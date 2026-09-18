//bearer token for updated API call
async function callAzureAI(displayPrompt, apiPrompt) {
  appendMessage("user", displayPrompt);
  conversationHistory.push({ role: "user", content: apiPrompt });

  const baseEndpoint = CONFIG.endpoint.replace(/\/+$/, "");
  const url = `${baseEndpoint}/openai/deployments/${CONFIG.deployment}/chat/completions?api-version=${CONFIG.apiVersion}`;
  console.log("Request URL:", url);
  //const url = `${CONFIG.endpoint}/openai/deployments/${CONFIG.deployment}/chat/completions?api-version=${CONFIG.apiVersion}`;

  try {
    //before calling entra ID, get the access token from openai
    const accessToken = await getAccessToken();
    const requestPayload = {
      messages: conversationHistory,
      response_format: { type: "json_object" },
      reasoning_effort: "low",            
      max_completion_tokens: 10000         
    };
    //send the request with the token in the Authorization header
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${accessToken}`
      },
      body: JSON.stringify(requestPayload)
    });

    if (!response.ok) {
      if (response.status === 401 || response.status === 403) {
        throw new Error("שגיאת הרשאות: ודא שלמשתמש יש תפקיד 'Cognitive Services OpenAI User' במשאב ה-Azure.");
      }
      if (response.status === 429) {
        const retryAfter = response.headers.get("Retry-After") || "כמה";
        throw new Error(`חריגה ממכסת בקשות (Rate Limit 429). נסה שוב בעוד ${retryAfter} שניות.`);
      }
      const errJson = await response.json().catch(() => null);
      const detailedMessage = errJson?.error?.message || (await response.text());
      throw new Error(`שגיאת API (${response.status}): ${detailedMessage}`);
    }

    const data = await response.json();
    let replyString = data.choices[0].message.content;
    
    conversationHistory.push({ role: "assistant", content: replyString });
    replyString = replyString.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
    const aiResponse = JSON.parse(replyString);

    appendMessage("assistant", aiResponse.chat_message, aiResponse.action);
    // Execute all requested actions sequentially
    if (Array.isArray(aiResponse.actions)) {
      for (const act of aiResponse.actions) {
        if (act.type === "format_text" && act.target_text) {
          await formatDocumentSubstring(act.target_text, act.format);}
        else if(act.type === "replace_all" && act.action_text) {
          await replaceEntireDocument(act.action_text);}
         else if (act.type === "replace_text" && act.target_text) {
          await replaceDocumentSubstring(act.target_text, act.action_text || "");
        } else if (act.type === "insert_end" && act.action_text) {
          await insertTextAtEnd(act.action_text);
        }
      }
    }

  } catch (error) {
    appendMessage("assistant", `שגיאה בתקשורת או בפענוח: ${error.message}`);
    console.error("Fetch error:", error);
    appendMessage(
      "assistant", 
      `שגיאה: ${error.message}\n\nכתובת היעד:\n${url}`
    );
  }
  
}
