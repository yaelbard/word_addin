console.log("🔥  api-client.js LOADED SUCCESSFULLY!");
// Initialize conversation history on the global window object
window.conversationHistory = [
  { 
    role: "system", 
    content: (typeof window.SYSTEM_PROMPT !== "undefined" ? window.SYSTEM_PROMPT : "") 
  }
];

// Bearer token for updated Azure AI API call
window.callAzureAI = async function(displayPrompt, apiPrompt) {
  // Ensure appendMessage is available before invoking
  if (typeof window.appendMessage === "function") {
    window.appendMessage("user", displayPrompt);
  }

  // Update system prompt if it was assigned after initial script execution
  if (
    window.conversationHistory.length > 0 && 
    !window.conversationHistory[0].content && 
    window.SYSTEM_PROMPT
  ) {
    window.conversationHistory[0].content = window.SYSTEM_PROMPT;
  }

  window.conversationHistory.push({ role: "user", content: apiPrompt });

  const config = window.CONFIG || {};
  if (!config.endpoint || !config.deployment || !config.apiVersion) {
    const errorMsg = "Configuration error: window.CONFIG is missing or incomplete.";
    console.error(errorMsg);
    if (typeof window.appendMessage === "function") {
      window.appendMessage("assistant", errorMsg);
    }
    return;
  }

  const baseEndpoint = config.endpoint.replace(/\/+$/, "");
  const url = `${baseEndpoint}/openai/deployments/${config.deployment}/chat/completions?api-version=${config.apiVersion}`;
  console.log("Request URL:", url);

  try {
    if (typeof window.getAccessToken !== "function") {
      throw new Error("Function getAccessToken is not defined on window.");
    }

    const accessToken = await window.getAccessToken();
    const requestPayload = {
      messages: window.conversationHistory,
      response_format: { type: "json_object" },
      reasoning_effort: "low",
      max_completion_tokens: 10000
    };

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
        throw new Error("Authorization error: Ensure the user has the 'Cognitive Services OpenAI User' role assigned on the Azure resource.");
      }
      if (response.status === 429) {
        const retryAfter = response.headers.get("Retry-After") || "a few";
        throw new Error(`Rate limit exceeded (Status 429). Please retry after ${retryAfter} seconds.`);
      }
      const errJson = await response.json().catch(() => null);
      const detailedMessage = errJson?.error?.message || (await response.text());
      throw new Error(`API Error (${response.status}): ${detailedMessage}`);
    }

    const data = await response.json();
    let replyString = data.choices[0].message.content;

    window.conversationHistory.push({ role: "assistant", content: replyString });
    replyString = replyString.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
    const aiResponse = JSON.parse(replyString);

    if (typeof window.appendMessage === "function") {
      window.appendMessage("assistant", aiResponse.chat_message, aiResponse.action);
    }

    // Execute all requested Word document actions sequentially
    if (Array.isArray(aiResponse.actions)) {
      for (const act of aiResponse.actions) {
        if (act.type === "format_text" && act.target_text && typeof window.formatDocumentSubstring === "function") {
          await window.formatDocumentSubstring(act.target_text, act.format);
        } else if (act.type === "replace_all" && act.action_text && typeof window.replaceEntireDocument === "function") {
          await window.replaceEntireDocument(act.action_text);
        } else if (act.type === "replace_text" && act.target_text && typeof window.replaceDocumentSubstring === "function") {
          await window.replaceDocumentSubstring(act.target_text, act.action_text || "");
        } else if (act.type === "insert_end" && act.action_text && typeof window.insertTextAtEnd === "function") {
          await window.insertTextAtEnd(act.action_text);
        }
      }
    }

  } catch (error) {
    console.error("Fetch error:", error);
    if (typeof window.appendMessage === "function") {
      window.appendMessage("assistant", `Communication/parsing error: ${error.message}\n\nEndpoint URL:\n${url}`);
    }
  }
};