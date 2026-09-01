/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, msal */

// 1. הגדרות תצורה: אין יותר API Key!
const CONFIG = {
  endpoint: "https://gpt-agents-1972-foundry.services.ai.azure.com/",
  deployment: "gpt-4o",
  apiVersion: "2024-08-01-preview",
  
  // פרטי ה-Entra ID שלך (ציבוריים, מותר להעלות ל-GitHub)
  clientId: "2603315b-9e9f-4b43-b2fd-3de9ff9c41bd", 
  tenantId: "73ea3442-65e1-4556-a609-904f5d2e45ab", 
  redirectUri: "https://yaelbard.github.io/word_addin/taskpane.htm"
};

// 2. הגדרת ספריית MSAL לאימות מול Entra ID
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// 3. פונקציה חכמה לקבלת טוקן גישה (מבקשת התחברות במידת הצורך)
async function getAccessToken() {
  // ה-Scope הנדרש כדי לגשת ישירות ל-Azure OpenAI באמצעות Entra ID
  const loginRequest = {
    scopes: ["https://cognitiveservices.azure.com/.default"] 
  };

  try {
    const accounts = msalInstance.getAllAccounts();
    
    // אם המשתמש לא מחובר כלל, פתח חלון התחברות
    if (accounts.length === 0) {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      return loginResponse.accessToken;
    }

    // ניסיון קבלת טוקן שקט (ללא חלון) למשתמש שכבר מחובר
    const silentResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    });
    return silentResponse.accessToken;

  } catch (error) {
    console.warn("Silent token acquisition failed. Acquiring token using popup...", error);
    // אם הטוקן השקט נכשל (למשל פג תוקף), פתח חלון לאישור מחדש
    const popupResponse = await msalInstance.acquireTokenPopup(loginRequest);
    return popupResponse.accessToken;
  }
}

// הנחיית מערכת
const SYSTEM_PROMPT = `אתה עוזר AI מקצועי, חכם, שירותי ורהוט המשולב בתוך Microsoft Word.
סגנון התשובה שלך (בשדה chat_message) צריך להיות מפורט, אינטליגנטי, מנומס וברמה גבוהה מאוד, בדיוק כמו בשיחה טבעית וזורמת. ענה תמיד בעברית מעולה.

חובה עליך להחזיר תמיד אך ורק אובייקט JSON תקין.
מבנה ה-JSON הנדרש:
{
  "chat_message": "כאן תכתוב את התשובה המלאה, העשירה והשירותית שלך למשתמש. אם התבקשת לנתח, להסביר או לענות על שאלה - עשה זאת כאן בהרחבה.",
  "action": "סוג הפעולה במסמך. אפשרויות: 'none', 'replace_selection', 'insert_end'",
  "action_text": "הטקסט לביצוע הפעולה במסמך. אם action הוא 'none', השאר שדה זה ריק ("")."
}

חוקי עריכה במסמך (חשוב מאוד!):
1. שיח שאלות ותשובות: אם המשתמש רק שואל שאלה, מתייעץ או אומר שלום, הגדר action כ-"none" ותן מענה עשיר ואיכותי ב-chat_message.
2. עריכה, שכתוב או מחיקה חלקית (החלפה): אם המשתמש מבקש לשנות את הטקסט המסומן (למשל לשכתב אותו או **למחוק מתוכו שורות ומילים**), עליך תמיד להגדיר את action כ-'replace_selection'. בשדה action_text החזר **אך ורק את הגרסה הסופית והמתוקנת** (ללא השורות שנמחקו). לעולם אל תחזיר את הטקסט המקורי יחד עם החדש, ואסור בשום אופן להשתמש ב-'insert_end' לעריכת טקסט קיים!
3. מחיקת טקסט מתוך בחירה: אם המשתמש מבקש למחוק שורות או מילים מסוימות מתוך הקטע המסומן, הגדר action כ-"replace_selection". בשדה action_text, החזר את הטקסט *שנותר* לאחר המחיקה. 
4. מחיקה מלאה: אם התבקשת למחוק את כל הטקסט המסומן לחלוטין, action יהיה 'replace_selection' ו-action_text יהיה מחרוזת ריקה ("").
5. הוספת טקסט חדש: רק אם המשתמש מבקש לכתוב תוכן חדש לחלוטין שלא מסתמך על הקיים, השתמש ב-'insert_end'.
6. חוסר סימון: אם המשתמש מבקש לערוך משהו ספציפי אבל לא מופיע לך "טקסט מסומן לעריכה" בבקשה, אל תנחש. הגדר action כ-"none" ובקש ממנו בנימוס ב-chat_message לסמן את הקטע הרלוונטי בעכבר.
`;

let conversationHistory = [
  { role: "system", content: SYSTEM_PROMPT }
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {    
    const sendBtn = document.getElementById("send-btn");
    const inputArea = document.getElementById("prompt-input");

    sendBtn.onclick = handleSend;

    inputArea.addEventListener("keydown", (event) => {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        handleSend();
      }
    });
  }
});

async function replaceSelectedText(newText) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    if (!newText || newText.trim() === "") {
      selection.delete(); // מפעיל מחיקה אמיתית של הקטע מהמסמך
    } else {
      selection.insertText(newText, Word.InsertLocation.replace); // דורס את הישן ושם את החדש
    }
    
    await context.sync();
  });
}

async function insertTextAtEnd(newText) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(newText, Word.InsertLocation.end);
    await context.sync();
  });
}

function appendMessage(role, text, actionType = "none") {
  const chatBox = document.getElementById("chat-box");
  const msgElement = document.createElement("div");
  msgElement.className = `msg ${role}`;
  msgElement.textContent = text;

  if (actionType !== "none") {
    const notice = document.createElement("div");
    notice.className = "system-notice";
    notice.textContent = "✅ בוצעה עריכה אוטומטית במסמך";
    msgElement.appendChild(notice);
  }

  chatBox.appendChild(msgElement);
  chatBox.scrollTop = chatBox.scrollHeight;
}

// עדכון פונקציית הקריאה ל-API לשימוש ב-Bearer Token
async function callAzureAI(displayPrompt, apiPrompt) {
  appendMessage("user", displayPrompt);
  conversationHistory.push({ role: "user", content: apiPrompt });

  const url = `${CONFIG.endpoint}/openai/deployments/${CONFIG.deployment}/chat/completions?api-version=${CONFIG.apiVersion}`;

  try {
    // 1. השגת הטוקן מ-Entra ID לפני קריאה ל-OpenAI
    const accessToken = await getAccessToken();

    // 2. שליחת הבקשה עם הטוקן בכותרת Authorization
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${accessToken}` // <--- השינוי המרכזי כאן
      },
      body: JSON.stringify({
        messages: conversationHistory,
        temperature: 0.7, 
        response_format: { type: "json_object" }
      })
    });

    if (!response.ok) {
      if (response.status === 401 || response.status === 403) {
        throw new Error("שגיאת הרשאות: ודא שלמשתמש יש תפקיד 'Cognitive Services OpenAI User' במשאב ה-Azure.");
      }
      throw new Error(`שגיאת API: ${response.status}`);
    }

    const data = await response.json();
    let replyString = data.choices[0].message.content;
    
    conversationHistory.push({ role: "assistant", content: replyString });

    if (replyString.startsWith("```json")) {
        replyString = replyString.replace(/^```json\n/, "").replace(/\n```$/, "");
    }

    const aiResponse = JSON.parse(replyString);

    appendMessage("assistant", aiResponse.chat_message, aiResponse.action);

    if (aiResponse.action === "replace_selection") {
      await replaceSelectedText(aiResponse.action_text);
    } else if (aiResponse.action === "insert_end" && aiResponse.action_text) {
      await insertTextAtEnd(aiResponse.action_text);
    }

  } catch (error) {
    appendMessage("assistant", `שגיאה בתקשורת או בפענוח: ${error.message}`);
  }
}

async function handleSend() {
  const input = document.getElementById("prompt-input");
  const userText = input.value.trim();
  if (!userText) return;

  input.value = "";

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    let fullPrompt = userText;

    if (selection.text && selection.text.trim().length > 0) {
      fullPrompt = `${userText}\n\n[טקסט מסומן לעריכה]:\n"${selection.text}"`;
    }

    callAzureAI(userText, fullPrompt);
  });
}
