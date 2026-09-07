/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


const CONFIG = {
  //endpoint: "https://gpt-agents-1972-foundry.services.ai.azure.com/api/projects/gpt-agents-1972-proj/openai/v1/responses",
  deployment: "gpt-5.6-sol",
  endpoint: "https://gpt-agents-1972-foundry.openai.azure.com",
  apiVersion: "2025-01-01-preview",
  clientId: "2603315b-9e9f-4b43-b2fd-3de9ff9c41bd", 
  tenantId: "73ea3442-65e1-4556-a609-904f5d2e45ab", 
  //redirectUri: "https://yaelbard.github.io/word_addin/src/taskpane/taskpane.html"
};
// defining the worker source for pdf.js to enable PDF text extraction
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const cleanUrl = window.location.origin + window.location.pathname;
const dynamicRedirectUri = cleanUrl.replace('taskpane.html', 'auth-redirect.html');

// MSAL configuration for Entra ID authentication
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: dynamicRedirectUri
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// smart function to get an access token (prompts for login if necessary)
async function getAccessToken() {
  // entra ID scopes required to directly access Azure OpenAI
  const loginRequest = {
    scopes: ["https://cognitiveservices.azure.com/.default"] 
  };

  try {
    const accounts = msalInstance.getAllAccounts();
    //if the user is not logged in at all, open a login window    
    if (accounts.length === 0) {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      return loginResponse.accessToken;
    }
    //try to get a silent token (no window) for the already logged-in user
    const silentResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    });
    return silentResponse.accessToken;

  } catch (error) {
    console.warn("Silent token acquisition failed. Acquiring token using popup...", error);
    //opens a popup for the user to re-authenticate if silent token acquisition fails (e.g., token expired)
    const popupResponse = await msalInstance.acquireTokenPopup(loginRequest);
    return popupResponse.accessToken;
  }
}

//system prompt for the AI assistant, defining its behavior and response format
const SYSTEM_PROMPT = `AI Word assistant. Analyze full doc context directly without relying on mouse selection. Use natural Hebrew.
Output STRICT JSON:
{
  "chat_message": "Hebrew reply",
  "actions": [
    {
      "type": "format_text|replace_text|insert_end",
      "target_text": "EXACT target text segment (under 200 chars, empty for insert_end)",
      "action_text": "new text or empty string",
      "format": {"bold": null, "underline": null, "italic": null, "font_name": null}
    }
  ]
}

Rules:
1. Pure styling (underline, bold, italic, font change): action="format_text", target_text="EXACT target word/phrase only", set requested fields in format object (e.g. {"underline": true}). Do not touch or rewrite other words.
2. Edit/Replace text content: action="replace_text", target_text="EXACT doc text segment (under 200 chars)", action_text="new plain text".
3. Chat/Q&A: action="none". Full rewrite: action="replace_all". Append: action="insert_end".
4. RTL & right-aligned Hebrew phrasing.`;

let conversationHistory = [
  { role: "system", content: SYSTEM_PROMPT }
];

// Variables to store the uploaded document data
let uploadedFile = null;
let uploadedFileText = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sendBtn = document.getElementById("send-btn");
    const inputArea = document.getElementById("prompt-input");
    const fileInput = document.getElementById("doc-upload");
    const removeFileBtn = document.getElementById("remove-file-btn");

    // Handle sending via button click
    sendBtn.addEventListener("click", handleSend);

    // Handle sending via Enter key (Shift + Enter for newline)
    inputArea.addEventListener("keydown", (event) => {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        handleSend();
      }
    });

    // Handle file upload and removal
    if (fileInput) {
      fileInput.addEventListener("change", handleFileUpload);
    }
    if (removeFileBtn) {
      removeFileBtn.addEventListener("click", clearAttachedFile);
    }
  }
});

// File upload handler
async function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  uploadedFile = file;

  // הצגת תגית הקובץ המצורף בממשק
  const card = document.getElementById("file-attached-card");
  const nameEl = document.getElementById("attached-file-name");
  const sizeEl = document.getElementById("attached-file-size");

  if (card && nameEl && sizeEl) {
    nameEl.innerText = file.name;
    sizeEl.innerText = formatFileSize(file.size);
    card.style.display = "flex";
  }

  const fileExtension = file.name.split(".").pop().toLowerCase();
  const reader = new FileReader();

  try {
    if (fileExtension === "docx") {
      //docx files (using mammoth.js)
      reader.onload = async (e) => {
        const arrayBuffer = e.target.result;
        const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
        uploadedFileText = result.value;
        console.log("חולץ טקסט מ-DOCX, אורך:", uploadedFileText.length);
      };
      reader.readAsArrayBuffer(file);

    } else if (fileExtension === "pdf") {
      //pdf files (using pdfjs-dist)
      reader.onload = async (e) => {
        const arrayBuffer = e.target.result;
        uploadedFileText = await extractTextFromPdf(arrayBuffer);
        console.log("חולץ טקסט מ-PDF, אורך:", uploadedFileText.length);
      };
      reader.readAsArrayBuffer(file);

    } else {
      //regular text files (txt, csv, etc.)
      reader.onload = (e) => {
        uploadedFileText = e.target.result;
      };
      reader.readAsText(file, "UTF-8");
    }
  } catch (err) {
    console.error("שגיאה בפענוח הקובץ:", err);
    appendMessage("assistant", `שגיאה בפענוח הקובץ: ${err.message}`);
  }
}

async function extractTextFromPdf(arrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdfDoc = await loadingTask.promise;
  let fullText = "";

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ");
    fullText += `\n--- עמוד ${pageNum} ---\n` + pageText;
  }

  return fullText.trim();
}
// Clear the attached file
function clearAttachedFile() {
  uploadedFile = null;
  uploadedFileText = "";
  const fileInput = document.getElementById("doc-upload");
  const card = document.getElementById("file-attached-card");

  if (fileInput) fileInput.value = "";
  if (card) card.style.display = "none";
}

// Format file size helper
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

async function applyContentToRange(context, targetRange, content, insertLocation) {
  let insertedRange;
  // Detect if content contains HTML tags
  const hasHtml = /<[a-z][\s\S]*>/i.test(content);

  if (hasHtml) {
    // Wrap with dir="rtl" to ensure proper bidirectional rendering
    const wrappedHtml = `<div dir="rtl" style="text-align: right;">${content}</div>`;
    insertedRange = targetRange.insertHtml(wrappedHtml, insertLocation);
  } else {
    insertedRange = targetRange.insertText(content, insertLocation);
  }

  insertedRange.paragraphs.load("items");
  await context.sync();

  insertedRange.paragraphs.items.forEach((p) => {
    p.alignment = Word.Alignment.right;
    try { p.isRightToLeft = true; } catch (e) {}
  });

  await context.sync();
}
async function replaceDocumentSubstring(targetText, replacementContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) {
      throw new Error("Target text is empty.");
    }

    // Direct search for segments under the 255-char limit
    if (cleanTarget.length <= 250) {
      const searchResults = body.search(cleanTarget, { matchCase: false });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > 0) {
        await applyContentToRange(context, searchResults.items[0], replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    // Paragraph scan fallback for longer segments
    const paragraphs = body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    let matchedPara = null;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const pText = paragraphs.items[i].text.trim();
      if (!pText) continue;

      if (pText.includes(cleanTarget) || cleanTarget.includes(pText)) {
        matchedPara = paragraphs.items[i];
        break;
      }
    }

    if (matchedPara) {
      await applyContentToRange(context, matchedPara, replacementContent, Word.InsertLocation.replace);
      return;
    }

    throw new Error("Target text not found in the document.");
  });
}

async function replaceEntireDocument(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await applyContentToRange(context, body, newContent, Word.InsertLocation.replace);
  });
}

async function insertTextAtEnd(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await applyContentToRange(context, body, "\n" + newContent, Word.InsertLocation.end);
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
async function formatDocumentSubstring(targetText, formatOptions) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) {
      throw new Error("Target text is empty.");
    }

    const searchResults = body.search(cleanTarget, { matchCase: false });
    searchResults.load("items");
    await context.sync();

    if (searchResults.items.length === 0) {
      throw new Error("Target text not found for formatting.");
    }

    // Apply styles strictly to the matched target range
    const targetRange = searchResults.items[0];

    if (formatOptions) {
      if (typeof formatOptions.bold === "boolean") {
        targetRange.font.bold = formatOptions.bold;
      }
      if (typeof formatOptions.underline === "boolean") {
        targetRange.font.underline = formatOptions.underline ? "Single" : "None";
      }
      if (typeof formatOptions.italic === "boolean") {
        targetRange.font.italic = formatOptions.italic;
      }
      if (formatOptions.font_name) {
        targetRange.font.name = formatOptions.font_name;
      }
    }

    await context.sync();
  });
}
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

    if (replyString.startsWith("```json")) {
        replyString = replyString.replace(/^```json\n/, "").replace(/\n```$/, "");
    }

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

async function handleSend() {
  const input = document.getElementById("prompt-input");
  const sendBtn = document.getElementById("send-btn");
  const loader = document.getElementById("loadingIndicator");

  const userText = input.value.trim();
  // Ensure we don't send empty requests unless a file is attached
  if (!userText && !uploadedFileText) return;

  input.value = "";

  // Show loading spinner and disable send button
  if (loader) loader.style.display = "block";
  if (sendBtn) sendBtn.disabled = true;

  try {
    await Word.run(async (context) => {
      // Ingest the entire Word document body
      const body = context.document.body;
      body.load("text");
      await context.sync();

      let fullPrompt = userText;

      // Append open Word document text
      if (body.text && body.text.trim().length > 0) {
        fullPrompt = `${fullPrompt}\n\n[Open Word Document Content]:\n"${body.text.trim()}"`;
      }

      // Append uploaded external file content
      if (uploadedFileText) {
        fullPrompt = `${fullPrompt}\n\n[Uploaded File Content (${uploadedFile.name})]:\n"${uploadedFileText.trim()}"`;
      }

      // Call the AI model
      await callAzureAI(userText, fullPrompt);

      // Reset attached file after successful send
      clearAttachedFile();
    });
  } catch (error) {
    console.error("Error during processing:", error);
  } finally {
    // Hide loading spinner and re-enable send button
    if (loader) loader.style.display = "none";
    if (sendBtn) sendBtn.disabled = false;
  }
}

async function sendPrompt() {
  const promptInput = document.getElementById("prompt-input");
  const sendBtn = document.getElementById("send-btn");
  const loader = document.getElementById("loadingIndicator");
  const responseBox = document.getElementById("chatResponse");

  const query = promptInput.value.trim();
  if (!query) return;

  // 1. הפעלת מצב טעינה והשבתת קלט
  loader.style.display = "block";
  sendBtn.disabled = true;
  responseBox.innerText = "";

  try {
    // קריאה לפונקציית ה-API שלך
    const reply = await callGptApi(query);
    responseBox.innerText = reply;
  } catch (err) {
    console.error(err);
    responseBox.innerText = "חלה שגיאה בקבלת המענה מהמודל.";
  } finally {
    // 2. כיבוי מצב הטעינה
    loader.style.display = "none";
    sendBtn.disabled = false;
  }
}

