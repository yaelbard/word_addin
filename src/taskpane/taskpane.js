/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


const CONFIG = {
  deployment: "gpt-5.6-sol",
  endpoint: "https://gpt-agents-1972-foundry.openai.azure.com",
  apiVersion: "2025-01-01-preview",
  clientId: "2603315b-9e9f-4b43-b2fd-3de9ff9c41bd", 
  tenantId: "73ea3442-65e1-4556-a609-904f5d2e45ab"
};

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const redirectPageUri = new URL("auth-redirect.html", window.location.href).href;

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: redirectPageUri,
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let isMsalInitialized = false;

async function initMsal() {
  if (!isMsalInitialized) {
    await msalInstance.initialize();
    isMsalInitialized = true;
  }
}

/**
 * פתיחת חלון אימות דרך Office Dialog API
 */
function openLoginDialog() {
  return new Promise((resolve, reject) => {
    // שמירת המזהים ישירות ב-localStorage המשותף
    if (!CONFIG || !CONFIG.clientId || !CONFIG.tenantId) {
      reject(new Error("CONFIG.clientId or CONFIG.tenantId is missing in taskpane.js"));
      return;
    }
    localStorage.setItem("entra_clientId", CONFIG.clientId);
    localStorage.setItem("entra_tenantId", CONFIG.tenantId);

    // נתיב הדיאלוג נשאר נקי וקבוע
    const dialogUrl = new URL("auth-redirect.html", window.location.href).href;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 70, width: 60, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Failed to open dialog: ${asyncResult.error.message}`));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const response = JSON.parse(arg.message);
            if (response.status === "success") {
              resolve(response.token);
            } else {
              reject(new Error(response.error || "Authentication failed"));
            }
          } catch (e) {
            reject(new Error("Malformed dialog response"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          if (arg.error === 12006) {
            reject(new Error("Login window was closed by the user."));
          }
        });
      }
    );
  });
}

/**
 * פונקציית האימות הראשית
 */
async function getAccessToken() {
  await initMsal();

  const loginRequest = {
    scopes: ["https://cognitiveservices.azure.com/.default"]
  };

  try {
    const accounts = msalInstance.getAllAccounts();

    // אם קיים חשבון שמור, מנסים להוציא טוקן באופן שקט ללא שום חלון
    if (accounts.length > 0) {
      const silentResponse = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0]
      });
      return silentResponse.accessToken;
    }
  } catch (silentError) {
    console.warn("Silent token acquisition failed. Opening dialog...", silentError);
  }

  // אם אין חשבון או שהטוקן השקט נכשל – פותחים דיאלוג ייעודי של אופיס
  return await openLoginDialog();
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
4. RTL & right-aligned Hebrew phrasing.
5. Line breaks: Never use multiple empty lines; at most a single newline.
`;


let conversationHistory = [
  { role: "system", content: SYSTEM_PROMPT }
];

// Variables to store the uploaded document data
let uploadedFiles = [];
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
      removeFileBtn.addEventListener("click", clearAttachedFiles);
    }
  }
});

async function handleFileUpload(event) {
  const newFiles = Array.from(event.target.files);
  if (!newFiles || newFiles.length === 0) return;

  const uniqueFiles = newFiles.filter(
    (newF) => !uploadedFiles.some((item) => item.file.name === newF.name && item.file.size === newF.size)
  );

  if (uniqueFiles.length === 0) {
    event.target.value = "";
    return;
  }

  try {
    const parsedItems = await Promise.all(
      uniqueFiles.map(async (file) => {
        const text = await parseSingleFile(file);
        return {
          id: `${file.name}-${file.size}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
          file: file,
          text: text
        };
      })
    );

    uploadedFiles = uploadedFiles.concat(parsedItems);
    renderAttachedFilesUI();

  } catch (err) {
    console.error("שגיאה בפענוח הקבצים:", err);
    appendMessage("assistant", `שגיאה בפענוח הקבצים: ${err.message}`);
  } finally {
    event.target.value = "";
  }
}

function removeSingleFile(fileId) {
  uploadedFiles = uploadedFiles.filter((item) => item.id !== fileId);
  renderAttachedFilesUI();
}

function clearAttachedFiles() {
  uploadedFiles = [];
  renderAttachedFilesUI();
  const fileInput = document.getElementById("doc-upload");
  if (fileInput) fileInput.value = "";
}

function renderAttachedFilesUI() {
  const container = document.getElementById("attached-files-container");
  if (!container) return;

  container.innerHTML = "";

  uploadedFiles.forEach((item) => {
    const card = document.createElement("div");
    card.className = "attached-file-chip";

    const infoDiv = document.createElement("div");
    infoDiv.className = "file-chip-info";

    const icon = document.createElement("span");
    icon.className = "file-icon";
    icon.textContent = "📄";

    const nameSpan = document.createElement("span");
    nameSpan.className = "file-chip-name";
    nameSpan.textContent = item.file.name;
    nameSpan.title = item.file.name;

    const sizeSpan = document.createElement("span");
    sizeSpan.className = "file-chip-size";
    sizeSpan.textContent = `(${formatFileSize(item.file.size)})`;

    infoDiv.appendChild(icon);
    infoDiv.appendChild(nameSpan);
    infoDiv.appendChild(sizeSpan);

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "remove-chip-btn";
    deleteBtn.title = "הסר קובץ";
    deleteBtn.textContent = "✕";
    deleteBtn.onclick = () => removeSingleFile(item.id);

    card.appendChild(infoDiv);
    card.appendChild(deleteBtn);
    container.appendChild(card);
  });
}

function parseSingleFile(file) {
  return new Promise((resolve, reject) => {
    const fileExtension = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();

    reader.onerror = () => reject(new Error(`קריאת הקובץ ${file.name} נכשלה`));

    if (fileExtension === "docx") {
      reader.onload = async (e) => {
        try {
          const text = await extractTextFromDocx(e.target.result);
          resolve(`--- תוכן קובץ Word: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);

    } else if (fileExtension === "pdf") {
      reader.onload = async (e) => {
        try {
          const text = await extractTextFromPdf(e.target.result);
          resolve(`--- תוכן קובץ PDF: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);

    } else if (fileExtension === "xlsx" || fileExtension === "xls") {
      reader.onload = async (e) => {
        try {
          const text = await extractTextFromExcel(e.target.result);
          resolve(`--- תוכן קובץ Excel: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(new Error(`שגיאה בפענוח קובץ אקסל ${file.name}: ${err.message}`));
        }
      };
      reader.readAsArrayBuffer(file);

    } else {
      // קובצי טקסט רגילים (txt, csv, md וכו')
      reader.onload = (e) => {
        resolve(`--- תוכן קובץ טקסט: ${file.name} ---\n${e.target.result.trim()}\n`);
      };
      reader.readAsText(file, "UTF-8");
    }
  });
}

async function extractTextFromDocx(arrayBuffer) {
  const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
  return result.value.trim();
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

async function extractTextFromExcel(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });
  let excelText = "";

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    if (csvContent.trim()) {
      excelText += `[גיליון: ${sheetName}]\n${csvContent}\n\n`;
    }
  });

  return excelText.trim();
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
async function handleSend() {
  const input = document.getElementById("prompt-input");
  const sendBtn = document.getElementById("send-btn");
  const loader = document.getElementById("loadingIndicator");

  const userText = input.value.trim();
  if (!userText && uploadedFiles.length === 0) return;
  input.value = "";
  if (loader) loader.style.display = "block";
  if (sendBtn) sendBtn.disabled = true;

  try {
    let docBodyText = ""; 
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      docBodyText = body.text ? body.text.trim() : "";
    });
    let fullPrompt = userText;
    if (docBodyText.length > 0) {
      fullPrompt = `${fullPrompt}\n\n[Open Word Document Content]:\n"${docBodyText}"`;
    }

    if (uploadedFiles.length > 0) {
    const fileNames = uploadedFiles.map((item) => item.file.name).join(", ");
    const filesCombinedText = uploadedFiles.map((item) => item.text).join("\n\n");
    fullPrompt = `${fullPrompt}\n\n[Uploaded Files Content (${fileNames})]:\n"${filesCombinedText.trim()}"`;
    }
    await callAzureAI(userText, fullPrompt);
    clearAttachedFiles();

  } catch (error) {
    console.error("Error during processing:", error);
    input.value = userText;
  } finally {
    if (loader) loader.style.display = "none";
    if (sendBtn) sendBtn.disabled = false;
  }
}

