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

