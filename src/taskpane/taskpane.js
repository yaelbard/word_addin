console.log("🔥  taskpane.js LOADED SUCCESSFULLY!");
// Initialize Office Add-in once the host application is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sendBtn = document.getElementById("send-btn");
    const inputArea = document.getElementById("prompt-input");
    const fileInput = document.getElementById("doc-upload");
    const removeFileBtn = document.getElementById("remove-file-btn");

    // Handle sending via button click
    if (sendBtn) {
      sendBtn.addEventListener("click", window.handleSend);
    }

    // Handle sending via Enter key (Shift + Enter for newline)
    if (inputArea) {
      inputArea.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          if (typeof window.handleSend === "function") {
            window.handleSend();
          }
        }
      });
    }

    // Handle file upload and removal events
    if (fileInput) {
      fileInput.addEventListener("change", (event) => {
        if (typeof window.handleFileUpload === "function") {
          window.handleFileUpload(event);
        }
      });
    }

    if (removeFileBtn) {
      removeFileBtn.addEventListener("click", () => {
        if (typeof window.clearAttachedFiles === "function") {
          window.clearAttachedFiles();
        }
      });
    }
  }
});

// Process user input, extract document/file context, and invoke Azure AI
window.handleSend = async function() {
  const input = document.getElementById("prompt-input");
  const sendBtn = document.getElementById("send-btn");
  const loader = document.getElementById("loadingIndicator");

  if (!input) return;

  const userText = input.value.trim();
  const files = Array.isArray(window.uploadedFiles) ? window.uploadedFiles : [];

  // Prevent sending if both input text and uploaded files are empty
  if (!userText && files.length === 0) return;

  input.value = "";
  if (loader) loader.style.display = "block";
  if (sendBtn) sendBtn.disabled = true;

  try {
    let docBodyText = "";

    // Read full active document text via Office Word API
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      docBodyText = body.text ? body.text.trim() : "";
    });

    let fullPrompt = userText;

    // Append current Word document content if present
    if (docBodyText.length > 0) {
      fullPrompt = `${fullPrompt}\n\n[Open Word Document Content]:\n"${docBodyText}"`;
    }

    // Append uploaded attachment contents if present
    if (files.length > 0) {
      const fileNames = files.map((item) => item.file.name).join(", ");
      const filesCombinedText = files.map((item) => item.text).join("\n\n");
      fullPrompt = `${fullPrompt}\n\n[Uploaded Files Content (${fileNames})]:\n"${filesCombinedText.trim()}"`;
    }

    // Call Azure AI completions endpoint via window
    if (typeof window.callAzureAI === "function") {
      await window.callAzureAI(userText, fullPrompt);
    } else {
      console.error("callAzureAI is not defined on window.");
    }

    // Reset attached files state and UI after successful request
    if (typeof window.clearAttachedFiles === "function") {
      window.clearAttachedFiles();
    }

  } catch (error) {
    console.error("Error during processing:", error);
    // Restore user input on failure
    input.value = userText;
  } finally {
    if (loader) loader.style.display = "none";
    if (sendBtn) sendBtn.disabled = false;
  }
};