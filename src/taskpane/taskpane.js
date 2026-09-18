
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
      fileInput.addEventListener("change", window.handleFileUpload);
    }
    if (removeFileBtn) {
      removeFileBtn.addEventListener("click", clearAttachedFiles);
    }
  }
});
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

