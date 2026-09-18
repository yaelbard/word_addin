
// Format file size helper
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
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
