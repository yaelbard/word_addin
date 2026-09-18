console.log("🔥 UPLOAD frontend_utils.js LOADED SUCCESSFULLY!");
// Format file size helper
window.formatFileSize = function(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
};

window.appendMessage = function(role, text, actionType = "none") {
  const chatBox = document.getElementById("chat-box");
  if (!chatBox) {
    console.warn("appendMessage: אלמנט chat-box לא נמצא ב-DOM");
    return;
  }

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
};