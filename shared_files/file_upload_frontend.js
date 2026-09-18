// Variables to store the uploaded document data on window
console.log("🔥 UPLOAD FILE SCRIPT LOADED SUCCESSFULLY!");
window.uploadedFiles = [];  
window.uploadedFileText = "";

window.handleFileUpload = async function(event) {
  const newFiles = Array.from(event.target.files);
  if (!newFiles || newFiles.length === 0) return;

  const uniqueFiles = newFiles.filter(
    (newF) => !window.uploadedFiles.some((item) => item.file.name === newF.name && item.file.size === newF.size)
  );

  if (uniqueFiles.length === 0) {
    event.target.value = "";
    return;
  }

  try {
    const parsedItems = await Promise.all(
      uniqueFiles.map(async (file) => {
        // קריאה לפונקציית הפענוח דרך window למקרה שהוגדרה בקובץ אחר
        const text = await (window.parseSingleFile ? window.parseSingleFile(file) : Promise.resolve(""));
        return {
          id: `${file.name}-${file.size}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
          file: file,
          text: text
        };
      })
    );

    window.uploadedFiles = window.uploadedFiles.concat(parsedItems);
    window.renderAttachedFilesUI();

  } catch (err) {
    console.error("שגיאה בפענוח הקבצים:", err);
    if (typeof window.appendMessage === "function") {
      window.appendMessage("assistant", `שגיאה בפענוח הקבצים: ${err.message}`);
    }
  } finally {
    event.target.value = "";
  }
};

window.removeSingleFile = function(fileId) {
  window.uploadedFiles = window.uploadedFiles.filter((item) => item.id !== fileId);
  window.renderAttachedFilesUI();
};

window.clearAttachedFiles = function() {
  window.uploadedFiles = [];
  window.renderAttachedFilesUI();
  const fileInput = document.getElementById("doc-upload");
  if (fileInput) fileInput.value = "";
};

window.renderAttachedFilesUI = function() {
  const container = document.getElementById("attached-files-container");
  if (!container) return;

  container.innerHTML = "";

  window.uploadedFiles.forEach((item) => {
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
    // שימוש בפונקציית העיצוב דרך window עם ברירת מחדל בטוחה
    const formattedSize = typeof window.formatFileSize === "function" 
      ? window.formatFileSize(item.file.size) 
      : `${item.file.size} B`;
    sizeSpan.textContent = `(${formattedSize})`;

    infoDiv.appendChild(icon);
    infoDiv.appendChild(nameSpan);
    infoDiv.appendChild(sizeSpan);

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "remove-chip-btn";
    deleteBtn.title = "הסר קובץ";
    deleteBtn.textContent = "✕";
    deleteBtn.onclick = () => window.removeSingleFile(item.id);

    card.appendChild(infoDiv);
    card.appendChild(deleteBtn);
    container.appendChild(card);
  });
};