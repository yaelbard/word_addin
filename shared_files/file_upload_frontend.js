
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
