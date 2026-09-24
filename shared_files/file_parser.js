console.log("🔥 UPLOAD fileparser.js LOADED SUCCESSFULLY!");

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

window.parseSingleFile = function(file) {
  return new Promise((resolve, reject) => {
    const fileExtension = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();

    reader.onerror = () => reject(new Error(`קריאת הקובץ ${file.name} נכשלה`));

    // 1. קובצי Word
    if (fileExtension === "docx") {
      reader.onload = async (e) => {
        try {
          const text = await window.extractTextFromDocx(e.target.result);
          resolve(`--- תוכן קובץ Word: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);

    // 2. קובצי PDF
    } else if (fileExtension === "pdf") {
      reader.onload = async (e) => {
        try {
          const text = await window.extractTextFromPdf(e.target.result);
          resolve(`--- תוכן קובץ PDF: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);

    // 3. קובצי Excel
    } else if (fileExtension === "xlsx" || fileExtension === "xls") {
      reader.onload = async (e) => {
        try {
          const text = await window.extractTextFromExcel(e.target.result);
          resolve(`--- תוכן קובץ Excel: ${file.name} ---\n${text.trim()}\n`);
        } catch (err) {
          reject(new Error(`שגיאה בפענוח קובץ אקסל ${file.name}: ${err.message}`));
        }
      };
      reader.readAsArrayBuffer(file);

    } else if (["jpg", "jpeg", "png", "bmp", "webp"].includes(fileExtension)) {
      reader.onload = async (e) => {
        try {
          const text = await window.extractTextFromImage(e.target.result);
          resolve(`--- תוכן תמונה: ${file.name} (OCR) ---\n${text.trim() || "[לא זוהה טקסט בתמונה]"}\n`);
        } catch (err) {
          reject(new Error(`שגיאה בפענוח תמונה ${file.name}: ${err.message}`));
        }
      };
      reader.readAsDataURL(file);

    } else if (fileExtension === "doc") {
      reject(new Error(`הקובץ "${file.name}" הוא בפורמט .doc ישן. יש לשמור אותו כ-docx ולהעלות שוב.`));
      return;

    // 6. קובצי טקסט פשוטים (txt, csv, md)
    } else {
      reader.onload = (e) => {
        resolve(`--- תוכן קובץ טקסט: ${file.name} ---\n${e.target.result.trim()}\n`);
      };
      reader.readAsText(file, "UTF-8");
    }
  });
};

// --- פונקציית חילוץ טקסט מתמונה באמצעות Tesseract.js ---
window.extractTextFromImage = async function(imageSource) {
  if (!window.Tesseract) {
    throw new Error("ספריית Tesseract.js לא נטענה בדף");
  }

  const result = await window.Tesseract.recognize(
    imageSource,
    'heb+eng',
    {
      logger: m => {
        if (m.status === 'recognizing text') {
          console.log(`[OCR Progress]: ${Math.round((m.progress || 0) * 100)}%`);
        }
      }
    }
  );

  return result.data.text.trim();
};

window.extractTextFromDocx = async function(arrayBuffer) {
  if (!window.mammoth) {
    throw new Error("ספריית mammoth לא נטענה בדף");
  }
  const result = await window.mammoth.extractRawText({ arrayBuffer: arrayBuffer });
  return result.value.trim();
};

window.extractTextFromPdf = async function(arrayBuffer) {
  if (!window.pdfjsLib) {
    throw new Error("ספריית pdfjsLib לא נטענה בדף");
  }

  if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  const loadingTask = window.pdfjsLib.getDocument({ data: arrayBuffer });
  const pdfDoc = await loadingTask.promise;
  let fullText = "";

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ").trim();

    // בדיקה אם קיים טקסט דיגיטלי בעמוד
    if (pageText.length > 30) {
      fullText += `\n--- עמוד ${pageNum} ---\n` + pageText;
    } else {
      // עמוד סרוק ב-PDF: רינדור ל-Canvas והעברה ל-OCR מקומי
      const viewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      await page.render({ canvasContext: ctx, viewport: viewport }).promise;

      try {
        const ocrText = await window.extractTextFromImage(canvas);
        fullText += `\n--- עמוד סרוק ${pageNum} (OCR) ---\n` + (ocrText || "[לא זוהה טקסט קריא בעמוד זה]");
      } catch (ocrErr) {
        console.error(`שגיאה ב-OCR עמוד ${pageNum}:`, ocrErr);
        fullText += `\n--- עמוד סרוק ${pageNum} ---\n[שגיאה בחילוץ טקסט מתמונה: ${ocrErr.message}]`;
      }
    }
  }

  return fullText.trim();
};

window.extractTextFromExcel = async function(arrayBuffer) {
  if (!window.XLSX) {
    throw new Error("ספריית XLSX לא נטענה בדף");
  }
  const data = new Uint8Array(arrayBuffer);
  const workbook = window.XLSX.read(data, { type: "array" });
  let excelText = "";

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const csvContent = window.XLSX.utils.sheet_to_csv(worksheet);
    if (csvContent.trim()) {
      excelText += `[גיליון: ${sheetName}]\n${csvContent}\n\n`;
    }
  });

  return excelText.trim();
}