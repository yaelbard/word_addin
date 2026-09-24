console.log("🔥 UPLOAD fileparser.js LOADED SUCCESSFULLY!");

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

window.parseSingleFile = function(file) {
  return new Promise((resolve, reject) => {
    const fileExtension = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(`קריאת הקובץ ${file.name} נכשלה`));
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
    } else {
      reader.onload = (e) => {
        resolve(`--- תוכן קובץ טקסט: ${file.name} ---\n${e.target.result.trim()}\n`);
      };
      reader.readAsText(file, "UTF-8");
    }
  });
};

window.extractTextFromImage = async function(imageSource) {
  if (!window.Tesseract) {
    throw new Error("ספריית Tesseract.js לא נטענה בדף");
  }
  const result = await window.Tesseract.recognize(imageSource, "heb+eng", {
    logger: (m) => {
      if (m.status === "recognizing text") {
        console.log(`[OCR Progress]: ${Math.round((m.progress || 0) * 100)}%`);
      }
    }
  });
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
  if (!window.Tesseract) {
    throw new Error("ספריית Tesseract.js לא נטענה בדף");
  }

  const loadingTask = window.pdfjsLib.getDocument({ data: arrayBuffer });
  const pdfDoc = await loadingTask.promise;
  let fullText = "";
  const lengthThreshold = 30;
  let ocrWorker = null;

  try {
    console.log("🔧 יוצר Tesseract Worker עבור ה-PDF...");
    ocrWorker = await window.Tesseract.createWorker("heb+eng", 1, {
      logger: (m) => {
        if (m.status === "loading language traineddata") {
          console.log(`[OCR] טוען שפות: ${Math.round((m.progress || 0) * 100)}%`);
        }
        if (m.status === "initializing api") {
          console.log(`[OCR] מאתחל מנוע: ${Math.round((m.progress || 0) * 100)}%`);
        }
        if (m.status === "recognizing text") {
          console.log(`[OCR] עיבוד: ${Math.round((m.progress || 0) * 100)}%`);
        }
      }
    });
    console.log("✅ Tesseract Worker מוכן לשימוש");

    for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
      const page = await pdfDoc.getPage(pageNum);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map((item) => item.str).join(" ").trim();
      const meaningfulCharCount = pageText.replace(/\s/g, "").length;

      if (meaningfulCharCount > lengthThreshold) {
        console.log(`📄 עמוד ${pageNum}: נמצאו ${meaningfulCharCount} תווים משמעותיים, משתמש בטקסט הדיגיטלי.`);
        fullText += `\n--- עמוד ${pageNum} ---\n` + pageText;
      } else {
        console.log(`📸 עמוד ${pageNum}: נמצאו רק ${meaningfulCharCount} תווים משמעותיים, מפעיל OCR...`);
        const viewport = page.getViewport({ scale: 2.0 });
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);

        try {
          await page.render({ canvasContext: ctx, viewport: viewport }).promise;
          const dataUrl = canvas.toDataURL("image/png");
          console.log(`🔍 OCR עמוד ${pageNum} באמצעות Worker משותף...`);
          const result = await ocrWorker.recognize(dataUrl);
          const ocrText = result.data && result.data.text ? result.data.text.trim() : "";
          fullText += `\n--- עמוד סרוק ${pageNum} (OCR) ---\n` + (ocrText || "[לא זוהה טקסט קריא בעמוד זה]");
        } catch (ocrErr) {
          console.error(`שגיאה ב-OCR עמוד ${pageNum}:`, ocrErr);
          fullText += `\n--- עמוד סרוק ${pageNum} ---\n` + `[שגיאה בחילוץ טקסט מתמונה: ${ocrErr.message || ocrErr}]`;
        } finally {
          canvas.width = 0;
          canvas.height = 0;
        }
      }
    }
  } finally {
    if (ocrWorker) {
      console.log("🧹 סוגר Tesseract Worker...");
      try {
        await ocrWorker.terminate();
        console.log("✅ Tesseract Worker נסגר בהצלחה");
      } catch (terminateErr) {
        console.warn("⚠️ שגיאה בסגירת Tesseract Worker:", terminateErr);
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
      excelText += `[גיליון: ${sheetName}]\n` + `${csvContent}\n\n`;
    }
  });

  return excelText.trim();
};