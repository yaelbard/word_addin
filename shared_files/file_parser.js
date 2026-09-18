if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
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
    } 
    else if (fileExtension === "doc") {
  reject(new Error(`הקובץ "${file.name}" הוא בפורמט .doc ישן. יש לשמור אותו כ-docx ולהעלות שוב.`));
  return;
    }
    else {
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
    const pageText = textContent.items.map((item) => item.str).join(" ").trim();

    // אם קיים טקסט דיגיטלי בעמוד
    if (pageText.length > 30) {
      fullText += `\n--- עמוד ${pageNum} ---\n` + pageText;
    } else {
      // עמוד סרוק / תמונה: יצירת תמונה באמצעות Canvas בפורמט Base64
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      await page.render({ canvasContext: ctx, viewport: viewport }).promise;
      const imageBase64 = canvas.toDataURL("image/jpeg", 0.85);

      fullText += `\n--- עמוד סרוק ${pageNum} (תמונה) ---\n[Image: ${imageBase64}]\n`;
    }
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
