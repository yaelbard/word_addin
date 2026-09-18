
function normalizeText(str) {
  if (!str) return "";
  return str
    .replace(/[\u200B-\u200D\uFEFF\u200E\u200F]/g, "") // הסרת תווי כיווניות נסתרים
    .replace(/[\u201C\u201D\u05F4"]/g, '"')             // איחוד כל סוגי הגרשיים
    .replace(/[\u2018\u2019\u05F3']/g, "'")             // איחוד גרש בודד
    .replace(/[\r\n\t]+/g, " ")                          // הפיכת ירידות שורה לרווח
    .replace(/\s+/g, " ")                               // איחוד רווחים כפולים
    .trim();
}
async function replaceDocumentSubstring(targetText, replacementContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) throw new Error("Target text is empty.");

    // שלב 1: חיפוש רגיל של Word (מהיר)
    if (cleanTarget.length <= 250) {
      const searchResults = body.search(cleanTarget, { 
        matchCase: false, 
        matchWholeWord: false 
      });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > 0) {
        await applyContentToRange(context, searchResults.items[0], replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    // שלב 2: סריקה לפי פסקאות מנורמלות (מתגבר על הבדלי מקלדת/גרשיים/רווחים)
    const paragraphs = body.paragraphs;
    paragraphs.load(["text", "items"]);
    await context.sync();

    const normalizedTarget = normalizeText(cleanTarget);

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const normalizedPText = normalizeText(p.text);

      if (!normalizedPText) continue;

      if (normalizedPText.includes(normalizedTarget) || normalizedTarget.includes(normalizedPText)) {
        await applyContentToRange(context, p, replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    // שלב 3: חיפוש תת-מחרוזת ראשונית (אם הסוף נחתך או שונה מעט)
    if (cleanTarget.length > 25) {
      const shortTarget = cleanTarget.substring(0, 25).trim();
      const fallbackSearch = body.search(shortTarget, { matchCase: false });
      fallbackSearch.load("items");
      await context.sync();

      if (fallbackSearch.items.length > 0) {
        await applyContentToRange(context, fallbackSearch.items[0], replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    throw new Error(`Target text not found: "${cleanTarget.substring(0, 35)}..."`);
  });
}
async function replaceEntireDocument(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await applyContentToRange(context, body, newContent, Word.InsertLocation.replace);
  });
}

async function insertTextAtEnd(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await applyContentToRange(context, body, "\n" + newContent, Word.InsertLocation.end);
  });
}
async function formatDocumentSubstring(targetText, formatOptions) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) throw new Error("Target text is empty.");

    let targetRange = null;

    // 1. חיפוש ישיר
    const searchResults = body.search(cleanTarget, { matchCase: false });
    searchResults.load("items");
    await context.sync();

    if (searchResults.items.length > 0) {
      targetRange = searchResults.items[0];
    } else {
      // 2. חיפוש חלקי אם המדויק נכשל
      if (cleanTarget.length > 25) {
        const shortTarget = cleanTarget.substring(0, 25).trim();
        const fallbackSearch = body.search(shortTarget, { matchCase: false });
        fallbackSearch.load("items");
        await context.sync();

        if (fallbackSearch.items.length > 0) {
          targetRange = fallbackSearch.items[0];
        }
      }
    }

    if (!targetRange) {
      throw new Error(`Target text not found for formatting: "${cleanTarget.substring(0, 35)}..."`);
    }

    if (formatOptions) {
      if (typeof formatOptions.bold === "boolean") targetRange.font.bold = formatOptions.bold;
      if (typeof formatOptions.underline === "boolean") targetRange.font.underline = formatOptions.underline ? "Single" : "None";
      if (typeof formatOptions.italic === "boolean") targetRange.font.italic = formatOptions.italic;
      if (formatOptions.font_name) targetRange.font.name = formatOptions.font_name;
    }

    await context.sync();
  });
}

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
