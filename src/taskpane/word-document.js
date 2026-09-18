console.log("🔥  word-document.js LOADED SUCCESSFULLY!");

/**
 * Strips hidden directional characters, standardizes quotes, and normalizes whitespace
 */
window.normalizeText = function(str) {
  if (!str) return "";
  return str
    .replace(/[\u200B-\u200D\uFEFF\u200E\u200F]/g, "") // Remove hidden directional marks
    .replace(/[\u201C\u201D\u05F4"]/g, '"')             // Normalize double quotation marks
    .replace(/[\u2018\u2019\u05F3']/g, "'")             // Normalize single quotation marks
    .replace(/[\r\n\t]+/g, " ")                         // Convert line breaks and tabs to spaces
    .replace(/\s+/g, " ")                               // Collapse multiple spaces into one
    .trim();
};

/**
 * Finds and replaces a target text substring within the active Word document
 */
window.replaceDocumentSubstring = async function(targetText, replacementContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) throw new Error("Target text is empty.");

    // Step 1: Standard direct Word search (fast path for shorter strings)
    if (cleanTarget.length <= 250) {
      const searchResults = body.search(cleanTarget, { 
        matchCase: false, 
        matchWholeWord: false 
      });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > 0) {
        await window.applyContentToRange(context, searchResults.items[0], replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    // Step 2: Scan paragraphs with normalized comparison (handles quote/spacing differences)
    const paragraphs = body.paragraphs;
    paragraphs.load(["text", "items"]);
    await context.sync();

    const normalizedTarget = window.normalizeText(cleanTarget);

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const normalizedPText = window.normalizeText(p.text);

      if (!normalizedPText) continue;

      if (normalizedPText.includes(normalizedTarget) || normalizedTarget.includes(normalizedPText)) {
        await window.applyContentToRange(context, p, replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    // Step 3: Fallback prefix search if the end was truncated or slightly altered
    if (cleanTarget.length > 25) {
      const shortTarget = cleanTarget.substring(0, 25).trim();
      const fallbackSearch = body.search(shortTarget, { matchCase: false });
      fallbackSearch.load("items");
      await context.sync();

      if (fallbackSearch.items.length > 0) {
        await window.applyContentToRange(context, fallbackSearch.items[0], replacementContent, Word.InsertLocation.replace);
        return;
      }
    }

    throw new Error(`Target text not found: "${cleanTarget.substring(0, 35)}..."`);
  });
};

/**
 * Replaces the entire body content of the active Word document
 */
window.replaceEntireDocument = async function(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await window.applyContentToRange(context, body, newContent, Word.InsertLocation.replace);
  });
};

/**
 * Appends new text or HTML content to the end of the Word document
 */
window.insertTextAtEnd = async function(newContent) {
  return Word.run(async (context) => {
    const body = context.document.body;
    await window.applyContentToRange(context, body, "\n" + newContent, Word.InsertLocation.end);
  });
};

/**
 * Applies font styling (bold, italic, underline, font family) to matching text
 */
window.formatDocumentSubstring = async function(targetText, formatOptions) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const cleanTarget = (targetText || "").trim();

    if (!cleanTarget) throw new Error("Target text is empty.");

    let targetRange = null;

    // 1. Exact search
    const searchResults = body.search(cleanTarget, { matchCase: false });
    searchResults.load("items");
    await context.sync();

    if (searchResults.items.length > 0) {
      targetRange = searchResults.items[0];
    } else {
      // 2. Partial prefix search if exact match fails
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

    // Apply requested formatting options
    if (formatOptions) {
      if (typeof formatOptions.bold === "boolean") targetRange.font.bold = formatOptions.bold;
      if (typeof formatOptions.underline === "boolean") targetRange.font.underline = formatOptions.underline ? "Single" : "None";
      if (typeof formatOptions.italic === "boolean") targetRange.font.italic = formatOptions.italic;
      if (formatOptions.font_name) targetRange.font.name = formatOptions.font_name;
    }

    await context.sync();
  });
};

/**
 * Injects HTML or plain text into a specified Word range and enforces RTL alignment
 */
window.applyContentToRange = async function(context, targetRange, content, insertLocation) {
  let insertedRange;
  // Detect if content contains HTML markup
  const hasHtml = /<[a-z][\s\S]*>/i.test(content);

  if (hasHtml) {
    // Wrap with dir="rtl" to guarantee proper bidirectional rendering
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
};