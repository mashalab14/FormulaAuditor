/*****
 * Google Sheets Add-on: Formula Auditor
 * Main Code File
 ****/

/**
 * Creates the add-on menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Formula Auditor')
    .addItem('🎨 Format Formulas', 'showFormatSidebar')
    .addToUi();
}

/**
 * Shows the Format Tool sidebar
 */
function showFormatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('formatFormulaCells')
    .setTitle('Format Tool')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/*****
 * Apply Formatting
 * F0005
 ****/
function applyFormatting(styles) {
  if (!styles || typeof styles !== "object") {
    return "⚠️ Error: No formatting styles received.";
  }
  const BATCH_SIZE = 1000; // ✅ Defined inside the function (local scope)
  const cache = CacheService.getScriptCache();
  cache.remove("isCancelled"); // ✅ Reset cancellation flag before starting

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  const range = sheet.getDataRange().trimWhitespace(); // ✅ Process only the used range
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const formulas = range.getFormulas(); // ✅ Get all formulas at once
  const formulaPositions = [];

  let formattedCells = 0;
  let applied = false;

  let fontWeights = range.getFontWeights();
  let fontStyles = range.getFontStyles();
  let fontLines = range.getFontLines();
  let fontColors = range.getFontColors();
  let bgColors = range.getBackgrounds();

 // ✅ Step 1: Pre-filter formula positions before looping
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      if (formulas[row][col] && formulas[row][col].startsWith("=")) {
        formulaPositions.push([row, col]);
      }
    }
  }

  // ✅ Step 2: Iterate only over formula cells
  for (let i = 0; i < formulaPositions.length; i++) {
    if (i % numCols === 0 && cache.get("isCancelled") === "true") { 
      Logger.log("❌ Formatting stopped due to user cancellation.");
      return "⚠️ Operation cancelled by the user.";
    }

    let [row, col] = formulaPositions[i]; // ✅ Directly access pre-filtered formula cells

    // ✅ Apply styles in batch
    if (styles.bold !== undefined) fontWeights[row][col] = styles.bold ? "bold" : "normal";
    if (styles.italic !== undefined) fontStyles[row][col] = styles.italic ? "italic" : "normal";
    if (styles.underline !== undefined) fontLines[row][col] = styles.underline ? "underline" : "none";
    if (styles.strikethrough !== undefined) fontLines[row][col] = styles.strikethrough ? "line-through" : "none";
    if (styles.textColor) fontColors[row][col] = styles.textColor;
    if (styles.bgColor) bgColors[row][col] = styles.bgColor;

    formattedCells++;
    applied = true;

    // ✅ Flush only if more than BATCH_SIZE cells have been updated
    if (formattedCells % BATCH_SIZE === 0) {
      SpreadsheetApp.flush();
    }
  }

  // ✅ Apply all formatting in **one batch update**
  range.setFontWeights(fontWeights);
  range.setFontStyles(fontStyles);
  range.setFontLines(fontLines);
  range.setFontColors(fontColors);
  range.setBackgrounds(bgColors);

  return applied
    ? `✅ Formatting applied to ${formattedCells} cells with formulas in Sheet:"${sheetName}"`
    : "⚠️ No formula cells found in this sheet.";
}

/**
 * Cancel Button
 * F0005
 */
function cancelFormatting() {
  const cache = CacheService.getScriptCache();
  cache.put("isCancelled", "true", 600); // ✅ Store cancellation flag for 10 minutes
  return "⚠️ Operation cancelled by the user.";
}
