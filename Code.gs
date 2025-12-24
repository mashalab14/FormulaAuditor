/*****
 * Google Sheets Add-on: Formula Auditor
 * Main Orchestrator File
 * 
 * This file coordinates all helper modules:
 * - validation.js: Input validation
 * - sheetContext.js: Sheet and range setup
 * - formulaDetection.js: Formula cell detection
 * - styleFetcher.js: Current style retrieval
 * - styleTransformer.js: Style transformation in memory
 * - styleWriter.js: Bulk style application
 * - debugLogger.js: Debug information logging
 * - cancellation.js: Cancellation handling
 ****/

/**
 * Creates the add-on menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Formula Auditor')
    .addItem('🎨 Format Formula Cells', 'showFormatSidebar')
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
 * Apply Formatting - Main Orchestrator
 * F0005
 * 
 * Coordinates the formatting workflow:
 * Step 1: Validation (validation.js)
 * Step 2: Setup (sheetContext.js, debugLogger.js)
 * Step 3: Detection (formulaDetection.js)
 * Step 4: Transformation (styleTransformer.js)
 * Step 5: Execution (styleWriter.js)
 ****/
function applyFormatting(styles) {
  // Step 1: Validation
  if (!isValidInput(styles)) {
    return "⚠️ Error: No formatting styles received.";
  }

  // Step 2: Setup
  const cache = CacheService.getScriptCache();
  cache.remove("isCancelled"); // Reset cancellation flag before starting

  const context = getSheetContext();
  logDebugInfo(context);

  // Step 3: Detection
  const formulaPositions = findFormulaCells(context.range);
  console.log("🔍 DEBUG: Total formula cells found:", formulaPositions.length);

  if (formulaPositions.length === 0) {
    return "⚠️ No formula cells found in this sheet.";
  }

  // Step 4: Transformation
  const currentStyles = fetchCurrentStyles(context.range);
  const formattedCount = applyStylesToMemory(
    formulaPositions,
    currentStyles,
    styles,
    context.numCols
  );

  // Step 5: Execution
  writeStylesToSheet(context.range, currentStyles);

  return `✅ Formatting applied to ${formattedCount} cells with formulas in Sheet:"${context.sheetName}"`;
}

