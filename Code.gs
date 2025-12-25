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
    // 1. Core Graph Features
    .addItem('🔍 Trace Precedents', 'showTraceDependents') 
    .addItem('🔄 Trace Dependents', 'showTraceDependents') // Placeholder
    .addItem('❌ Trace Errors', 'showTraceErrors')         // Placeholder
    .addItem('🏷️ Named Range Dependency', 'showTraceDependents')
    
    .addSeparator()
    
    // 2. Sidebar Utilities
    .addItem('🎨 Format Formula Cells', 'showFormatSidebar')
    .addItem('👀 Watch Window', 'showTraceDependents')
    .addItem('👀 Watch Window (M1)', 'showWatchWindowM1')
    .addItem('👀 Watch Window (M2)', 'showWatchWindowM2')
    
    .addToUi();
}
// --- Placeholders for Future Logic ---
function showTraceDependents() {
  SpreadsheetApp.getUi().alert("Coming Soon: Trace Dependents Graph");
}

function showTraceErrors() {
  SpreadsheetApp.getUi().alert("Coming Soon: Trace Errors Graph");
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

// --- MILESTONE 1 TEST FUNCTIONS ---

function showWatchWindowM1() {
  const html = HtmlService.createHtmlOutputFromFile('WatchWindow_M1')
    .setTitle('Watch Window M1')
    .setWidth(300); // Note: sidebar width may not be respected; keep as-is for now.
  SpreadsheetApp.getUi().showSidebar(html);
}

function testFetchA1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getRange(1, 1).getDisplayValue();
}

// --- MILESTONE 2: MANUAL WATCHER LOGIC ---

function showWatchWindowM2() {
  const html = HtmlService.createHtmlOutputFromFile('WatchWindow_M2')
    .setTitle('Watch Window (M2)')
    .setWidth(300); // Note: sidebar width may not be respected; keep as-is.
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Adds the active cell (active sheet + active cell) to the user's private watch list.
 * Stores sheetId + row + col to survive renames and inserts.
 */
function addActiveCellToWatchM2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  const newWatch = {
    id: Utilities.getUuid(),
    sheetId: sheet.getSheetId(),
    row: cell.getRow(),
    col: cell.getColumn(),
    createdAt: Date.now()
  };

  const userProps = PropertiesService.getUserProperties();
  const currentList = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  currentList.push(newWatch);
  userProps.setProperty('FA_WATCHES', JSON.stringify(currentList));
}

/**
 * Batch reads all watched cells and returns a normalized list including live display values.
 * Returns items shaped for the UI:
 * - ok: { id, sheetId, row, col, sheetName, cellRef, value, status: 'OK' }
 * - missing sheet: { id, sheetId, row, col, status: 'SHEET_MISSING', value: 'Error' }
 */
function getWatchListM2() {
  const userProps = PropertiesService.getUserProperties();
  const savedList = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  if (!Array.isArray(savedList) || savedList.length === 0) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return savedList.map(item => {
    const sheet = ss.getSheetById(item.sheetId);

    if (!sheet) {
      return {
        id: item.id,
        sheetId: item.sheetId,
        row: item.row,
        col: item.col,
        status: 'SHEET_MISSING',
        value: 'Error'
      };
    }

    const range = sheet.getRange(item.row, item.col);

    return {
      id: item.id,
      sheetId: item.sheetId,
      row: item.row,
      col: item.col,
      sheetName: sheet.getName(),
      cellRef: range.getA1Notation(),
      value: range.getDisplayValue(),
      status: 'OK'
    };
  });
}

/** Removes a specific watch item by ID from the user's private watch list. */
function removeWatchItemM2(idToRemove) {
  const userProps = PropertiesService.getUserProperties();
  const list = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  const next = (Array.isArray(list) ? list : []).filter(item => item.id !== idToRemove);
  userProps.setProperty('FA_WATCHES', JSON.stringify(next));
}

