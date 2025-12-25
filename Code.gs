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
    .addItem('👀 Watch Window', 'showWatchWindow')
    
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

// --- MILESTONE 3: PROFESSIONAL WATCH WINDOW ---

function showWatchWindow() {
  const html = HtmlService.createHtmlOutputFromFile('WatchWindow')
    .setTitle('Formula Auditor')
    .setWidth(320); // Note: sidebar width may not be respected; keep as-is.
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Adds the active cell to FA_WATCHES with dedupe.
 * Dedupe rule: same sheetId + row + col should not be added twice.
 * Returns: { added: boolean, reason?: string }
 */
function addActiveCellToWatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  const userProps = PropertiesService.getUserProperties();
  const list = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  const arr = Array.isArray(list) ? list : [];

  const sheetId = sheet.getSheetId();
  const row = cell.getRow();
  const col = cell.getColumn();

  const isDup = arr.some(x => x && x.sheetId === sheetId && x.row === row && x.col === col);
  if (isDup) {
    return { added: false, reason: 'DUPLICATE' };
  }

  const newWatch = {
    id: Utilities.getUuid(),
    sheetId: sheetId,
    row: row,
    col: col,
    createdAt: Date.now()
  };

  arr.push(newWatch);
  userProps.setProperty('FA_WATCHES', JSON.stringify(arr));
  return { added: true };
}

/**
 * Returns a full refresh payload for the UI.
 * Must be ONE call to return everything needed to render the list.
 *
 * Return shape:
 * {
 *   serverNow: number,
 *   items: Array<
 *     | { id, sheetId, row, col, status: 'SHEET_MISSING' }
 *     | { id, sheetId, row, col, sheetName, cellRef, value, status: 'OK', isErrorValue: boolean }
 *   >
 * }
 */
function getWatchList() {
  const userProps = PropertiesService.getUserProperties();
  const saved = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  const arr = Array.isArray(saved) ? saved : [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const items = arr.map(item => {
    const sheet = ss.getSheetById(item.sheetId);

    if (!sheet) {
      return {
        id: item.id,
        sheetId: item.sheetId,
        row: item.row,
        col: item.col,
        status: 'SHEET_MISSING'
      };
    }

    const range = sheet.getRange(item.row, item.col);
    const value = range.getDisplayValue();

    const isErrorValue =
      typeof value === 'string' &&
      value.length > 0 &&
      value[0] === '#'; // simple and reliable for Sheets error display strings

    return {
      id: item.id,
      sheetId: item.sheetId,
      row: item.row,
      col: item.col,
      sheetName: sheet.getName(),
      cellRef: range.getA1Notation(),
      value: value,
      status: 'OK',
      isErrorValue: isErrorValue
    };
  });

  return {
    serverNow: Date.now(),
    items: items
  };
}

/** Removes a specific watch item by ID from FA_WATCHES. Returns { removed: boolean }. */
function removeWatchItem(idToRemove) {
  const userProps = PropertiesService.getUserProperties();
  const saved = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  const arr = Array.isArray(saved) ? saved : [];

  const next = arr.filter(x => x && x.id !== idToRemove);
  userProps.setProperty('FA_WATCHES', JSON.stringify(next));

  return { removed: next.length !== arr.length };
}

/** Removes all watches whose sheet is missing. Returns { removedCount: number }. */
function removeMissingWatches() {
  const userProps = PropertiesService.getUserProperties();
  const saved = JSON.parse(userProps.getProperty('FA_WATCHES') || '[]');
  const arr = Array.isArray(saved) ? saved : [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let removedCount = 0;
  const next = arr.filter(item => {
    const sheet = ss.getSheetById(item.sheetId);
    if (!sheet) {
      removedCount += 1;
      return false;
    }
    return true;
  });

  userProps.setProperty('FA_WATCHES', JSON.stringify(next));
  return { removedCount: removedCount };
}

/**
 * Jump to a watched cell by sheetId + row + col.
 * Must survive sheet rename.
 * Returns { ok: boolean, reason?: string }.
 */
function jumpToWatchCell(sheetId, row, col) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetById(sheetId);
  if (!sheet) return { ok: false, reason: 'SHEET_MISSING' };

  try {
    ss.setActiveSheet(sheet);
    const range = sheet.getRange(row, col);
    sheet.setActiveRange(range);
    return { ok: true };
  } catch (e) {
    return { ok: false, reason: 'RANGE_ERROR' };
  }
}

