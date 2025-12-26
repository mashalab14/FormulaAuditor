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
    // F001-Precedence: Add Trace Precedents menu item
    .addItem('🔍 Trace Precedents', 'showTracePrecedents') 
    .addItem('🔄 Trace Dependents', 'showTraceDependents') // Placeholder
    .addItem('❌ Trace Errors', 'showTraceErrors')         // Placeholder
//    .addItem('🏷️ Named Range Dependency', 'showTraceDependents')
    
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
// F001-Precedence: Placeholder for Trace Precedents logic
function showTracePrecedents() {
  const html = HtmlService.createHtmlOutputFromFile('DependencyGraph.html')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Trace Precedents');
}
/**
 * Shows the unified Formula Auditor sidebar with specified initial tab
 */
function showFormulaAuditorSidebar(initialTab) {
  const template = HtmlService.createTemplateFromFile('FormulaAuditorSidebar');
  template.initialTab = initialTab || 'format';
  const html = template.evaluate()
    .setTitle('Formula Auditor')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the Format Tool sidebar (legacy entry point)
 */
function showFormatSidebar() {
  showFormulaAuditorSidebar('format');
}

/**
 * Shows the Watch Window sidebar (legacy entry point)
 */
function showWatchWindow() {
  showFormulaAuditorSidebar('watch');
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

// --- WATCH WINDOW ---

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

// F001-Precedence: Backend tracer logic for "Trace Precedents"
function getErrorTracerData() {
  const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  const sheet = cell.getSheet();
  const formula = cell.getFormula();
  const sheetName = sheet.getName();
  const selectedRef = `${sheetName}!${cell.getA1Notation()}`;

  if (!formula) {
    return {
      error: "Active cell does not contain a formula.",
      selectedRef: selectedRef
    };
  }

  const traced = getDependencies(cell); // Assumes getDependencies is defined elsewhere

  return {
    selectedRef: selectedRef,
    sheetName: sheetName,
    selectedCell: cell.getA1Notation(),
    precedents: traced.precedents || [],
    precedentLevels: traced.levelMap || {},
    cellDetails: traced.cellMap || {},
    edges: traced.edges || [],
    errorPathNodes: traced.errorPathNodes || [],
    errorPathEdges: traced.errorPathEdges || []
  };
}
// F001-Precedence: Client-side logic migrated to server-side

function parseFormula(formula) {
  try {
    return extractReferences(formula);
  } catch (e) {
    Logger.log("parseFormula error: " + e);
    return [];
  }
}

function extractReferences(formula) {
  const regex = /(?:'([^']+)'|([A-Za-z0-9_]+))?!?\$?([A-Z]+)\$?([0-9]+)/g;
  const references = new Set();
  let match;
  while ((match = regex.exec(formula)) !== null) {
    const sheet = match[1] || match[2] || '';
    const col = match[3];
    const row = match[4];
    const ref = sheet ? `${sheet}!${col}${row}` : `${col}${row}`;
    references.add(ref);
  }
  return Array.from(references);
}

function getDependencies(cell, visited = new Set(), levelMap = {}, edges = [], cellMap = {}, level = 0) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = cell.getSheet();
  const sheetName = sheet.getName();
  const ref = `${sheetName}!${cell.getA1Notation()}`;
  if (visited.has(ref)) return;

  visited.add(ref);
  levelMap[ref] = level;
  const formula = cell.getFormula();

  cellMap[ref] = {
    sheet: sheetName,
    cell: cell.getA1Notation(),
    value: cell.getDisplayValue(),
    formula: formula,
    type: formula ? 'formulaCell' : 'nonFormulaCell',
    isErrorValue: typeof cell.getDisplayValue() === 'string' && cell.getDisplayValue().startsWith('#')
  };

  if (!formula) return;

  const references = parseFormula(formula);
  references.forEach(refStr => {
    let range;
    try {
      range = ss.getRange(refStr);
    } catch (e) {
      Logger.log("Invalid reference: " + refStr);
      return;
    }

    const rangeSheet = range.getSheet();
    const rangeSheetName = rangeSheet.getName();
    const rangeRef = `${rangeSheetName}!${range.getA1Notation()}`;
    edges.push({ from: rangeRef, to: ref });

    getDependencies(range, visited, levelMap, edges, cellMap, level + 1);
  });

  return {
    precedents: Array.from(visited).filter(x => x !== ref),
    levelMap,
    cellMap,
    edges
  };
}