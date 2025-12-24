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

/*****
 * Helper Functions - Single Responsibility
 ****/

/**
 * Validates if the styles object is valid
 */
function isValidInput(styles) {
  return styles && typeof styles === "object";
}

/**
 * Returns sheet context with range, name, dimensions
 */
function getSheetContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Force full grid range for safety
  const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  return {
    sheet: sheet,
    range: range,
    sheetName: sheetName,
    numRows: numRows,
    numCols: numCols
  };
}

/**
 * Scans the grid and returns array of [row, col] for all formula cells
 */
function findFormulaCells(range) {
  const formulas = range.getFormulas();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const formulaPositions = [];

  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cellValue = formulas[row][col];
      // Safety Check: Ensure proper detection of formulas
      if (cellValue && cellValue.toString().startsWith("=")) {
        formulaPositions.push([row, col]);
        // Debug Logging: Print first 3 formulas detected
        if (formulaPositions.length <= 3) {
          console.log(`🔍 DEBUG: Formula ${formulaPositions.length} at [${row}, ${col}]:`, cellValue);
        }
      }
    }
  }

  return formulaPositions;
}

/**
 * Fetches all current formatting styles from the range
 */
function fetchCurrentStyles(range) {
  return {
    fontWeights: range.getFontWeights(),
    fontStyles: range.getFontStyles(),
    fontLines: range.getFontLines(),
    fontColors: range.getFontColors(),
    bgColors: range.getBackgrounds()
  };
}

/**
 * Applies new styles to memory arrays and handles cancellation
 */
function applyStylesToMemory(positions, currentStyles, newStyles, totalCols) {
  const BATCH_SIZE = 1000;
  const cache = CacheService.getScriptCache();
  let formattedCells = 0;

  for (let i = 0; i < positions.length; i++) {
    // Check for cancellation periodically
    if (i % totalCols === 0 && cache.get("isCancelled") === "true") {
      Logger.log("❌ Formatting stopped due to user cancellation.");
      throw new Error("⚠️ Operation cancelled by the user.");
    }

    const [row, col] = positions[i];

    // Apply styles in batch to memory
    if (newStyles.bold !== undefined) {
      currentStyles.fontWeights[row][col] = newStyles.bold ? "bold" : "normal";
    }
    if (newStyles.italic !== undefined) {
      currentStyles.fontStyles[row][col] = newStyles.italic ? "italic" : "normal";
    }
    if (newStyles.underline !== undefined) {
      currentStyles.fontLines[row][col] = newStyles.underline ? "underline" : "none";
    }
    if (newStyles.strikethrough !== undefined) {
      currentStyles.fontLines[row][col] = newStyles.strikethrough ? "line-through" : "none";
    }
    if (newStyles.textColor) {
      currentStyles.fontColors[row][col] = newStyles.textColor;
    }
    if (newStyles.bgColor) {
      currentStyles.bgColors[row][col] = newStyles.bgColor;
    }

    formattedCells++;

    // Flush periodically for large datasets
    if (formattedCells % BATCH_SIZE === 0) {
      SpreadsheetApp.flush();
    }
  }

  return formattedCells;
}

/**
 * Writes all style changes to the sheet in one batch update
 */
function writeStylesToSheet(range, styleData) {
  range.setFontWeights(styleData.fontWeights);
  range.setFontStyles(styleData.fontStyles);
  range.setFontLines(styleData.fontLines);
  range.setFontColors(styleData.fontColors);
  range.setBackgrounds(styleData.bgColors);
}

/**
 * Logs debug information about the sheet and range
 */
function logDebugInfo(context) {
  console.log("🔍 DEBUG: Active Sheet Name:", context.sheetName);
  console.log("🔍 DEBUG: Data Range Dimensions:", context.numRows, "rows x", context.numCols, "cols");
  
  // Aggressive verification: Check A1 specifically
  console.log("🔍 DEBUG: A1 Value:", context.range.getCell(1, 1).getValue());
  console.log("🔍 DEBUG: A1 Formula:", context.range.getCell(1, 1).getFormula());
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
