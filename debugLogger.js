/*****
 * Debug Logging Helper Functions
 * F0005
 ****/

/**
 * Logs debug information about the sheet and range
 * @param {Object} context - Sheet context object with sheet info
 */
function logDebugInfo(context) {
  console.log("🔍 DEBUG: Active Sheet Name:", context.sheetName);
  console.log("🔍 DEBUG: Data Range Dimensions:", context.numRows, "rows x", context.numCols, "cols");
  
  // Aggressive verification: Check A1 specifically
  console.log("🔍 DEBUG: A1 Value:", context.range.getCell(1, 1).getValue());
  console.log("🔍 DEBUG: A1 Formula:", context.range.getCell(1, 1).getFormula());
}
