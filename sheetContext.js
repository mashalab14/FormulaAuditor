/*****
 * Sheet Context Helper Functions
 * F0005
 ****/

/**
 * Returns sheet context with range, name, dimensions
 * @return {Object} - Object containing sheet, range, sheetName, numRows, numCols
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
