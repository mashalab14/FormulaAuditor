/*****
 * Formula Detection Helper Functions
 * F0005
 ****/

/**
 * Scans the grid and returns array of [row, col] for all formula cells
 * @param {Range} range - The range to scan for formulas
 * @return {Array} - Array of [row, col] coordinates for formula cells
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
