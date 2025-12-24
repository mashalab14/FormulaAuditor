/*****
 * Style Transformation Helper Functions
 * F0005
 ****/

/**
 * Applies new styles to memory arrays and handles cancellation
 * @param {Array} positions - Array of [row, col] formula positions
 * @param {Object} currentStyles - Current style arrays
 * @param {Object} newStyles - New styles to apply
 * @param {number} totalCols - Total number of columns for cancellation check
 * @return {number} - Number of cells formatted
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
