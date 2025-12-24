/*****
 * Style Writing Helper Functions
 * F0005
 ****/

/**
 * Writes all style changes to the sheet in one batch update
 * @param {Range} range - The range to apply styles to
 * @param {Object} styleData - Object containing all style arrays
 */
function writeStylesToSheet(range, styleData) {
  range.setFontWeights(styleData.fontWeights);
  range.setFontStyles(styleData.fontStyles);
  range.setFontLines(styleData.fontLines);
  range.setFontColors(styleData.fontColors);
  range.setBackgrounds(styleData.bgColors);
}
