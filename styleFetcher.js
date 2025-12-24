/*****
 * Style Fetching Helper Functions
 * F0005
 ****/

/**
 * Fetches all current formatting styles from the range
 * @param {Range} range - The range to fetch styles from
 * @return {Object} - Object containing all formatting arrays
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
