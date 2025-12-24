/*****
 * Validation Helper Functions
 * F0005
 ****/

/**
 * Validates if the styles object is valid
 * @param {Object} styles - The styles object to validate
 * @return {boolean} - True if valid, false otherwise
 */
function isValidInput(styles) {
  return styles && typeof styles === "object";
}
