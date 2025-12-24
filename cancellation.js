/*****
 * Cancellation Helper Functions
 * F0005
 ****/

/**
 * Cancel Button - Sets cancellation flag in cache
 * @return {string} - Cancellation message
 */
function cancelFormatting() {
  const cache = CacheService.getScriptCache();
  cache.put("isCancelled", "true", 600); // Store cancellation flag for 10 minutes
  return "⚠️ Operation cancelled by the user.";
}
