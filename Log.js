/**
 * Configurable logging with levels and function filtering.
 *
 * Configuration is per-document (each spreadsheet has its own settings).
 *
 * Configuration via Document Properties:
 *   logLevel: "DEBUG" | "INFO" | "WARN" | "ERROR" | "OFF" (default: "INFO")
 *   logMode: "blacklist" | "whitelist" (default: "blacklist")
 *   logFunctions: comma-separated function tags to filter (default: "")
 *
 * Usage:
 *   log.debug("functionName", "Detailed message");
 *   log.info("functionName", "General info");
 *   log.warn("functionName", "Warning message");
 *   log.error("functionName", "Error message");
 *
 * Example configuration (in Script Properties):
 *   logLevel: "INFO"
 *   logMode: "blacklist"
 *   logFunctions: "recommendSpread,recommendClose,XLookupByKeys"
 */

const LogLevel = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  OFF: 4
};

/**
 * Get current log configuration from Script Properties.
 * Caches config for performance (refreshed each execution).
 */
function getLogConfig_() {
  if (getLogConfig_.cache) return getLogConfig_.cache;

  const props = PropertiesService.getDocumentProperties();
  const levelStr = (props.getProperty("logLevel") || "INFO").toUpperCase();
  const mode = (props.getProperty("logMode") || "blacklist").toLowerCase();
  const functionsStr = props.getProperty("logFunctions") || "";

  const functions = functionsStr
    .split(",")
    .map(s => s.trim().toLowerCase())
    .filter(s => s.length > 0);

  getLogConfig_.cache = {
    level: LogLevel[levelStr] ?? LogLevel.INFO,
    mode: mode,
    functions: new Set(functions)
  };

  return getLogConfig_.cache;
}

/**
 * Check if logging is enabled for the given level and tag.
 */
function shouldLog_(level, tag) {
  const config = getLogConfig_();

  // Check level first
  if (level < config.level) return false;

  // Check function filter
  if (config.functions.size === 0) return true;

  const tagLower = (tag || "").toLowerCase();
  const isInList = config.functions.has(tagLower);

  if (config.mode === "whitelist") {
    return isInList; // Only log if tag is in whitelist
  } else {
    return !isInList; // Log unless tag is in blacklist
  }
}

/**
 * Format and output a log message.
 */
function logMessage_(level, levelName, tag, message) {
  if (!shouldLog_(level, tag)) return;

  const prefix = tag ? `[${levelName}] ${tag}: ` : `[${levelName}] `;
  Logger.log(prefix + message);
}

/**
 * Log object with level methods.
 */
const log = {
  /**
   * Debug level - verbose details for troubleshooting.
   * @param {string} tag - Function or module name for filtering
   * @param {string} message - Log message
   */
  debug: function(tag, message) {
    logMessage_(LogLevel.DEBUG, "DEBUG", tag, message);
  },

  /**
   * Info level - general operational messages.
   * @param {string} tag - Function or module name for filtering
   * @param {string} message - Log message
   */
  info: function(tag, message) {
    logMessage_(LogLevel.INFO, "INFO", tag, message);
  },

  /**
   * Warn level - potential issues that don't stop execution.
   * @param {string} tag - Function or module name for filtering
   * @param {string} message - Log message
   */
  warn: function(tag, message) {
    logMessage_(LogLevel.WARN, "WARN", tag, message);
  },

  /**
   * Error level - errors that affect functionality.
   * @param {string} tag - Function or module name for filtering
   * @param {string} message - Log message
   */
  error: function(tag, message) {
    logMessage_(LogLevel.ERROR, "ERROR", tag, message);
  }
};

/**
 * Configure logging programmatically (alternative to Script Properties).
 *
 * @param {Object} options
 * @param {string} [options.level] - "DEBUG", "INFO", "WARN", "ERROR", "OFF"
 * @param {string} [options.mode] - "blacklist" or "whitelist"
 * @param {string[]} [options.functions] - Array of function tags to filter
 */
function configureLogging(options) {
  const props = PropertiesService.getDocumentProperties();

  if (options.level) {
    props.setProperty("logLevel", options.level.toUpperCase());
  }
  if (options.mode) {
    props.setProperty("logMode", options.mode.toLowerCase());
  }
  if (options.functions) {
    props.setProperty("logFunctions", options.functions.join(","));
  }

  // Clear cache to pick up new config
  getLogConfig_.cache = null;
}

/**
 * Show current logging configuration.
 */
function showLogConfig() {
  const config = getLogConfig_();
  const levelName = Object.keys(LogLevel).find(k => LogLevel[k] === config.level);

  Logger.log("=== Log Configuration ===");
  Logger.log(`Level: ${levelName}`);
  Logger.log(`Mode: ${config.mode}`);
  Logger.log(`Functions: ${config.functions.size > 0 ? Array.from(config.functions).join(", ") : "(none)"}`);
}

/**
 * Helper to quickly blacklist noisy functions.
 * @param {...string} tags - Function tags to blacklist
 */
function blacklistLogs(...tags) {
  const props = PropertiesService.getDocumentProperties();
  const existing = props.getProperty("logFunctions") || "";
  const existingSet = new Set(existing.split(",").map(s => s.trim()).filter(s => s));

  tags.forEach(t => existingSet.add(t));

  props.setProperty("logMode", "blacklist");
  props.setProperty("logFunctions", Array.from(existingSet).join(","));
  getLogConfig_.cache = null;

  Logger.log(`Blacklisted: ${tags.join(", ")}`);
}

/**
 * Helper to set whitelist mode with specific functions.
 * @param {...string} tags - Function tags to whitelist (only these will log)
 */
function whitelistLogs(...tags) {
  const props = PropertiesService.getDocumentProperties();

  props.setProperty("logMode", "whitelist");
  props.setProperty("logFunctions", tags.join(","));
  getLogConfig_.cache = null;

  Logger.log(`Whitelist mode - only logging: ${tags.join(", ")}`);
}

/**
 * Reset logging to defaults (INFO level, no filtering).
 */
function resetLogConfig() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty("logLevel");
  props.deleteProperty("logMode");
  props.deleteProperty("logFunctions");
  getLogConfig_.cache = null;

  Logger.log("Log configuration reset to defaults (INFO level, no filtering)");
}
