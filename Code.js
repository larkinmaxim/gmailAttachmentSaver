// ===== SCRIPT PROPERTIES SETUP =====

/**
 * Setup Script Properties for secure credential storage
 * Run this function once to configure authentication credentials
 */
function setupScriptProperties() {
  console.log("=== SETTING UP SCRIPT PROPERTIES ===");
  
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    
    // Set the required configuration properties
    scriptProperties.setProperties({
      // Jira configuration
      'DEFAULT_JIRA_URL': 'https://support.transporeon.com',
      
      // Storage configuration
      'SETTINGS_STORAGE_KEY': 'JIRA_SETTINGS',
      
      // Google Drive folder configuration
      'CUSTOMERS_FOLDER_ID': '1pi6fJzg5nncV7ADrcivI0gEAa9cjZLQp',
      
      // Logging configuration
      // 0 = ERROR, 1 = WARN, 2 = INFO (default), 3 = DEBUG
      'DEFAULT_LOG_LEVEL': '2'
    });
    
    console.log("✅ Script properties configured successfully:");
    console.log("  - DEFAULT_JIRA_URL: " + scriptProperties.getProperty('DEFAULT_JIRA_URL'));
    console.log("  - SETTINGS_STORAGE_KEY: " + scriptProperties.getProperty('SETTINGS_STORAGE_KEY'));
    console.log("  - CUSTOMERS_FOLDER_ID: " + scriptProperties.getProperty('CUSTOMERS_FOLDER_ID'));
    console.log("  - DEFAULT_LOG_LEVEL: " + scriptProperties.getProperty('DEFAULT_LOG_LEVEL') + " (" + getLogLevelName(parseInt(scriptProperties.getProperty('DEFAULT_LOG_LEVEL'), 10)) + ")");
    
    return true;
    
  } catch (error) {
    console.error("❌ Failed to setup script properties:", error.message);
    return false;
  }
}

// ===== FOLDER CONFIGURATION CONSTANTS =====

/**
 * Standard project subfolders created for each new project
 */
var PROJECT_SUBFOLDERS = [
  "01_System_Design",
  "02_Meet_Recordings",
  "03_Correspondence",
  "04_Project_Documentation"
];

/**
 * Subfolders created inside 04_Project_Documentation
 */
var DOC_SUBFOLDERS = [
  "Project_Management",
  "Custom_Bundle",
  "Carrier_Onboarding"
];

/**
 * Get Default Jira URL from Script Properties
 */
function getDefaultJiraURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('DEFAULT_JIRA_URL');
  
  if (!url) {
    // Fallback to hardcoded value for backward compatibility
    return 'https://support.transporeon.com';
  }
  
  return url;
}

/**
 * Get Settings Storage Key from Script Properties
 */
function getSettingsStorageKey() {
  var properties = PropertiesService.getScriptProperties();
  var key = properties.getProperty('SETTINGS_STORAGE_KEY');
  
  if (!key) {
    // Fallback to default value
    return 'JIRA_SETTINGS';
  }
  
  return key;
}

/**
 * Get Customers Folder ID from Script Properties
 * This is the parent folder where all customer folders are stored
 */
function getCustomersFolderId() {
  var properties = PropertiesService.getScriptProperties();
  var folderId = properties.getProperty('CUSTOMERS_FOLDER_ID');
  
  if (!folderId) {
    // Fallback to default value for backward compatibility
    return '1pi6fJzg5nncV7ADrcivI0gEAa9cjZLQp';
  }
  
  return folderId;
}

// ===== PERFORMANCE CONFIGURATION =====

/**
 * Log levels for controlling output verbosity
 * 0 = ERROR - Only critical errors
 * 1 = WARN  - Errors and warnings
 * 2 = INFO  - Normal operation logs (default)
 * 3 = DEBUG - Verbose diagnostic logging
 */
var LOG_LEVELS = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

/**
 * Default log level - can be overridden by user settings
 * Set to INFO (2) for production, DEBUG (3) for troubleshooting
 */
var DEFAULT_LOG_LEVEL = LOG_LEVELS.INFO;

/**
 * Enable/disable verbose diagnostic logging (legacy - for backward compatibility)
 * @deprecated Use LOG_LEVEL setting instead
 */
var ENABLE_VERBOSE_LOGGING = false;

/**
 * Cache duration for folder lookups (in milliseconds)
 * Default: 5 minutes = 300000ms
 */
var FOLDER_CACHE_DURATION = 300000;

/**
 * Get current log level - checks User Properties first (per-user override),
 * then Script Properties (global default), then falls back to hardcoded default
 * @returns {number} Current log level (0-3)
 */
function getLogLevel() {
  try {
    // First check User Properties (per-user override)
    var userProperties = PropertiesService.getUserProperties();
    var userLogLevel = userProperties.getProperty('LOG_LEVEL');
    if (userLogLevel !== null) {
      var level = parseInt(userLogLevel, 10);
      if (level >= 0 && level <= 3) {
        return level;
      }
    }
    
    // Then check Script Properties (global default)
    var scriptProperties = PropertiesService.getScriptProperties();
    var scriptLogLevel = scriptProperties.getProperty('DEFAULT_LOG_LEVEL');
    if (scriptLogLevel !== null) {
      var level = parseInt(scriptLogLevel, 10);
      if (level >= 0 && level <= 3) {
        return level;
      }
    }
  } catch (e) {
    // Fall through to hardcoded default
  }
  return DEFAULT_LOG_LEVEL;
}

/**
 * Get default log level from Script Properties
 * @returns {number} Default log level (0-3)
 */
function getDefaultLogLevel() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var logLevel = scriptProperties.getProperty('DEFAULT_LOG_LEVEL');
    if (logLevel !== null) {
      var level = parseInt(logLevel, 10);
      if (level >= 0 && level <= 3) {
        return level;
      }
    }
  } catch (e) {
    // Fall through to constant
  }
  return DEFAULT_LOG_LEVEL;
}

/**
 * Set default log level in Script Properties (applies to all users)
 * @param {number} level - Log level (0-3)
 */
function setDefaultLogLevel(level) {
  try {
    if (level < 0 || level > 3) {
      console.error("Invalid log level. Must be 0-3.");
      return false;
    }
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('DEFAULT_LOG_LEVEL', String(level));
    console.log("✅ Default log level set to: " + getLogLevelName(level) + " (applies to all users)");
    return true;
  } catch (e) {
    console.error("Failed to set default log level:", e.message);
    return false;
  }
}

/**
 * Set log level in user settings (per-user override)
 * @param {number} level - Log level (0-3)
 */
function setLogLevel(level) {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('LOG_LEVEL', String(level));
    console.log("✅ Log level set to: " + getLogLevelName(level));
  } catch (e) {
    console.error("Failed to set log level:", e.message);
  }
}

/**
 * Clear user's log level override (revert to Script Properties default)
 */
function clearUserLogLevel() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('LOG_LEVEL');
    console.log("✅ User log level cleared. Now using default: " + getLogLevelName(getDefaultLogLevel()));
  } catch (e) {
    console.error("Failed to clear user log level:", e.message);
  }
}

/**
 * Get human-readable log level name
 * @param {number} level - Log level number
 * @returns {string} Log level name
 */
function getLogLevelName(level) {
  switch (level) {
    case 0: return "ERROR";
    case 1: return "WARN";
    case 2: return "INFO";
    case 3: return "DEBUG";
    default: return "UNKNOWN";
  }
}

/**
 * Log error message (always shown)
 * @param {string} message - Message to log
 */
function logError(message) {
  console.error("❌ " + message);
}

/**
 * Log warning message (shown at WARN level and above)
 * @param {string} message - Message to log
 */
function logWarn(message) {
  if (getLogLevel() >= LOG_LEVELS.WARN) {
    console.log("⚠️ " + message);
  }
}

/**
 * Log info message (shown at INFO level and above)
 * @param {string} message - Message to log
 */
function logInfo(message) {
  if (getLogLevel() >= LOG_LEVELS.INFO) {
    console.log(message);
  }
}

/**
 * Log debug message (shown only at DEBUG level)
 * @param {string} message - Message to log
 */
function logDebug(message) {
  if (getLogLevel() >= LOG_LEVELS.DEBUG) {
    console.log("🔍 " + message);
  }
}

/**
 * Conditional logging - only logs if verbose logging is enabled
 * @deprecated Use logDebug() instead
 */
function verboseLog(message) {
  if (ENABLE_VERBOSE_LOGGING || getLogLevel() >= LOG_LEVELS.DEBUG) {
    console.log(message);
  }
}

// ===== FOLDER SEARCH OPTIMIZATION =====

/**
 * Fast folder search using Drive API query instead of iteration
 * This is MUCH faster than iterating through all folders
 * 
 * @param {string} parentFolderId - The parent folder ID to search in
 * @param {string} searchPattern - Text to search for in folder name
 * @returns {Object} Search result with folder info
 */
function findFolderByNamePattern(parentFolderId, searchPattern) {
  try {
    var startTime = new Date().getTime();
    
    // Use Drive API query - optimized for Shared Drives
    var query = "mimeType='application/vnd.google-apps.folder' and '" + parentFolderId + "' in parents and trashed=false and name contains '" + searchPattern + "'";
    
    var results = Drive.Files.list({
      q: query,
      fields: 'files(id,name)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      corpora: 'allDrives',  // IMPORTANT: Optimizes Shared Drive search
      pageSize: 5  // We only need the first match
    });
    
    var duration = new Date().getTime() - startTime;
    logDebug("⏱️ Drive search: " + duration + "ms, found " + (results.files ? results.files.length : 0) + " for '" + searchPattern + "'");
    
    if (results.files && results.files.length > 0) {
      var folder = results.files[0];
      return {
        success: true,
        found: true,
        folder: {
          id: folder.id,
          name: folder.name
        },
        duration: duration
      };
    }
    
    return {
      success: true,
      found: false,
      duration: duration
    };
    
  } catch (error) {
    logError("Drive search failed: " + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get cached folder lookup or perform search
 * Caches folder IDs to avoid repeated searches
 * 
 * @param {string} cacheKey - Unique cache key (e.g., "customer_246028")
 * @param {string} parentFolderId - Parent folder to search in
 * @param {string} searchPattern - Pattern to search for
 * @returns {Object} Folder result (from cache or fresh search)
 */
function getCachedOrSearchFolder(cacheKey, parentFolderId, searchPattern) {
  try {
    var cache = CacheService.getUserCache();
    var cached = cache.get(cacheKey);
    
    if (cached) {
      var cachedData = JSON.parse(cached);
      logDebug("📦 Cache HIT for " + cacheKey + " - folder: " + cachedData.name);
      return {
        success: true,
        found: true,
        fromCache: true,
        folder: cachedData
      };
    }
    
    logDebug("📦 Cache MISS for " + cacheKey + " - searching...");
    
    // Perform search
    var searchResult = findFolderByNamePattern(parentFolderId, searchPattern);
    
    if (searchResult.success && searchResult.found) {
      // Cache the result (6 hours = 21600 seconds max for Apps Script cache)
      cache.put(cacheKey, JSON.stringify(searchResult.folder), 21600);
      logDebug("📦 Cached folder: " + searchResult.folder.name);
    }
    
    return searchResult;
    
  } catch (error) {
    logError("Cache error, falling back to search: " + error.message);
    return findFolderByNamePattern(parentFolderId, searchPattern);
  }
}

/**
 * Clear folder cache for a specific customer or all
 * @param {string} customerId - Optional customer ID to clear, or null for all
 */
function clearFolderCache(customerId) {
  try {
    var cache = CacheService.getUserCache();
    if (customerId) {
      cache.remove("customer_" + customerId);
      cache.remove("project_" + customerId);
      console.log("Cleared cache for customer:", customerId);
    } else {
      // Can't clear all in Apps Script, but individual keys can be removed
      console.log("Cache cleared for specified keys");
    }
  } catch (error) {
    console.error("Error clearing cache:", error.message);
  }
}

/**
 * Clear all folder cache entries (called from UI)
 * @returns {CardService.ActionResponse} Notification response
 */
function clearAllFolderCache() {
  try {
    var cache = CacheService.getUserCache();
    
    // Remove known cache key patterns
    // Note: Apps Script cache doesn't support listing all keys,
    // so we clear common patterns
    var clearedCount = 0;
    
    // Try to clear any cached folder lookups
    // The cache will naturally expire after 6 hours anyway
    cache.removeAll(['customer_', 'project_']);
    
    console.log("✅ Folder cache cleared");
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✅ Folder cache cleared! Next folder lookups will be fresh."))
      .build();
      
  } catch (error) {
    console.error("Error clearing cache:", error.message);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("⚠️ Cache clear attempted: " + error.message))
      .build();
  }
}

// ===== JIRA SUMMARY PARSING FUNCTIONS =====

/**
 * Parse Jira summary field to extract folder names
 * @param {string} summary - Jira summary field value
 * @param {string} ticketKey - Jira ticket key (e.g., "CXPRODELIVERY-4750")
 * @returns {Object} Parsed components and folder names
 */
function parseJiraSummaryForFolders(summary, ticketKey) {
  verboseLog("=".repeat(80));
  verboseLog("    PARSE JIRA SUMMARY FOR FOLDER NAMES");
  verboseLog("=".repeat(80));
  
  if (!summary) {
    console.error("❌ No summary provided");
    return { success: false, error: "No summary provided" };
  }
  
  if (!ticketKey) {
    console.error("❌ No ticket key provided");
    return { success: false, error: "No ticket key provided" };
  }
  
  verboseLog("📋 Input: " + ticketKey + " - " + summary.substring(0, 50) + "...");
  
  var result = {
    success: false,
    original: {
      summary: summary,
      ticketKey: ticketKey
    },
    parsed: {
      customerId: null,
      customerName: null,
      projectType: null,
      scopeDescription: null
    },
    folderNames: {
      customerFolder: null,
      projectFolder: null
    }
  };
  
  // Pattern: [CustomerID] - [Customer Name] | [Project Type] | [Scope Description]
  var parts = summary.split(' | ');
  
  if (parts.length < 2) {
    console.error("❌ Summary doesn't match expected pattern");
    result.error = "Summary doesn't match expected pattern (missing ' | ' separators)";
    return result;
  }
  
  // Part 1: "[CustomerID] - [Customer Name]"
  var customerPart = parts[0].trim();
  var customerMatch = customerPart.match(/^(\d+)\s*-\s*(.+)$/);
  
  if (customerMatch) {
    result.parsed.customerId = customerMatch[1].trim();
    result.parsed.customerName = customerMatch[2].trim();
  } else {
    verboseLog("⚠️ Could not parse customer ID from: " + customerPart);
    result.parsed.customerName = customerPart;
  }
  
  // Part 2: "[Project Type]"
  if (parts.length >= 2) {
    result.parsed.projectType = parts[1].trim();
  }
  
  // Part 3+: "[Scope Description]"
  if (parts.length >= 3) {
    result.parsed.scopeDescription = parts.slice(2).join(' | ').trim();
  }
  
  // Generate folder names
  if (result.parsed.customerId && result.parsed.customerName) {
    result.folderNames.customerFolder = result.parsed.customerId + " - " + result.parsed.customerName;
  } else if (result.parsed.customerName) {
    result.folderNames.customerFolder = result.parsed.customerName;
  }
  
  var projectParts = [ticketKey];
  if (result.parsed.projectType) {
    projectParts.push(result.parsed.projectType);
  }
  if (result.parsed.scopeDescription) {
    projectParts.push(result.parsed.scopeDescription);
  }
  result.folderNames.projectFolder = projectParts.join(' | ');
  
  console.log("📁 Parsed: Customer=" + result.parsed.customerId + ", Project=" + result.parsed.projectType);
  
  result.success = true;
  return result;
}

/**
 * Sanitize folder name for Google Drive
 * @param {string} folderName - Original folder name
 * @returns {string} Sanitized folder name
 */
function sanitizeFolderName(folderName) {
  if (!folderName) return "";
  
  var sanitized = folderName
    .replace(/[\/\\:*?"<>|]/g, '-')  // Replace forbidden chars with dash
    .replace(/\s+/g, ' ')             // Normalize whitespace
    .replace(/-+/g, '-')              // Remove duplicate dashes
    .replace(/^-|-$/g, '')            // Remove leading/trailing dashes
    .trim();
  
  // Limit length
  if (sanitized.length > 200) {
    sanitized = sanitized.substring(0, 200).trim();
  }
  
  return sanitized;
}

/**
 * Fetch Jira ticket and parse summary for folder names
 * @param {string} ticketKey - Jira ticket key
 * @returns {Object} Parse result with folder names
 */
function getJiraTicketFolderNames(ticketKey, cachedSettings) {
  verboseLog("=".repeat(80));
  verboseLog("    GET FOLDER NAMES FROM JIRA TICKET");
  verboseLog("=".repeat(80));
  
  if (!ticketKey) {
    return { success: false, error: "No ticket key provided" };
  }
  
  try {
    // Use cached settings if provided, otherwise load fresh
    var settings = cachedSettings || getUserSettings();
    
    if (!settings.jiraUrl || !settings.jiraToken) {
      console.error("❌ Jira not configured");
      return { success: false, error: "Jira not configured" };
    }
    
    var issueUrl = settings.jiraUrl + '/rest/api/2/issue/' + ticketKey;
    verboseLog("🔍 Fetching ticket " + ticketKey + " from Jira...");
    
    var response = UrlFetchApp.fetch(issueUrl, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + settings.jiraToken,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error("❌ Jira API failed:", response.getResponseCode());
      return { success: false, error: "Jira API failed: " + response.getResponseCode() };
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (!data.fields || !data.fields.summary) {
      console.error("❌ No summary field in ticket");
      return { success: false, error: "No summary field in ticket" };
    }
    
    var summary = data.fields.summary;
    console.log("📋 " + ticketKey + ": " + summary.substring(0, 60) + "...");
    
    var parseResult = parseJiraSummaryForFolders(summary, ticketKey);
    
    parseResult.ticketData = {
      key: data.key,
      id: data.id,
      status: data.fields.status ? data.fields.status.name : null
    };
    
    return parseResult;
    
  } catch (error) {
    console.error("❌ Error:", error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Create customer folder in Customers directory
 * @param {string} customerFolderName - Name for the customer folder
 * @returns {Object} Created folder info
 */
function createCustomerFolder(customerFolderName) {
  console.log("=".repeat(80));
  console.log("    CREATE CUSTOMER FOLDER");
  console.log("=".repeat(80));
  
  console.log("📁 Customers Folder ID:", getCustomersFolderId());
  console.log("📁 New Folder Name:", customerFolderName);
  
  try {
    var customersFolder = DriveApp.getFolderById(getCustomersFolderId());
    console.log("✅ Customers folder accessed:", customersFolder.getName());
    
    console.log("\nCreating customer folder...");
    var customerFolder = customersFolder.createFolder(customerFolderName);
    
    console.log("✅ Customer folder created!");
    console.log("   Name:", customerFolder.getName());
    console.log("   ID:", customerFolder.getId());
    console.log("   URL:", customerFolder.getUrl());
    
    return {
      success: true,
      created: true,
      folder: {
        name: customerFolder.getName(),
        id: customerFolder.getId(),
        url: customerFolder.getUrl()
      }
    };
    
  } catch (error) {
    console.error("❌ Failed to create customer folder");
    console.error("Error:", error.message);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Create project folder with standard subfolders
 * @param {string} parentFolderId - ID of the customer folder
 * @param {string} projectFolderName - Name for the project folder
 * @returns {Object} Created folder info
 */
function createProjectFolderStructure(parentFolderId, projectFolderName) {
  console.log("=".repeat(80));
  console.log("    CREATE PROJECT FOLDER STRUCTURE");
  console.log("=".repeat(80));
  
  console.log("📁 Parent Folder ID:", parentFolderId);
  console.log("📁 Project Folder Name:", projectFolderName);
  
  try {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    console.log("✅ Parent folder accessed:", parentFolder.getName());
    
    // Create main project folder
    console.log("\n1️⃣ Creating project folder...");
    var projectFolder = parentFolder.createFolder(projectFolderName);
    console.log("   ✅ Created:", projectFolder.getName());
    console.log("   ID:", projectFolder.getId());
    console.log("   URL:", projectFolder.getUrl());
    
    // Create main subfolders
    console.log("\n2️⃣ Creating subfolders...");
    var createdSubfolders = {};
    
    for (var i = 0; i < PROJECT_SUBFOLDERS.length; i++) {
      var subfolderName = PROJECT_SUBFOLDERS[i];
      var subfolder = projectFolder.createFolder(subfolderName);
      createdSubfolders[subfolderName] = subfolder;
      console.log("   ✅", subfolderName);
    }
    
    // Create documentation subfolders inside 04_Project_Documentation
    console.log("\n3️⃣ Creating documentation subfolders...");
    var docFolder = createdSubfolders["04_Project_Documentation"];
    
    for (var i = 0; i < DOC_SUBFOLDERS.length; i++) {
      var docSubfolderName = DOC_SUBFOLDERS[i];
      docFolder.createFolder(docSubfolderName);
      console.log("   ✅ 04_Project_Documentation/" + docSubfolderName);
    }
    
    console.log("\n✅ Folder structure created successfully!");
    console.log("🔗 Project Folder URL:", projectFolder.getUrl());
    
    return {
      success: true,
      projectFolder: {
        name: projectFolder.getName(),
        id: projectFolder.getId(),
        url: projectFolder.getUrl()
      },
      subfolders: PROJECT_SUBFOLDERS,
      docSubfolders: DOC_SUBFOLDERS
    };
    
  } catch (error) {
    console.error("❌ Failed to create folder structure");
    console.error("Error:", error.message);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Complete workflow: Get folder names from Jira and search/create folders
 * OPTIMIZED version with fast folder search and caching
 * 
 * @param {string} ticketKey - Jira ticket key
 * @param {boolean} createIfMissing - Whether to create folders if they don't exist
 * @param {Object} cachedSettings - Optional pre-loaded settings to avoid redundant loading
 * @returns {Object} Workflow result with folder info
 */
function completeJiraToFolderWorkflowWithCreate(ticketKey, createIfMissing, cachedSettings) {
  var workflowStart = new Date().getTime();
  console.log("\n🚀 JIRA → FOLDER WORKFLOW (OPTIMIZED) for " + ticketKey);
  
  if (!ticketKey) {
    return { success: false, error: "No ticket key provided" };
  }
  
  if (createIfMissing === undefined) {
    createIfMissing = true;
  }
  
  var result = {
    ticketKey: ticketKey,
    success: false,
    customerFolder: null,
    projectFolder: null,
    created: {
      customerFolder: false,
      projectFolder: false
    }
  };
  
  // Step 1: Get folder names from Jira (pass cached settings for performance)
  logDebug("STEP 1: PARSE JIRA TICKET");
  
  var parseResult = getJiraTicketFolderNames(ticketKey, cachedSettings);
  
  if (!parseResult.success) {
    logError("Failed to get folder names from Jira");
    result.error = parseResult.error;
    return result;
  }
  
  var customerFolderName = sanitizeFolderName(parseResult.folderNames.customerFolder);
  var projectFolderName = sanitizeFolderName(parseResult.folderNames.projectFolder);
  var customerId = parseResult.parsed.customerId;
  
  logInfo("📁 Looking for: Customer=" + customerId + ", Ticket=" + ticketKey);
  
  // Step 2: Find or create customer folder (OPTIMIZED with fast search + caching)
  logDebug("=".repeat(60));
  logDebug("STEP 2: FIND/CREATE CUSTOMER FOLDER (OPTIMIZED)");
  logDebug("=".repeat(60));
  
  var customerFolderObj = null;
  
  try {
    logDebug("Fast searching for customer ID: " + customerId);
    
    // Use cached/fast search instead of slow iteration
    var cacheKey = "customer_" + customerId;
    var searchResult = getCachedOrSearchFolder(cacheKey, getCustomersFolderId(), customerId);
    
    if (searchResult.success && searchResult.found) {
      // Get DriveApp folder object from the ID
      customerFolderObj = DriveApp.getFolderById(searchResult.folder.id);
      logInfo("✅ Found customer folder: " + searchResult.folder.name);
      if (searchResult.fromCache) {
        logDebug("   (from cache - instant!)");
      }
    }
    
    if (customerFolderObj) {
      result.customerFolder = {
        exists: true,
        name: customerFolderObj.getName(),
        id: customerFolderObj.getId(),
        url: customerFolderObj.getUrl()
      };
    } else if (createIfMissing) {
      logInfo("❌ Customer folder not found");
      logInfo("📝 Creating: " + customerFolderName);
      
      var createResult = createCustomerFolder(customerFolderName);
      
      if (createResult.success) {
        result.customerFolder = {
          exists: true,
          name: createResult.folder.name,
          id: createResult.folder.id,
          url: createResult.folder.url
        };
        result.created.customerFolder = true;
        customerFolderObj = DriveApp.getFolderById(createResult.folder.id);
      } else {
        logError("Failed to create customer folder");
        result.error = createResult.error;
        return result;
      }
    } else {
      logInfo("❌ Customer folder not found (creation disabled)");
      result.customerFolder = {
        exists: false,
        suggestedName: customerFolderName
      };
    }
    
  } catch (error) {
    logError("Error: " + error.message);
    result.error = error.message;
    return result;
  }
  
  // Step 3: Find or create project folder (OPTIMIZED with fast search + caching)
  if (result.customerFolder && result.customerFolder.exists) {
    logDebug("=".repeat(60));
    logDebug("STEP 3: FIND/CREATE PROJECT FOLDER (OPTIMIZED)");
    logDebug("=".repeat(60));
    
    try {
      logDebug("Fast searching for ticket: " + ticketKey);
      
      // Use cached/fast search instead of slow iteration
      var projectCacheKey = "project_" + ticketKey;
      var projectSearchResult = getCachedOrSearchFolder(projectCacheKey, result.customerFolder.id, ticketKey);
      var existingProjectFolder = null;
      
      if (projectSearchResult.success && projectSearchResult.found) {
        existingProjectFolder = DriveApp.getFolderById(projectSearchResult.folder.id);
        logInfo("✅ Found project folder: " + projectSearchResult.folder.name);
        if (projectSearchResult.fromCache) {
          logDebug("   (from cache - instant!)");
        }
      }
      
      if (existingProjectFolder) {
        result.projectFolder = {
          exists: true,
          name: existingProjectFolder.getName(),
          id: existingProjectFolder.getId(),
          url: existingProjectFolder.getUrl()
        };
      } else if (createIfMissing) {
        logInfo("❌ Project folder not found");
        logInfo("📝 Creating: " + projectFolderName);
        
        var createResult = createProjectFolderStructure(result.customerFolder.id, projectFolderName);
        
        if (createResult.success) {
          result.projectFolder = {
            exists: true,
            name: createResult.projectFolder.name,
            id: createResult.projectFolder.id,
            url: createResult.projectFolder.url
          };
          result.created.projectFolder = true;
        } else {
          logError("Failed to create project folder");
          result.projectFolder = {
            exists: false,
            error: createResult.error
          };
        }
      } else {
        logInfo("❌ Project folder not found (creation disabled)");
        result.projectFolder = {
          exists: false,
          suggestedName: projectFolderName
        };
      }
      
    } catch (error) {
      logError("Error: " + error.message);
    }
  }
  
  // Final Summary
  var workflowDuration = new Date().getTime() - workflowStart;
  
  result.success = true;
  result.parsed = parseResult.parsed;
  result.folderNames = {
    customerFolder: customerFolderName,
    projectFolder: projectFolderName
  };
  
  // Concise summary
  var customerStatus = result.customerFolder && result.customerFolder.exists 
    ? (result.created.customerFolder ? "🆕 CREATED" : "✅ EXISTS") 
    : "❌ NOT FOUND";
  var projectStatus = result.projectFolder && result.projectFolder.exists 
    ? (result.created.projectFolder ? "🆕 CREATED" : "✅ EXISTS") 
    : "❌ NOT FOUND";
  
  console.log("✅ Workflow completed in " + workflowDuration + "ms");
  console.log("   Customer: " + customerStatus + " | Project: " + projectStatus);
  
  return result;
}

// ===== MAIN GMAIL ADD-ON FUNCTIONS =====

function buildAddOn(e) {
    console.log("=== BUILD ADD-ON CALLED ===");
    
    try {
      var card = CardService.newCardBuilder();
      var section = CardService.newCardSection()
        .setHeader("Jira Project");
      
      // Check if settings are configured
      console.log("Checking user settings configuration...");
      var settings = getUserSettings();
      
      console.log("Settings validation results:");
      console.log("- Jira Token present:", !!settings.jiraToken);
      console.log("- Jira URL present:", !!settings.jiraUrl);
      console.log("- Customers Folder ID:", getCustomersFolderId());
      
      if (!settings.jiraToken || !settings.jiraUrl) {
        console.log("SETTINGS INCOMPLETE - showing setup screen");
        console.log("Missing:");
        if (!settings.jiraToken) console.log("- Jira Token");
        if (!settings.jiraUrl) console.log("- Jira URL");
        
        return [buildSettingsCard(true)]; // true = first time setup
      }
      
      console.log("Settings validation PASSED - proceeding with main interface");
      
      // Get user's TPM projects and active tickets
      try {
        console.log("Loading TPM projects...");
        var projectResult = getMyJiraProjects();
        var projects = projectResult.projects || {};
        var issues = projectResult.issues || [];
        
        console.log("DEBUG: Projects loaded:", Object.keys(projects).length);
        console.log("DEBUG: Issues loaded:", issues.length);
        
        if (Object.keys(projects).length > 0 && issues.length > 0) {
          console.log("Creating TPM project UI with", issues.length, "tickets...");
          
          // Compact header with just settings button
          var headerSection = CardService.newCardSection();
          
          // Compact settings button (gear only)
          var settingsButtonSet = CardService.newButtonSet()
            .addButton(CardService.newTextButton()
              .setText("⚙️")
              .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
          headerSection.addWidget(settingsButtonSet);
          
          card.addSection(headerSection);
          
          // Create dropdown with active TPM tickets WITH onChange action
          console.log("Creating dropdown selection widget with dynamic info...");
          var ticketSelection = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setFieldName("selectedTicket")
            .setTitle("Select Active Ticket")
            .setOnChangeAction(CardService.newAction()
              .setFunctionName("showTicketDetails")
              .setParameters({threadId: (e && e.gmail) ? e.gmail.threadId : ""}));
          
          // Add "Manual Entry" option first
          console.log("Adding manual entry option...");
          ticketSelection.addItem("Manual Entry - Enter ticket number below", "manual", true);
          
          // Sort and add all active TPM tickets to dropdown
          issues.sort(function(a, b) {
            return a.key.localeCompare(b.key);
          });
          
          console.log("Adding", issues.length, "TPM tickets to dropdown...");
          
          var successfulItems = 0;
          for (var i = 0; i < issues.length; i++) {
            try {
              var issue = issues[i];
              
              // Create compact display text with emoji and shortened info
              var displayText = formatCompactTicketDisplay(issue);
              
              console.log("Adding ticket:", displayText);
              ticketSelection.addItem(displayText, issue.key, false);
              successfulItems++;
              
            } catch (itemError) {
              console.error("Error adding dropdown item for", issue.key, ":", itemError);
              // Continue with other items
              continue;
            }
          }
          
          console.log("Successfully added", successfulItems, "items to dropdown");
          section.addWidget(ticketSelection);
          
          // Add manual ticket number input
          var ticketInput = CardService.newTextInput()
            .setFieldName("manualTicketNumber")
            .setTitle("Manual Ticket Entry")
            .setHint("e.g., 6500 (will become CXPRODELIVERY-6500)");
          section.addWidget(ticketInput);
          
          // Add helpful tip
          var statusInfo = CardService.newTextParagraph()
            .setText("💡 Tip: Select from dropdown to see full ticket details");
          section.addWidget(statusInfo);
          
          console.log("Created main ticket selection interface with dropdown and dynamic info");
          
        } else {
          console.log("No TPM projects/issues found, using fallback UI");
          var debugInfo = CardService.newTextParagraph()
            .setText("⚠️ Could not load TPM tickets. Check your JQL or connection settings.");
          section.addWidget(debugInfo);
          
                // Settings button for troubleshooting
        var settingsButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("⚙️ Check Settings")
            .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
        section.addWidget(settingsButtonSet);
          
          var ticketInput = CardService.newTextInput()
            .setFieldName("jiraTicket")
            .setTitle("Manual Ticket Entry")
            .setHint("e.g., CXPRODELIVERY-6310");
          section.addWidget(ticketInput);
        }
        
      } catch (jiraError) {
        console.error('Failed to load TPM projects:', jiraError);
        var errorInfo = CardService.newTextParagraph()
          .setText("⚠️ Jira API error: " + jiraError.message + ". Check your settings.");
        section.addWidget(errorInfo);
        
        // Settings button for troubleshooting
        var settingsButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("⚙️ Fix Settings")
            .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
        section.addWidget(settingsButtonSet);
        
        var ticketInput = CardService.newTextInput()
          .setFieldName("jiraTicket")
          .setTitle("Full Jira Ticket")
          .setHint("e.g., CXPRODELIVERY-6310");
        section.addWidget(ticketInput);
      }
      
      // Handle Gmail context for attachments
      if (!e || !e.gmail || !e.gmail.threadId) {
        console.log("No Gmail context - using basic interface");
        var infoSection = CardService.newCardSection()
          .setHeader("Open an email to see attachments");
        
        var infoText = CardService.newTextParagraph()
          .setText("Please open an email with attachments to use this add-on.");
        infoSection.addWidget(infoText);
        card.addSection(infoSection);
        
        // Add save button
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Save to Project Folder")
            .setOnClickAction(CardService.newAction().setFunctionName("saveSelectedAttachmentsToGDrive")));
        section.addWidget(saveButtonSet);
        
        card.addSection(section);
        
        return [card.build()];
      }
      
      console.log("Gmail context found, loading attachments...");
      
      // Get attachments and Google Docs links from the email thread
      try {
        var threadId = e.gmail.threadId;
        var thread = GmailApp.getThreadById(threadId);
        var messages = thread.getMessages();
        var allAttachments = [];
        var allGoogleDocsLinks = [];
        
        for (var i = 0; i < messages.length; i++) {
          var message = messages[i];
          var attachments = message.getAttachments();
          
          // Process regular attachments
          for (var j = 0; j < attachments.length; j++) {
            var attachment = attachments[j];
            allAttachments.push({
              name: attachment.getName(),
              size: attachment.getSize(),
              attachment: attachment,
              messageIndex: i,
              type: 'regular'
            });
          }
          
          // Extract Google Docs links from message content
          var googleDocsLinks = extractGoogleDocsLinks(message);
          for (var k = 0; k < googleDocsLinks.length; k++) {
            var docLink = googleDocsLinks[k];
            allGoogleDocsLinks.push({
              name: docLink.title,
              size: 0, // Google Docs links don't have a file size
              url: docLink.url,
              docType: docLink.type,
              messageIndex: i,
              type: 'google_docs'
            });
          }
        }
        
        // Combine regular attachments and Google Docs links
        var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
        
                if (combinedAttachments.length > 0) {
            // Group attachments by file extension (including Google Docs links)
            var attachmentsByExtension = {};
            var extensionCounts = {};
            
            for (var i = 0; i < combinedAttachments.length; i++) {
              var attachment = combinedAttachments[i];
              var fileName = attachment.name;
              
              // Determine extension/category
              var extension;
              if (attachment.type === 'google_docs') {
                // Group Google Docs by their type (docs, sheets, slides, etc.)
                extension = 'google-' + attachment.docType;
              } else {
                // Regular file extension
                extension = fileName.lastIndexOf('.') > -1 ? 
                  fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              }
              
              if (!attachmentsByExtension[extension]) {
                attachmentsByExtension[extension] = [];
                extensionCounts[extension] = 0;
              }
              
              attachmentsByExtension[extension].push({
                attachment: attachment,
                index: i
              });
              extensionCounts[extension]++;
            }
            
            // Get stored attachment selections for this thread
            var storedSelections = getStoredAttachmentSelections(threadId);
            
            // Sort extensions alphabetically for consistent display
            var sortedExtensions = Object.keys(attachmentsByExtension).sort();
            
            // Create a section for each file extension
            for (var extIndex = 0; extIndex < sortedExtensions.length; extIndex++) {
              var extension = sortedExtensions[extIndex];
              var attachmentsInGroup = attachmentsByExtension[extension];
              var count = extensionCounts[extension];
              
              // Create section header with extension and count (show total in first section)
              var extDisplayName = extension === 'no-extension' ? 'Files without extension' : 
                extension.toUpperCase() + ' files';
              var headerText = extDisplayName + " (" + count + ")";
              if (extIndex === 0) {
                headerText = "📎 Select Attachments & Docs: " + headerText + " (Total: " + combinedAttachments.length + ")";
              }
              
              var extensionSection = CardService.newCardSection()
                .setHeader(headerText)
                .setCollapsible(true)
                .setNumUncollapsibleWidgets(count > 5 ? 2 : count); // Show first 2 if more than 5 files
              
              // Add checkboxes for each attachment in this extension group
              for (var j = 0; j < attachmentsInGroup.length; j++) {
                var attachmentData = attachmentsInGroup[j];
                var attachment = attachmentData.attachment;
                var originalIndex = attachmentData.index;
                
                // Check if this attachment was previously selected using unique key
                var isSelected = false; // Default to false - let user explicitly choose
                var uniqueKey = attachment.name + "_" + originalIndex;
                if (storedSelections && storedSelections.hasOwnProperty(uniqueKey)) {
                  isSelected = storedSelections[uniqueKey];
                }
                
                // Create display text based on attachment type
                var displayText;
                if (attachment.type === 'google_docs') {
                  var docTypeIcon = getGoogleDocsIcon(attachment.docType);
                  displayText = docTypeIcon + " " + attachment.name + " (Google " + attachment.docType.charAt(0).toUpperCase() + attachment.docType.slice(1) + ")";
                } else {
                  displayText = attachment.name + " (" + formatFileSize(attachment.size) + ")";
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(displayText, originalIndex.toString(), isSelected);
                
                extensionSection.addWidget(checkbox);
              }
              
              card.addSection(extensionSection);
            }
          } else {
          var noAttachSection = CardService.newCardSection()
            .setHeader("No Attachments Found");
          
          var noAttachText = CardService.newTextParagraph()
            .setText("This email thread doesn't contain any attachments.");
          noAttachSection.addWidget(noAttachText);
          card.addSection(noAttachSection);
        }
      } catch (attachmentError) {
        console.error("Error loading attachments:", attachmentError);
      }
      
      // Add main section to card first
      card.addSection(section);
      
      // ===== SUBFOLDER SELECTION & SAVE BUTTON SECTION =====
      var saveSection = CardService.newCardSection()
        .setHeader("📁 Select project sub folder");
      
      // Check if subfolder feature is enabled
      var useEnhancedUI = shouldUseEnhancedUI();
      
      if (useEnhancedUI) {
        // Add subfolder selection dropdown (no title to avoid text overlap)
        var subfolderDropdown = CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("selectedSubfolder")
          .addItem("Project Root (Default)", "", true)
          .addItem("01_System_Design", "01_System_Design", false)
          .addItem("02_Meet_Recordings", "02_Meet_Recordings", false)
          .addItem("03_Correspondence", "03_Correspondence", false)
          .addItem("04_Project_Documentation", "04_Project_Documentation", false)
          .addItem("04_Project_Documentation > Project_Management", "04_Project_Documentation/Project_Management", false)
          .addItem("04_Project_Documentation > Carrier_Onboarding", "04_Project_Documentation/Carrier_Onboarding", false);
        
        saveSection.addWidget(subfolderDropdown);
        
        // Add helpful text with better spacing
        var helperText = CardService.newTextParagraph()
          .setText("<font color=\"#888888\"><i>Folders created automatically if needed</i></font>");
        saveSection.addWidget(helperText);
      }
      
      var saveAction = CardService.newAction()
        .setFunctionName("saveSelectedAttachmentsToGDrive")
        .setParameters({threadId: (e && e.gmail) ? e.gmail.threadId : ""});
      
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("💾 Save to Project Folder")
          .setOnClickAction(saveAction));
      
      saveSection.addWidget(saveButtonSet);
      card.addSection(saveSection);
      
      console.log("Card built successfully");
      return [card.build()];
      
    } catch (error) {
      console.error("Critical error in buildAddOn:", error);
      
      // Fallback card with settings option
      var fallbackCard = CardService.newCardBuilder();
      var fallbackSection = CardService.newCardSection()
        .setHeader("Error - Check Settings");
      
      var errorText = CardService.newTextParagraph()
        .setText("Error: " + error.message);
      fallbackSection.addWidget(errorText);
      
      var settingsButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("⚙️ Configure Settings")
          .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
      fallbackSection.addWidget(settingsButtonSet);
      
      var ticketInput = CardService.newTextInput()
        .setFieldName("jiraTicket")
        .setTitle("Jira Ticket")
        .setHint("e.g., CXPRODELIVERY-6310");
      fallbackSection.addWidget(ticketInput);
      
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("Save to Folder")
          .setOnClickAction(CardService.newAction().setFunctionName("saveSelectedAttachmentsToGDrive")));
      fallbackSection.addWidget(saveButtonSet);
      
      fallbackCard.addSection(fallbackSection);
      return [fallbackCard.build()];
    }
  }
  
  // ===== SETTINGS FUNCTIONS =====
  
  function buildSettingsCard(isFirstTime) {
    console.log("=== BUILD SETTINGS CARD ===");
    
    try {
      var card = CardService.newCardBuilder();
      console.log("Card builder created successfully");
      
      // Get current settings first and validate
      var settings = getUserSettings();
      console.log("Retrieved settings:", settings ? "Success" : "Failed");
    
    // Header section
    var headerSection = CardService.newCardSection()
      .setHeader(isFirstTime ? "⚙️ Initial Setup Required" : "⚙️ Settings");
    
    if (isFirstTime) {
      var setupInfo = CardService.newTextParagraph()
        .setText("Welcome! Please configure your Jira connection to get started.");
      headerSection.addWidget(setupInfo);
    }
    
    card.addSection(headerSection);
    console.log("Header section added successfully");
    
    // Settings section
    console.log("Creating settings section...");
    var settingsSection = CardService.newCardSection()
      .setHeader("Jira Configuration");
    console.log("Settings section created successfully");
    
    // Jira URL input
    console.log("Creating Jira URL input...");
    var jiraUrlInput = CardService.newTextInput()
      .setFieldName("jiraUrl")
      .setTitle("Jira URL")
      .setHint("e.g., https://your-company.atlassian.net")
      .setValue(settings.jiraUrl || getDefaultJiraURL());
    settingsSection.addWidget(jiraUrlInput);
    console.log("Jira URL input added successfully");
    
    // Jira Token input
    var jiraTokenInput = CardService.newTextInput()
      .setFieldName("jiraToken")
      .setTitle("Jira API Token")
      .setHint("Your personal Jira API token")
      .setValue(settings.jiraToken || "");
    settingsSection.addWidget(jiraTokenInput);
    
    // Instructions for getting token
    var tokenHelp = CardService.newTextParagraph()
      .setText("📖 How to get API token:\n1. Go to Jira → Profile → Security\n2. Create API token\n3. Copy and paste above");
    settingsSection.addWidget(tokenHelp);
    
    // JQL input - using regular text input instead of multiline
    var jqlInput = CardService.newTextInput()
      .setFieldName("customJql")
      .setTitle("Custom JQL Query")
      .setHint("Optional: Custom JQL to filter your tickets")
      .setValue(settings.customJql || getDefaultJQL());
    settingsSection.addWidget(jqlInput);
    
    // JQL help
    var jqlHelp = CardService.newTextParagraph()
      .setText("💡 JQL Examples:\n• assignee = currentUser()\n• project = MYPROJECT AND status = 'In Progress'\n• 'Technical Project Manager' in (currentUser())");
    settingsSection.addWidget(jqlHelp);
    
    card.addSection(settingsSection);
    
    console.log("Settings section added to card successfully");
    
    // Action buttons section
    console.log("Creating button section...");
    var buttonSection = CardService.newCardSection();
    console.log("Button section created successfully");
    
    // Save button
    var saveButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("💾 Save Settings")
        .setOnClickAction(CardService.newAction().setFunctionName("saveSettings")));
    buttonSection.addWidget(saveButtonSet);
    
    // Test buttons
    var testButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🧪 Test Jira Connection")
        .setOnClickAction(CardService.newAction().setFunctionName("testJiraConnection")));
    buttonSection.addWidget(testButtonSet);
    
    // Folder Lookup Test Button
    var folderTestButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("📁 Test Folder Lookup")
        .setOnClickAction(CardService.newAction().setFunctionName("testFolderLookup")));
    buttonSection.addWidget(folderTestButtonSet);
    
    // Back button (only if not first time)
    if (!isFirstTime) {
      var backButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("← Back to Main")
          .setOnClickAction(CardService.newAction().setFunctionName("backToMain")));
      buttonSection.addWidget(backButtonSet);
    }
    
    card.addSection(buttonSection);
    
    // Advanced Settings section
    var advancedSection = CardService.newCardSection()
      .setHeader("⚙️ Advanced Settings")
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0);
    
    // Log Level dropdown
    var currentLogLevel = getLogLevel();
    var logLevelDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("logLevel")
      .setTitle("Log Level")
      .addItem("ERROR - Only errors", "0", currentLogLevel === 0)
      .addItem("WARN - Errors & warnings", "1", currentLogLevel === 1)
      .addItem("INFO - Normal (recommended)", "2", currentLogLevel === 2)
      .addItem("DEBUG - Verbose (troubleshooting)", "3", currentLogLevel === 3);
    advancedSection.addWidget(logLevelDropdown);
    
    // Log level help text
    var logLevelHelp = CardService.newTextParagraph()
      .setText("💡 Tip: Use DEBUG level only when troubleshooting.\nLower log levels improve performance.");
    advancedSection.addWidget(logLevelHelp);
    
    // Clear Cache button
    var clearCacheButton = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🗑️ Clear Folder Cache")
        .setOnClickAction(CardService.newAction().setFunctionName("clearAllFolderCache")));
    advancedSection.addWidget(clearCacheButton);
    
    card.addSection(advancedSection);
    
    // Current settings display (if configured)
    if (settings.jiraUrl && settings.jiraToken) {
      var currentSection = CardService.newCardSection()
        .setHeader("Current Configuration")
        .setCollapsible(true)
        .setNumUncollapsibleWidgets(0);
      
      var currentInfo = CardService.newTextParagraph()
        .setText("🔗 URL: " + settings.jiraUrl + "\n🔑 Token: " + maskToken(settings.jiraToken) + "\n📋 JQL: " + (settings.customJql ? "Custom" : "Default"));
      currentSection.addWidget(currentInfo);
      
      card.addSection(currentSection);
    }
    
      console.log("About to build settings card");
      var builtCard = card.build();
      console.log("Settings card built successfully");
      return builtCard;
      
    } catch (error) {
      console.error("Error in buildSettingsCard:", error);
      
      // Return a basic fallback card
      var fallbackCard = CardService.newCardBuilder();
      var fallbackSection = CardService.newCardSection()
        .setHeader("Settings Error");
      
      var errorText = CardService.newTextParagraph()
        .setText("Error loading settings: " + error.message);
      fallbackSection.addWidget(errorText);
      
      fallbackCard.addSection(fallbackSection);
      return fallbackCard.build();
    }
  }
  
  function showSettings(e) {
    console.log("=== SHOW SETTINGS ===");
    
    try {
      // Build a simplified settings card with single section
      var card = CardService.newCardBuilder();
      
      // Get current settings
      var settings = getUserSettings();
      
      // Single section with everything
      var mainSection = CardService.newCardSection()
        .setHeader("⚙️ Jira Settings");
      
      // Jira URL input with current value
      var jiraUrlInput = CardService.newTextInput()
        .setFieldName("jiraUrl")
        .setTitle("Jira URL")
        .setHint("e.g., https://your-company.atlassian.net")
        .setValue(settings.jiraUrl || getDefaultJiraURL());
      mainSection.addWidget(jiraUrlInput);
      
      // Jira Token input with masked value for security
      var maskedToken = "";
      if (settings.jiraToken) {
        // Show only first 4 and last 4 characters with asterisks in between
        if (settings.jiraToken.length > 8) {
          maskedToken = settings.jiraToken.substring(0, 4) + "****" + settings.jiraToken.substring(settings.jiraToken.length - 4);
        } else {
          maskedToken = "****" + settings.jiraToken.substring(settings.jiraToken.length - 2);
        }
      }
      
      var jiraTokenInput = CardService.newTextInput()
        .setFieldName("jiraToken")
        .setTitle("Jira API Token")
        .setHint("Leave blank to keep current token, or enter new token")
        .setValue(maskedToken);
      mainSection.addWidget(jiraTokenInput);
      
      // Instructions
      var tokenHelp = CardService.newTextParagraph()
        .setText("📖 How to get API token:\n1. Go to Jira → Profile → Manage Account → Security\n2. Create API token\n3. Copy the generated token (NOT your password)\n\n🔒 Security: Current token is masked for security. Leave unchanged to keep current token.");
      mainSection.addWidget(tokenHelp);
      
      // JQL input as regular text input
      var jqlInput = CardService.newTextInput()
        .setFieldName("customJql")
        .setTitle("Custom JQL Query")
        .setHint("Optional: Custom JQL to filter your tickets")
        .setValue(settings.customJql || getDefaultJQL());
      mainSection.addWidget(jqlInput);
      
      // JQL help
      var jqlHelp = CardService.newTextParagraph()
        .setText("💡 JQL Examples:\n• assignee = currentUser()\n• project = MYPROJECT AND status = 'In Progress'\n• 'Technical Project Manager' in (currentUser())");
      mainSection.addWidget(jqlHelp);
      
      
      // Save button
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("💾 Save Settings")
          .setOnClickAction(CardService.newAction().setFunctionName("saveSettings")));
      mainSection.addWidget(saveButtonSet);
      
      // Test buttons
      var testButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("🧪 Test Jira Connection")
          .setOnClickAction(CardService.newAction().setFunctionName("testJiraConnection")));
      mainSection.addWidget(testButtonSet);
      
      // Folder Lookup Test button
      var folderTestButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("📁 Test Folder Lookup")
          .setOnClickAction(CardService.newAction().setFunctionName("testFolderLookup")));
      mainSection.addWidget(folderTestButtonSet);
      
      // Back button
      var backButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("← Back to Main")
          .setOnClickAction(CardService.newAction().setFunctionName("backToMain")));
      mainSection.addWidget(backButtonSet);
      
      card.addSection(mainSection);
      
      var builtCard = card.build();
      console.log("Settings card built successfully");
      
      // Try with updateCard instead of pushCard
      var navigation = CardService.newNavigation().updateCard(builtCard);
      console.log("Navigation created successfully");
      
      var response = CardService.newActionResponseBuilder()
        .setNavigation(navigation)
        .build();
      
      console.log("Action response built successfully");
      return response;
      
    } catch (error) {
      console.error("Error in showSettings:", error);
      
      // Return a simple notification instead of navigation
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("Error opening settings: " + error.message))
        .build();
    }
  }
  
  function saveSettings(e) {
    console.log("=== SAVE SETTINGS ===");
    console.log("Form input:", JSON.stringify(e.formInput));
    
    try {
      var jiraUrl = e.formInput.jiraUrl;
      var jiraToken = e.formInput.jiraToken;
      var customJql = e.formInput.customJql;
      
      // Get existing settings to check for current token
      var existingSettings = getUserSettings();
      
      // Handle masked token - if user didn't change it, keep the existing token
      var finalToken = jiraToken;
      if (existingSettings.jiraToken && jiraToken && jiraToken.includes("****")) {
        // User left the masked token unchanged, keep the existing token
        finalToken = existingSettings.jiraToken;
      }
      
      if (!jiraUrl || (!finalToken && !existingSettings.jiraToken)) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please provide both Jira URL and API token"))
          .build();
      }
      
      // Clean up URL (remove trailing slash)
      if (jiraUrl.endsWith('/')) {
        jiraUrl = jiraUrl.slice(0, -1);
      }

      // Validate URL format
      if (!jiraUrl.startsWith('http')) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please provide a valid URL starting with http"))
          .build();
      }
      
      
      // Save settings
      var settings = {
        jiraUrl: jiraUrl,
        jiraToken: finalToken,
        customJql: customJql || getDefaultJQL(),
        savedAt: new Date().toISOString()
      };
      
      setUserSettings(settings);
      
      // Save log level if provided
      var logLevel = e.formInput.logLevel;
      if (logLevel !== undefined && logLevel !== null) {
        setLogLevel(parseInt(logLevel, 10));
      }
      
      console.log("Settings saved successfully");
      
      // Success notification for Jira settings
      var currentLogLevelName = getLogLevelName(getLogLevel());
      var successMsg = "✅ Settings Saved Successfully!\n";
      successMsg += "• Jira URL: " + jiraUrl + "\n";
      successMsg += "• API Token: " + (finalToken ? "Configured" : "Not Set") + "\n";
      successMsg += "• JQL Query: " + (customJql ? "Custom" : "Default") + "\n";
      successMsg += "• Log Level: " + currentLogLevelName + "\n\n";
      successMsg += "📁 Folder Integration: Using Jira summary parsing\n";
      successMsg += "💡 Use 'Test Connections' to verify functionality";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(successMsg))
        .build();
        
    } catch (error) {
      console.error("Error saving settings:", error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Error saving settings: " + error.message))
        .build();
    }
  }
  
  function testJiraConnection(e) {
    console.log("=== TEST JIRA CONNECTION ===");
    
    try {
      var settings = getUserSettings();
      
      if (!settings.jiraUrl || !settings.jiraToken) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please save your settings first"))
          .build();
      }
      
      // Test the connection
      var testResult = testJiraAPI(settings);
      
      if (testResult.success) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("✅ Connection successful! Logged in as: " + testResult.userInfo))
          .build();
      } else {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Connection failed: " + testResult.error))
          .build();
      }
      
    } catch (error) {
      console.error("Error testing connection:", error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Test failed: " + error.message))
        .build();
    }
  }
  
  /**
   * Test folder lookup functionality using Jira summary parsing
   */
  function testFolderLookup(e) {
    console.log("=== TEST FOLDER LOOKUP ===");
    
    try {
      var results = [];
      
      // Test 1: Check Customers folder access
      console.log("Testing Customers folder access...");
      try {
        var customersFolder = DriveApp.getFolderById(getCustomersFolderId());
        results.push("✅ Customers Folder: Accessible");
        results.push("   Name: " + customersFolder.getName());
        results.push("   ID: " + getCustomersFolderId());
      } catch (folderError) {
        results.push("❌ Customers Folder: " + folderError.message);
        
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Folder Lookup Test FAILED\n\n" + results.join("\n") + "\n\n💡 Check folder permissions or ID"))
          .build();
      }
      
      // Test 2: Test Jira connection and summary parsing
      console.log("Testing Jira summary parsing...");
      var testTicket = "CXPRODELIVERY-4750"; // Known test ticket
      
      try {
        var parseResult = getJiraTicketFolderNames(testTicket);
        
        if (parseResult.success) {
          results.push("");
          results.push("✅ Jira Summary Parsing: Success");
          results.push("   Customer: " + parseResult.parsed.customerId + " - " + parseResult.parsed.customerName);
          results.push("   Project Type: " + parseResult.parsed.projectType);
          results.push("");
          results.push("📁 Generated Folder Names:");
          results.push("   Customer: " + parseResult.folderNames.customerFolder);
          results.push("   Project: " + parseResult.folderNames.projectFolder);
        } else {
          results.push("⚠️ Jira Summary Parsing: " + parseResult.error);
        }
      } catch (parseError) {
        results.push("⚠️ Jira Summary Parsing: " + parseError.message);
      }
      
      // Test 3: Test folder search (preview only)
      console.log("Testing folder search...");
      try {
        var workflowResult = completeJiraToFolderWorkflowWithCreate(testTicket, false);
        
        if (workflowResult.success) {
          results.push("");
          results.push("✅ Folder Search Complete:");
          
          if (workflowResult.customerFolder && workflowResult.customerFolder.exists) {
            results.push("   📁 Customer: EXISTS - " + workflowResult.customerFolder.name);
          } else if (workflowResult.customerFolder) {
            results.push("   📁 Customer: NOT FOUND (would create: " + workflowResult.customerFolder.suggestedName + ")");
          }
          
          if (workflowResult.projectFolder && workflowResult.projectFolder.exists) {
            results.push("   📁 Project: EXISTS - " + workflowResult.projectFolder.name);
          } else if (workflowResult.projectFolder) {
            results.push("   📁 Project: NOT FOUND (would create: " + workflowResult.projectFolder.suggestedName + ")");
          }
        } else {
          results.push("⚠️ Folder Search: " + workflowResult.error);
        }
      } catch (searchError) {
        results.push("⚠️ Folder Search: " + searchError.message);
      }
      
      var successMsg = "✅ Folder Lookup Test Complete!\n\n" + results.join("\n");
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(successMsg))
        .build();
        
    } catch (error) {
      console.error("❌ Folder lookup test failed:", error.message);
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Folder Lookup Test FAILED\n• Error: " + error.message))
        .build();
    }
  }
  
  function backToMain(e) {
    console.log("=== BACK TO MAIN ===");
    
    try {
      var cards = buildAddOn(e);
      
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(cards[0]))
        .build();
        
    } catch (error) {
      console.error("Error in backToMain:", error);
      
      // Fallback: just pop the current card
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popCard())
        .build();
    }
  }
  
  // ===== DYNAMIC TICKET DETAILS FUNCTION =====
  
  function showTicketDetails(e) {
    console.log("=== SHOW TICKET DETAILS CALLED ===");
    console.log("Form input:", JSON.stringify(e.formInput));
    console.log("Parameters:", JSON.stringify(e.parameters));
    
    try {
      // Get threadId first (needed for dropdown onChange actions)
      var threadId = "";
      if (e.parameters && e.parameters.threadId) {
        threadId = e.parameters.threadId;
        console.log("Got threadId from parameters:", threadId);
      } else if (e.gmail && e.gmail.threadId) {
        threadId = e.gmail.threadId;
        console.log("Got threadId from gmail:", threadId);
      } else {
        console.log("No threadId found in parameters or gmail context");
      }
      
  
      
      var selectedTicketKey = null;
      
      // Try to get selected ticket from form input
      if (e.formInput && e.formInput.selectedTicket) {
        selectedTicketKey = e.formInput.selectedTicket;
        console.log("Got ticket from formInput:", selectedTicketKey);
      } else if (e.parameters && e.parameters.selectedTicket) {
        selectedTicketKey = e.parameters.selectedTicket;
        console.log("Got ticket from parameters:", selectedTicketKey);
      } else {
        console.log("No selected ticket found in formInput or parameters");
      }
      
      console.log("Final selected ticket key:", selectedTicketKey);
      
      if (!selectedTicketKey || selectedTicketKey === "manual") {
        console.log("Manual entry selected or no ticket selected, rebuilding main card");
        var cards = buildAddOn(e);
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().updateCard(cards[0]))
          .build();
      }
      
      // Get the selected ticket details
      console.log("Getting Jira projects...");
      var projectResult = getMyJiraProjects();
      console.log("Got", projectResult.issues ? projectResult.issues.length : 0, "issues");
      var selectedTicket = null;
      
      for (var i = 0; i < projectResult.issues.length; i++) {
        if (projectResult.issues[i].key === selectedTicketKey) {
          selectedTicket = projectResult.issues[i];
          break;
        }
      }
      
      if (!selectedTicket) {
        console.log("Selected ticket not found:", selectedTicketKey);
        console.log("Available tickets:", projectResult.issues.map(function(issue) { return issue.key; }));
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().popCard())
          .build();
      }
      
      console.log("Found selected ticket:", selectedTicket.key);
      
      // Build card with ticket details
      console.log("Building card...");
      var card = CardService.newCardBuilder();
      
      // Compact header with just settings button
      console.log("Creating header section...");
      var headerSection = CardService.newCardSection();
      
      // Compact settings button (gear only)
      var settingsButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("⚙️")
          .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
      headerSection.addWidget(settingsButtonSet);
      
      card.addSection(headerSection);
      console.log("Header section added");
      
      // Main section with dropdown
      console.log("Creating main section...");
      var mainSection = CardService.newCardSection();
      
      // Recreate dropdown with current selection
      console.log("Creating ticket dropdown...");
      var ticketDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName("selectedTicket")
        .setTitle("Select Active Ticket")
        .setOnChangeAction(CardService.newAction()
          .setFunctionName("showTicketDetails")
          .setParameters({threadId: threadId}));
      
      // Add manual entry option
      console.log("Adding manual entry option...");
      ticketDropdown.addItem("Manual Entry - Enter ticket number below", "manual", false);
      
      // Add all tickets with current one selected
      console.log("Sorting and adding", projectResult.issues.length, "tickets to dropdown...");
      projectResult.issues.sort(function(a, b) {
        return a.key.localeCompare(b.key);
      });
      
      for (var i = 0; i < projectResult.issues.length; i++) {
        try {
          var issue = projectResult.issues[i];
          console.log("Processing ticket", i + 1, "of", projectResult.issues.length, ":", issue.key);
          var displayText = formatCompactTicketDisplay(issue);
          var isSelected = (issue.key === selectedTicketKey);
          ticketDropdown.addItem(displayText, issue.key, isSelected);
          console.log("Successfully added ticket:", issue.key);
        } catch (itemError) {
          console.error("Error processing ticket", issue ? issue.key : "unknown", ":", itemError);
          // Continue with other tickets
        }
      }
      console.log("Finished adding tickets to dropdown");
      
      console.log("Adding dropdown to main section...");
      mainSection.addWidget(ticketDropdown);
      console.log("Dropdown added to main section");
      
  
      
      // Add manual ticket input (pre-filled with selected ticket number)
      console.log("Adding manual ticket input...");
      var ticketNumber = selectedTicketKey.replace('CXPRODELIVERY-', '');
      var ticketInput = CardService.newTextInput()
        .setFieldName("manualTicketNumber")
        .setTitle("Manual Ticket Entry")
        .setHint("e.g., 6500 (will become CXPRODELIVERY-6500)")
        .setValue(ticketNumber);
      mainSection.addWidget(ticketInput);
      console.log("Manual ticket input added");
      
      // Create ticket details section
      console.log("Creating ticket details section...");
      var detailsSection = CardService.newCardSection()
        .setHeader("📋 Selected Ticket Details");
      
      // Get Jira URL for ticket link
      var settings = getUserSettings();
      var jiraTicketUrl = settings.jiraUrl + '/browse/' + selectedTicket.key;
      
      // Add clickable ticket link button
      var ticketLinkButton = CardService.newTextButton()
        .setText("🔗 " + selectedTicket.key)
        .setOpenLink(CardService.newOpenLink()
          .setUrl(jiraTicketUrl)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
          .setOnClose(CardService.OnClose.NOTHING));
      detailsSection.addWidget(ticketLinkButton);
      
      // Add ticket details as text
      var statusEmoji = getStatusEmoji(selectedTicket.status);
      var detailsText = "\n📝 Summary:\n" + selectedTicket.summary + "\n\n";
      detailsText += "📊 Status: " + statusEmoji + " " + selectedTicket.status + "\n";
      detailsText += "🏷️ Type: " + selectedTicket.issueType;
      
      var detailsParagraph = CardService.newTextParagraph().setText(detailsText);
      detailsSection.addWidget(detailsParagraph);
      console.log("Ticket details section created with clickable link");
      
      // Add customer folder link to details
      console.log("Looking up customer folder for ticket:", selectedTicket.key);
      try {
        // Search for existing customer folder (don't create if missing)
        var folderResult = completeJiraToFolderWorkflowWithCreate(selectedTicket.key, false);
        
        if (folderResult.success && folderResult.customerFolder && folderResult.customerFolder.exists) {
          console.log("✅ Customer folder found:", folderResult.customerFolder.name);
          console.log("   URL:", folderResult.customerFolder.url);
          
          // Add visual separator
          detailsSection.addWidget(CardService.newTextParagraph()
            .setText("\n─────────────────────"));
          
          // Add customer folder as a clickable button
          var customerFolderButton = CardService.newTextButton()
            .setText("📁 Open Customer Folder")
            .setOpenLink(CardService.newOpenLink()
              .setUrl(folderResult.customerFolder.url)
              .setOpenAs(CardService.OpenAs.FULL_SIZE)
              .setOnClose(CardService.OnClose.NOTHING));
          
          // Add label and button
          detailsSection.addWidget(CardService.newTextParagraph()
            .setText("Customer Folder:\n" + folderResult.customerFolder.name));
          detailsSection.addWidget(customerFolderButton);
          
          console.log("✅ Customer folder link added to details");
        } else {
          console.log("⚠️ Customer folder not found (will be created on first upload)");
          
          // Show expected folder name
          if (folderResult.customerFolder && folderResult.customerFolder.suggestedName) {
            detailsSection.addWidget(CardService.newTextParagraph()
              .setText("\n─────────────────────\nCustomer Folder:\n_" + folderResult.customerFolder.suggestedName + "_\n_(will be created on first upload)_"));
          }
        }
      } catch (folderError) {
        console.error("❌ Error looking up customer folder:", folderError.message);
        // Don't fail the whole card if folder lookup fails
      }
      
      // Handle Gmail context for attachments (threadId already determined above)
      
      if (threadId) {
        console.log("Processing attachments for threadId:", threadId);
        try {
          var thread = GmailApp.getThreadById(threadId);
          var messages = thread.getMessages();
          var allAttachments = [];
          var allGoogleDocsLinks = [];
          console.log("Found", messages.length, "messages in thread");
          
          for (var i = 0; i < messages.length; i++) {
            var message = messages[i];
            var attachments = message.getAttachments();
            console.log("Message", i + 1, "has", attachments.length, "attachments");
            
            // Process regular attachments
            for (var j = 0; j < attachments.length; j++) {
              var attachment = attachments[j];
              allAttachments.push({
                name: attachment.getName(),
                size: attachment.getSize(),
                attachment: attachment,
                messageIndex: i,
                type: 'regular'
              });
            }
            
            // Extract Google Docs links from message content
            var googleDocsLinks = extractGoogleDocsLinks(message);
            console.log("Message", i + 1, "has", googleDocsLinks.length, "Google Docs links");
            for (var k = 0; k < googleDocsLinks.length; k++) {
              var docLink = googleDocsLinks[k];
              allGoogleDocsLinks.push({
                name: docLink.title,
                size: 0, // Google Docs links don't have a file size
                url: docLink.url,
                docType: docLink.type,
                messageIndex: i,
                type: 'google_docs'
              });
            }
          }
          
          // Combine regular attachments and Google Docs links
          var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
          
          console.log("Total attachments and Google Docs links found:", combinedAttachments.length);
          
          if (combinedAttachments.length > 0) {
            console.log("Creating attachment section...");
            
            // First, check if we have current form selections to preserve
            var currentSelections = {};
            var hasCurrentSelections = false;
            
            if (e.formInput) {
              console.log("=== CHECKING CURRENT FORM SELECTIONS ===");
              for (var i = 0; i < combinedAttachments.length; i++) {
                var formFieldName = "attachment_" + i;
                // If field exists in form input → selected = true, if missing → selected = false
                var isCurrentlySelected = e.formInput.hasOwnProperty(formFieldName) && e.formInput[formFieldName] && e.formInput[formFieldName].length > 0;
                
                // Use unique key to handle duplicate filenames (filename + index)
                var uniqueKey = combinedAttachments[i].name + "_" + i;
                currentSelections[uniqueKey] = isCurrentlySelected;
                hasCurrentSelections = true;
                console.log("Current form state for", combinedAttachments[i].name, "- field:", formFieldName, "- exists:", e.formInput.hasOwnProperty(formFieldName), "- selected:", isCurrentlySelected, "- stored as:", uniqueKey);
              }
            }
            
            // If we have current selections, store them for future use
            if (hasCurrentSelections) {
              console.log("Storing current form selections for future use");
              storeAttachmentSelections(threadId, currentSelections);
            }
            
            // Get stored attachment selections for this thread (might be what we just stored)
            console.log("Getting stored attachment selections...");
            var storedSelections = getStoredAttachmentSelections(threadId);
            console.log("Stored selections:", storedSelections);
            
            // Group attachments by file extension
            var attachmentsByExtension = {};
            var extensionCounts = {};
            
            for (var i = 0; i < combinedAttachments.length; i++) {
              var attachment = combinedAttachments[i];
              var fileName = attachment.name;
              
              // Determine extension/category
              var extension;
              if (attachment.type === 'google_docs') {
                // Group Google Docs by their type (docs, sheets, slides, etc.)
                extension = 'google-' + attachment.docType;
              } else {
                // Regular file extension
                extension = fileName.lastIndexOf('.') > -1 ? 
                fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              }
              
              if (!attachmentsByExtension[extension]) {
                attachmentsByExtension[extension] = [];
                extensionCounts[extension] = 0;
              }
              
              attachmentsByExtension[extension].push({
                attachment: attachment,
                index: i
              });
              extensionCounts[extension]++;
            }
            
            // Sort extensions alphabetically for consistent display
            var sortedExtensions = Object.keys(attachmentsByExtension).sort();
            console.log("Attachment extensions found:", sortedExtensions);
            
            // Create a section for each file extension
            for (var extIndex = 0; extIndex < sortedExtensions.length; extIndex++) {
              var extension = sortedExtensions[extIndex];
              var attachmentsInGroup = attachmentsByExtension[extension];
              var count = extensionCounts[extension];
              
              console.log("Processing extension group:", extension, "with", count, "files");
              
              // Create section header with extension and count (show total in first section)
              var extDisplayName = extension === 'no-extension' ? 'Files without extension' : 
                extension.toUpperCase() + ' files';
              var headerText = extDisplayName + " (" + count + ")";
              if (extIndex === 0) {
                headerText = "📎 Select Attachments & Docs: " + headerText + " (Total: " + combinedAttachments.length + ")";
              }
              
              var extensionSection = CardService.newCardSection()
                .setHeader(headerText)
                .setCollapsible(true)
                .setNumUncollapsibleWidgets(count > 5 ? 2 : count); // Show first 2 if more than 5 files
              
              // Add checkboxes for each attachment in this extension group
              for (var j = 0; j < attachmentsInGroup.length; j++) {
                var attachmentData = attachmentsInGroup[j];
                var attachment = attachmentData.attachment;
                var originalIndex = attachmentData.index;
                
                console.log("Processing attachment", originalIndex, ":", attachment.name);
                
                // Check if this attachment was previously selected
                var isSelected = false; // Default to false - let user explicitly choose
                
                // Use unique key to handle duplicate filenames (filename + index)
                var uniqueKey = attachment.name + "_" + originalIndex;
                
                if (storedSelections && storedSelections.hasOwnProperty(uniqueKey)) {
                  isSelected = storedSelections[uniqueKey];
                  console.log("Found stored selection for", attachment.name, "(", uniqueKey, "):", isSelected);
                } else {
                  console.log("No stored selection found for", attachment.name, "(", uniqueKey, ") - using default (unselected):", isSelected);
                }
                
                // Create display text based on attachment type
                var displayText;
                if (attachment.type === 'google_docs') {
                  var docTypeIcon = getGoogleDocsIcon(attachment.docType);
                  displayText = docTypeIcon + " " + attachment.name + " (Google " + attachment.docType.charAt(0).toUpperCase() + attachment.docType.slice(1) + ")";
                } else {
                  displayText = attachment.name + " (" + formatFileSize(attachment.size) + ")";
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(displayText, originalIndex.toString(), isSelected);
                
                extensionSection.addWidget(checkbox);
              }
              
              console.log("Adding", extension, "section to card");
              card.addSection(extensionSection);
            }
            
            console.log("All attachment sections added to card");
          } else {
            console.log("No attachments found, skipping attachment section");
          }
        } catch (attachmentError) {
          console.error("Error loading attachments in showTicketDetails:", attachmentError);
        }
      } else {
        console.log("No threadId, skipping attachment processing");
      }
      
      // Add sections to card in the correct order
      console.log("Adding main section to card...");
      card.addSection(mainSection);
      
      // ===== SUBFOLDER SELECTION & SAVE BUTTON =====
      // Check if subfolder feature is enabled
      var useEnhancedUI = shouldUseEnhancedUI();
      
      if (useEnhancedUI) {
        // Create separate section for subfolder selection
        var subfolderSection = CardService.newCardSection()
          .setHeader("📁 Select project sub folder");
        
        // Dropdown without title to avoid text overlap
        var subfolderDropdown = CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("selectedSubfolder")
          .addItem("Project Root (Default)", "", true)
          .addItem("01_System_Design", "01_System_Design", false)
          .addItem("02_Meet_Recordings", "02_Meet_Recordings", false)
          .addItem("03_Correspondence", "03_Correspondence", false)
          .addItem("04_Project_Documentation", "04_Project_Documentation", false)
          .addItem("04_Project_Documentation > Project_Management", "04_Project_Documentation/Project_Management", false)
          .addItem("04_Project_Documentation > Carrier_Onboarding", "04_Project_Documentation/Carrier_Onboarding", false);
        
        subfolderSection.addWidget(subfolderDropdown);
        
        // Add helpful text with better spacing
        var helperText = CardService.newTextParagraph()
          .setText("<font color=\"#888888\"><i>Folders created automatically if needed</i></font>");
        subfolderSection.addWidget(helperText);
        
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("💾 Save to Project Folder")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("saveSelectedAttachmentsToGDrive")
              .setParameters({threadId: threadId})));
        
        subfolderSection.addWidget(saveButtonSet);
        console.log("Adding save attachments section to card...");
        card.addSection(subfolderSection);
      } else {
        // Legacy UI - simple save button
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Save to Project Folder")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("saveSelectedAttachmentsToGDrive")
              .setParameters({threadId: threadId})));
        mainSection.addWidget(saveButtonSet);
      }
      
      // Add details section LAST (at the very bottom)
      console.log("Adding details section to card (at bottom)...");
      card.addSection(detailsSection);
      
      console.log("All sections added. Built card with ticket details for:", selectedTicket.key);
      
      var builtCard = card.build();
      console.log("Card built successfully");
      
      var navigation = CardService.newNavigation().updateCard(builtCard);
      console.log("Navigation created");
      
      var response = CardService.newActionResponseBuilder()
        .setNavigation(navigation)
        .build();
      console.log("Response built, returning...");
      
      return response;
        
    } catch (error) {
      console.error("Error in showTicketDetails:", error);
      console.error("Error stack:", error.stack);
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popCard())
        .build();
    }
  }
  
  // ===== SAVE FUNCTION =====
  
  function saveSelectedAttachmentsToGDrive(e) {
    try {
      console.log("=== SAVE FUNCTION CALLED ===");
      console.log("Save operation timestamp:", new Date().toISOString());
      console.log("Form input:", JSON.stringify(e.formInput));
      console.log("Parameters:", JSON.stringify(e.parameters));
      
      // Add comprehensive diagnostic information for troubleshooting (only if verbose logging enabled)
      if (ENABLE_VERBOSE_LOGGING) {
        console.log("\n=== DIAGNOSTIC INFO FOR TROUBLESHOOTING ===");
        logDiagnosticInfo();
      }
      
      var selectedTicket = e.formInput.selectedTicket;
      var manualTicketNumber = e.formInput.manualTicketNumber;
      var fullTicket = e.formInput.jiraTicket;
      var threadId = e.parameters ? e.parameters.threadId : "";
      
      // Determine final ticket
      var finalTicket;
      if (selectedTicket && selectedTicket !== "manual") {
        finalTicket = selectedTicket;
        console.log("Using selected ticket:", finalTicket);
      } else if (manualTicketNumber) {
        finalTicket = "CXPRODELIVERY-" + manualTicketNumber;
        console.log("Using manual ticket:", finalTicket);
      } else if (fullTicket) {
        finalTicket = fullTicket;
        console.log("Using full ticket:", finalTicket);
      } else {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please select a ticket or enter a manual ticket number"))
          .build();
      }
      
      if (!threadId) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please open an email to save attachments"))
          .build();
      }
      
      console.log("Loading attachments from thread:", threadId);
      
      // Get all attachments and Google Docs links from the thread
      var thread = GmailApp.getThreadById(threadId);
      var messages = thread.getMessages();
      var allAttachments = [];
      var allGoogleDocsLinks = [];
      
      for (var i = 0; i < messages.length; i++) {
        var message = messages[i];
        var attachments = message.getAttachments();
        
        // Process regular attachments
        for (var j = 0; j < attachments.length; j++) {
          var attachment = attachments[j];
          allAttachments.push({
            name: attachment.getName(),
            size: attachment.getSize(),
            attachment: attachment,
            messageIndex: i,
            type: 'regular'
          });
        }
        
        // Extract Google Docs links from message content
        var googleDocsLinks = extractGoogleDocsLinks(message);
        var emailSubject = message.getSubject();
        for (var k = 0; k < googleDocsLinks.length; k++) {
          var docLink = googleDocsLinks[k];
          allGoogleDocsLinks.push({
            name: docLink.title,
            size: 0, // Google Docs links don't have a file size
            url: docLink.url,
            docType: docLink.type,
            messageIndex: i,
            type: 'google_docs',
            emailSubject: emailSubject
          });
        }
      }
      
      // Combine regular attachments and Google Docs links
      var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
      
      console.log("Found", combinedAttachments.length, "total attachments and Google Docs links");
      
      if (combinedAttachments.length === 0) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("No attachments found in this thread"))
          .build();
      }
      
      // Find selected attachments from form input and store selection state
      var selectedAttachments = [];
      var selectionState = {};
      
      console.log("=== PROCESSING ATTACHMENT SELECTIONS FOR SAVE ===");
      console.log("ThreadId for storage:", threadId);
      console.log("Total attachments to process:", combinedAttachments.length);
      console.log("Form input keys:", Object.keys(e.formInput));
      
      for (var i = 0; i < combinedAttachments.length; i++) {
        var checkboxValue = e.formInput["attachment_" + i];
        var isSelected = checkboxValue && checkboxValue.length > 0;
        var attachmentName = combinedAttachments[i].name;
        var attachmentType = combinedAttachments[i].type || 'regular';
        
        console.log("Attachment", i, "(", attachmentName, ") - type:", attachmentType, "- checkbox value:", checkboxValue, "- selected:", isSelected);
        
        // Store selection state for this attachment using unique key
        var uniqueKey = attachmentName + "_" + i;
        selectionState[uniqueKey] = isSelected;
        
        if (isSelected) {
          selectedAttachments.push(combinedAttachments[i]);
          console.log("Added to selected list:", attachmentName, "(" + attachmentType + ")");
        }
      }
      
      console.log("Final selection state to store:", JSON.stringify(selectionState));
      
      // Store the current selection state for this thread
      storeAttachmentSelections(threadId, selectionState);
      
      console.log("Selected", selectedAttachments.length, "attachments for saving");
      
      if (selectedAttachments.length === 0) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please select at least one attachment"))
          .build();
      }
      
  // Jira Summary Parsing - Find/Create folder from ticket summary (OPTIMIZED)
      console.log("\n🔍 Finding project folder for " + finalTicket + "...");
      
      // Load settings once and pass to workflow for better performance
      var cachedSettings = getUserSettings();
      var folderWorkflowResult = completeJiraToFolderWorkflowWithCreate(finalTicket, true, cachedSettings);
      
      verboseLog("📋 Folder Workflow Result: " + (folderWorkflowResult.success ? "Success" : "Failed"));
      
      if (!folderWorkflowResult.success) {
        // Handle folder lookup/creation errors
        console.error("\n❌ STEP 1 FAILED: Folder Lookup/Creation");
        console.error("  Error:", folderWorkflowResult.error);
        console.error("  Ticket:", finalTicket);
        console.error("\n⏹️ PROCESS STOPPED: Cannot proceed without project folder");
        
        var userFriendlyError = "❌ Cannot find/create project folder for " + finalTicket + "\n\n";
        userFriendlyError += "Error: " + (folderWorkflowResult.error || "Unknown error") + "\n\n";
        userFriendlyError += "💡 Possible solutions:\n";
        userFriendlyError += "• Check if Jira ticket summary follows expected format\n";
        userFriendlyError += "• Verify you have access to Customers folder\n";
        userFriendlyError += "• Contact your project manager for assistance";
        
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText(userFriendlyError))
          .build();
      }
      
      // Check if project folder was found/created
      if (!folderWorkflowResult.projectFolder || !folderWorkflowResult.projectFolder.exists) {
        console.error("\n❌ STEP 1 FAILED: Project Folder Not Available");
        
        var errorMsg = "❌ Project folder not available for " + finalTicket + "\n\n";
        if (folderWorkflowResult.projectFolder && folderWorkflowResult.projectFolder.suggestedName) {
          errorMsg += "Suggested name: " + folderWorkflowResult.projectFolder.suggestedName + "\n\n";
        }
        errorMsg += "💡 Please contact your project manager";
        
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText(errorMsg))
          .build();
      }
      
      console.log("✅ Project folder found: " + folderWorkflowResult.projectFolder.name);
      
      // Verify folder access
      var folderResult = getFolderByIdSafely(folderWorkflowResult.projectFolder.id);
      
      verboseLog("📋 Folder Access: " + (folderResult.success ? "Granted" : "Denied"));
      
      if (!folderResult.success) {
        // Handle folder access error
        console.error("\n❌ STEP 2 FAILED: Folder Access Denied");
        console.error("  Folder ID:", folderWorkflowResult.projectFolder.id);
        console.error("  Error:", folderResult.error);
        console.error("\n⏹️ PROCESS STOPPED: Cannot save files without folder access");
        
        // Build and show request access card
        var requestAccessCard = buildRequestAccessCard(finalTicket, folderWorkflowResult.projectFolder.id, folderResult.error);
        
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().pushCard(requestAccessCard))
          .build();
      }
      
      var projectRootFolder = folderResult.folder;
      
      // Subfolder selection
      var selectedSubfolder = e.formInput.selectedSubfolder || "";
      var validation = validateSubfolderSelection(selectedSubfolder);
      if (!validation.valid) {
        selectedSubfolder = validation.sanitized;
      }
      
      // Get or create target subfolder
      var targetFolderResult = getOrCreateProjectSubfolder(projectRootFolder, selectedSubfolder);
      verboseLog("📂 Subfolder: " + (targetFolderResult.success ? targetFolderResult.path : "ERROR"));
      
      if (!targetFolderResult.success) {
        // Attempt fallback strategy
        var fallbackResult = handleSubfolderCreationFailure(projectRootFolder, selectedSubfolder, targetFolderResult);
        
        if (!fallbackResult.success) {
          var subfolderError = getSubfolderCreationErrorMessage(targetFolderResult.error, projectRootFolder, selectedSubfolder);
          return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(subfolderError))
            .build();
        }
        
        console.log("⚠️ Using fallback: " + fallbackResult.actualPath);
        targetFolderResult = fallbackResult;
      }
      
      var ticketFolder = targetFolderResult.folder;
      var actualFolderPath = targetFolderResult.path;
      
      // Save selected attachments
      console.log("💾 Saving " + selectedAttachments.length + " file(s) to: " + actualFolderPath);
      
      var savedCount = 0;
      var skippedCount = 0;
      var savedFiles = [];
      var skippedFiles = [];
      
      for (var i = 0; i < selectedAttachments.length; i++) {
        try {
          var attachmentData = selectedAttachments[i];
          var fileName = attachmentData.name;
          var attachmentType = attachmentData.type || 'regular';
          
          verboseLog("📎 " + (i + 1) + "/" + selectedAttachments.length + ": " + fileName + " (" + attachmentType + ")");
          
          var file;
          var finalFileName;
          
          if (attachmentType === 'google_docs') {
            // Handle Google Docs links
            file = copyGoogleDocToFolder({
              title: attachmentData.name,
              url: attachmentData.url,
              type: attachmentData.docType,
              emailSubject: attachmentData.emailSubject
            }, ticketFolder);
            finalFileName = file.getName();
            savedCount++;
            savedFiles.push(finalFileName);
            
          } else {
            // Handle regular attachments - check for duplicates
            var existingFiles = ticketFolder.getFilesByName(fileName);
            
            if (existingFiles.hasNext()) {
              var existingFile = existingFiles.next();
              var existingSize = existingFile.getSize();
              var newSize = attachmentData.size;
              
              if (existingSize === newSize) {
                // Same size, likely duplicate - skip
                skippedCount++;
                skippedFiles.push(fileName);
                continue;
              } else {
                // Different size, create with timestamp
                var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
                var fileExtension = fileName.lastIndexOf('.') > -1 ? fileName.substring(fileName.lastIndexOf('.')) : '';
                var baseName = fileName.lastIndexOf('.') > -1 ? fileName.substring(0, fileName.lastIndexOf('.')) : fileName;
                var newFileName = baseName + "_" + timestamp + fileExtension;
                
                var saveResult = saveAttachmentWithAdvancedDrive(attachmentData.attachment, ticketFolder.getId(), newFileName);
                
                if (saveResult.success) {
                  file = DriveApp.getFileById(saveResult.fileId);
                  finalFileName = newFileName;
                  savedCount++;
                  savedFiles.push(finalFileName);
                } else {
                  throw new Error("Failed to save file: " + saveResult.error);
                }
              }
            } else {
              // File doesn't exist, save normally
              var saveResult = saveAttachmentWithAdvancedDrive(attachmentData.attachment, ticketFolder.getId(), fileName);
              
              if (saveResult.success) {
                file = DriveApp.getFileById(saveResult.fileId);
                finalFileName = file.getName();
                savedCount++;
                savedFiles.push(finalFileName);
              } else {
                throw new Error("Failed to save file: " + saveResult.error);
              }
            }
          }
        } catch (saveError) {
          console.error("❌ Failed to save " + attachmentData.name + ": " + saveError.message);
        }
      }
      
      // Concise summary (verbose details only if enabled)
      console.log("✅ SAVE COMPLETE: " + savedCount + "/" + selectedAttachments.length + " files saved to " + actualFolderPath);
      if (skippedCount > 0) {
        console.log("   ⏭️ Skipped " + skippedCount + " duplicate(s)");
      }
      
      if (ENABLE_VERBOSE_LOGGING && savedFiles.length > 0) {
        console.log("   Saved: " + savedFiles.slice(0, 3).join(", ") + (savedFiles.length > 3 ? " +" + (savedFiles.length - 3) + " more" : ""));
      }
      
      // Enhanced notification message with detailed folder info and subfolder path
      var notificationText = "✅ Attachment Save Complete!\n\n";
      
      // Project folder information
      if (folderWorkflowResult.created && folderWorkflowResult.created.projectFolder) {
        notificationText += "📁 New project folder created: " + folderResult.name + "\n";
      } else {
        notificationText += "📁 Project folder: " + folderResult.name + "\n";
      }
      
      // Subfolder information
      if (actualFolderPath && actualFolderPath !== "Project Root") {
        notificationText += "📂 Subfolder: " + actualFolderPath + "\n";
        if (targetFolderResult.created) {
          notificationText += "🆕 Created new subfolder(s): " + targetFolderResult.createdFolders.join(", ") + "\n";
        }
        if (targetFolderResult.fallbackUsed) {
          notificationText += "⚠️ Fallback: " + targetFolderResult.warning + "\n";
        }
      } else {
        notificationText += "📂 Location: Project Root\n";
      }
      
      // File statistics
      notificationText += "📎 Files processed: " + selectedAttachments.length + "\n";
      notificationText += "💾 Files saved: " + savedCount + "\n";
      
      if (skippedCount > 0) {
        notificationText += "⚠️ Duplicates skipped: " + skippedCount + "\n";
      }
      
      // Project information
      notificationText += "🎫 Project: " + finalTicket + "\n";
      
      // Success summary
      var successRate = Math.round((savedCount / selectedAttachments.length) * 100);
      notificationText += "✨ Success rate: " + successRate + "%";
      
      if (savedFiles.length > 0 && savedFiles.length <= 3) {
        notificationText += "\n\n📋 Saved files:\n• " + savedFiles.join("\n• ");
      } else if (savedFiles.length > 3) {
        notificationText += "\n\n📋 Recent files: " + savedFiles.slice(0, 2).join(", ") + " and " + (savedFiles.length - 2) + " more";
      }
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(notificationText))
        .build();
        
    } catch (error) {
      console.error("=== SAVE ERROR ===");
      console.error("Error:", error);
      console.error("Stack:", error.stack);
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Error: " + error.message))
        .build();
    }
  }
  
  // ===== SETTINGS STORAGE FUNCTIONS =====
  
  function getUserSettings() {
    console.log("=== LOADING USER SETTINGS ===");
    console.log("Timestamp:", new Date().toISOString());
    
    try {
      var userProperties = PropertiesService.getUserProperties();
      var settingsJson = userProperties.getProperty(getSettingsStorageKey());
      
      console.log("Settings JSON from storage:", settingsJson ? "found" : "not found");
      
      if (settingsJson) {
        console.log("Raw settings JSON length:", settingsJson.length);
        var settings = JSON.parse(settingsJson);
        console.log("Parsed settings keys:", Object.keys(settings));
        
        console.log("Final settings configuration:");
        console.log("- Jira URL:", settings.jiraUrl ? "configured" : "NOT SET");
        console.log("- Jira Token:", settings.jiraToken ? "configured (length: " + settings.jiraToken.length + ")" : "NOT SET");
        console.log("- JQL Query:", settings.jqlQuery ? "configured" : "NOT SET");
        console.log("- Customers Folder ID:", getCustomersFolderId());
        
        return settings;
      } else {
        console.log("No existing settings found - returning defaults for new user");
        var defaultSettings = {};
        console.log("Default settings:", JSON.stringify(defaultSettings));
        return defaultSettings;
      }
    } catch (error) {
      console.error("CRITICAL: Error getting user settings:", error.name, "-", error.message);
      console.error("Settings error stack:", error.stack);
      return {};
    }
  }
  
  function setUserSettings(settings) {
    try {
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty(getSettingsStorageKey(), JSON.stringify(settings));
      console.log("Settings saved to user properties");
    } catch (error) {
      console.error("Error setting user settings:", error);
      throw error;
    }
  }
  
  function getDefaultJQL() {
    return 'project = CXPRODELIVERY AND issuetype in (Project, "Project (Standard Solution)") AND status in (HYPERCARE, "Order received", "Test system available", "Project go-live/productive start", "System Design Assigned", "Implementation Assigned", "Implementation Order Assigned", "Test System Available (Implementation Order)", "Handover Check Needed", "HYPERCARE (WITH CHECK)", "LIVE SYSTEM AVAILABLE", "System Design Order Received", "System Design Started", "Requirements Clarified", "Implementation Started") AND "Technical Project Manager" in (currentUser())';
  }
  
  function maskToken(token) {
    if (!token || token.length < 8) return "****";
    return token.substring(0, 4) + "****" + token.substring(token.length - 4);
  }
  
  function testJiraAPI(settings) {
    try {
      var testUrl = settings.jiraUrl + '/rest/api/2/myself';
      
      var response = UrlFetchApp.fetch(testUrl, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + settings.jiraToken,
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      });
      
      var responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        var data = JSON.parse(response.getContentText());
        return {
          success: true,
          userInfo: data.displayName || data.name || 'User',
          email: data.emailAddress || 'Unknown'
        };
      } else {
        var errorData = response.getContentText();
        return {
          success: false,
          error: "HTTP " + responseCode + ": " + errorData
        };
      }
      
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  // ===== FOLDER ACCESS REQUEST CARD =====
  
  /**
   * Build a card for requesting folder access
   * @param {string} ticketKey - The Jira ticket key
   * @param {string} folderId - The Google Drive folder ID
   * @param {string} errorMessage - The detailed error message
   * @returns {Card} A card with request access button
   */
  function buildRequestAccessCard(ticketKey, folderId, errorMessage) {
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle("🔒 Folder Access Required")
        .setSubtitle(ticketKey));
    
    var section = CardService.newCardSection();
    
    // Error explanation
    var errorText = CardService.newTextParagraph()
      .setText("<b>⚠️ Cannot Access Project Folder</b><br><br>" +
               "The project folder was found, but you don't have permission to access it yet.");
    section.addWidget(errorText);
    
    // Folder info
    var folderInfo = CardService.newKeyValue()
      .setTopLabel("Project Ticket")
      .setContent(ticketKey)
      .setIcon(CardService.Icon.TICKET);
    section.addWidget(folderInfo);
    
    var folderIdInfo = CardService.newKeyValue()
      .setTopLabel("Folder ID")
      .setContent(folderId)
      .setIcon(CardService.Icon.DESCRIPTION);
    section.addWidget(folderIdInfo);
    
    // Technical error (collapsible)
    var technicalError = CardService.newTextParagraph()
      .setText("<font color=\"#888888\"><i>Technical details: " + errorMessage + "</i></font>");
    section.addWidget(technicalError);
    
    card.addSection(section);
    
    // Action section
    var actionSection = CardService.newCardSection()
      .setHeader("🔧 What to Do");
    
    var instructions = CardService.newTextParagraph()
      .setText("Click the button below to open the folder in Google Drive and request access:");
    actionSection.addWidget(instructions);
    
    // Request access button with direct folder URL
    var folderUrl = "https://drive.google.com/drive/folders/" + folderId;
    var requestAccessButton = CardService.newTextButton()
      .setText("🔗 Open Folder & Request Access")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(folderUrl)
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING));
    
    var buttonSet = CardService.newButtonSet()
      .addButton(requestAccessButton);
    actionSection.addWidget(buttonSet);
    
    // Additional instructions
    var additionalHelp = CardService.newTextParagraph()
      .setText("<br><b>After requesting access:</b><br>" +
               "• The folder owner will receive your request<br>" +
               "• You'll get an email when access is granted<br>" +
               "• Come back and try saving again<br><br>" +
               "<b>Alternative:</b> Contact your project manager for immediate access");
    actionSection.addWidget(additionalHelp);
    
    card.addSection(actionSection);
    
    return card.build();
  }
  
  // ===== SUBFOLDER ERROR HANDLING =====
  
  // ===== SUBFOLDER ERROR HANDLING =====
  
  /**
   * Get error message for subfolder creation failures
   * @param {string} error - The error message
   * @param {Folder} projectFolder - The project root folder
   * @param {string} subfolderPath - The subfolder path
   * @returns {string} User-friendly error message
   */
  function getSubfolderCreationErrorMessage(error, projectFolder, subfolderPath) {
    var errorMsg = "❌ Cannot create subfolder in project folder\n\n";
    errorMsg += "📁 Project Folder: " + projectFolder.getName() + "\n";
    errorMsg += "🎯 Subfolder Path: " + subfolderPath + "\n";
    errorMsg += "📋 Technical Error: " + error + "\n\n";
    
    errorMsg += "🔧 Possible Solutions:\n";
    errorMsg += "• Check if you have write permissions to the project folder\n";
    errorMsg += "• Try saving to 'Project Root' as a fallback\n";
    errorMsg += "• Contact project admin if folder permissions are restricted\n\n";
    
    errorMsg += "💡 Tip: You can still save attachments by selecting 'Project Root'";
    
    return errorMsg;
  }
  
  // ===== FOLDER ACCESS FUNCTIONS =====
  
  function getFolderByIdSafely(folderId) {
    try {
      verboseLog("Accessing folder: " + folderId);
      
      // Try Advanced Drive API first (supports Shared Drives)
      var fileMetadata;
      var folder;
      var folderName;
      
      try {
        fileMetadata = Drive.Files.get(folderId, {
          supportsAllDrives: true,
          fields: 'id,name,driveId,parents'
        });
        
        folder = DriveApp.getFolderById(folderId);
        folderName = fileMetadata.name;
        
      } catch (advancedApiError) {
        // Fallback to standard DriveApp (for My Drive folders)
        verboseLog("Fallback to DriveApp: " + advancedApiError.message);
        
        folder = DriveApp.getFolderById(folderId);
        folderName = folder.getName();
      }
      
      // Detect if folder is in Shared Drive (simplified - skip hierarchy traversal for performance)
      var isSharedDrive = false;
      var sharedDriveName = null;
      var folderPath = [folderName];
      
      // Use Advanced Drive API metadata to detect Shared Drive (fast!)
      if (fileMetadata && fileMetadata.driveId) {
        isSharedDrive = true;
        verboseLog("📂 Shared Drive folder: " + folderName);
      }
      
      return {
        success: true,
        folder: folder,
        name: folderName,
        isSharedDrive: isSharedDrive,
        sharedDriveName: sharedDriveName,
        folderPath: folderPath.join(" > ")
      };
    } catch (error) {
      console.error("❌ Cannot access folder " + folderId + ": " + error.message);
      return {
        success: false,
        error: 'Cannot access folder (ID: ' + folderId + '): ' + error.message
      };
    }
  }
  
  function logDiagnosticInfo() {
    console.log("=== DIAGNOSTIC INFORMATION DUMP ===");
    console.log("Timestamp:", new Date().toISOString());
    console.log("Apps Script execution info:");
    
    try {
      // Log execution environment (with permission error handling)
      try {
        console.log("- Time zone:", Session.getScriptTimeZone());
      } catch (tzError) {
        console.log("- Time zone: Error getting timezone -", tzError.message);
      }
      
      try {
        console.log("- Active user email:", Session.getActiveUser().getEmail());
      } catch (userError) {
        console.log("- Active user email: Permission error -", userError.message);
        console.log("- Note: userinfo.email OAuth scope may need authorization");
      }
      
      try {
        console.log("- Script permissions:", Session.getScriptTimeZone() ? "timezone authorized" : "timezone not authorized");
      } catch (permError) {
        console.log("- Script permissions: Error checking permissions -", permError.message);
      }
      
      // Log settings
      console.log("\nUser settings diagnostic:");
      var settings = getUserSettings();
      console.log("- Settings object type:", typeof settings);
      console.log("- Settings keys:", Object.keys(settings || {}));
      if (settings) {
        console.log("- Jira URL configured:", !!settings.jiraUrl);
        console.log("- Jira Token configured:", !!settings.jiraToken);
        console.log("- Customers Folder ID:", getCustomersFolderId());
      }
      
      // Log folder access
      console.log("\nFolder access diagnostic:");
      console.log("- Customers Folder ID:", getCustomersFolderId());
      console.log("- URL fetch available:", typeof UrlFetchApp !== 'undefined');
      
      // Log Gmail context (with API access error handling)
      console.log("\nAPI availability diagnostic:");
      try {
        if (typeof GmailApp !== 'undefined') {
          console.log("- Gmail API: Available");
          // Test basic Gmail access
          try {
            var gmailQuota = GmailApp.getRemainingDailyQuota();
            console.log("- Gmail quota remaining:", gmailQuota);
          } catch (gmailError) {
            console.log("- Gmail access: Permission error -", gmailError.message);
          }
        } else {
          console.log("- Gmail API: Not available (manual testing mode)");
        }
      } catch (gmailCheckError) {
        console.log("- Gmail API: Error checking availability -", gmailCheckError.message);
      }
      
      try {
        if (typeof DriveApp !== 'undefined') {
          console.log("- Drive API: Available");
          // Test basic Drive access
          try {
            var driveQuota = DriveApp.getStorageUsed();
            console.log("- Drive storage used:", driveQuota, "bytes");
          } catch (driveError) {
            console.log("- Drive access: Permission error -", driveError.message);
          }
        } else {
          console.log("- Drive API: Not available");
        }
      } catch (driveCheckError) {
        console.log("- Drive API: Error checking availability -", driveCheckError.message);
      }
      
      try {
        console.log("- UrlFetch API: Available -", typeof UrlFetchApp !== 'undefined');
        console.log("- CardService API: Available -", typeof CardService !== 'undefined');
        console.log("- PropertiesService API: Available -", typeof PropertiesService !== 'undefined');
      } catch (apiCheckError) {
        console.log("- API availability check error:", apiCheckError.message);
      }
      
      console.log("=== END DIAGNOSTIC INFO ===");
    } catch (diagError) {
      console.error("Error during diagnostic logging:", diagError.message);
    }
  }
  
  /**
   * Reauthorization function for users
   * Run this function to trigger OAuth reauthorization
   * 
   * This is useful when:
   * - OAuth scopes have been added/modified
   * - User revoked permissions
   * - Permission errors occur
   * - Testing access after deployment
   * 
   * Usage: Simply run this function from Apps Script Editor
   */
  function reauthorizeUser() {
    console.log("=".repeat(80));
    console.log("    USER REAUTHORIZATION");
    console.log("=".repeat(80));
    console.log("Started:", new Date().toISOString());
    console.log("\n📋 This function will test all OAuth permissions");
    console.log("   If permissions are missing, you'll be prompted to reauthorize\n");
    
    var authResults = {
      timestamp: new Date().toISOString(),
      user: null,
      permissions: {},
      needsReauth: false,
      errors: []
    };
    
    try {
      // Get user info first
      console.log("👤 Identifying user...");
      try {
        var userEmail = Session.getActiveUser().getEmail();
        var userTimezone = Session.getScriptTimeZone();
        authResults.user = {
          email: userEmail,
          timezone: userTimezone
        };
        console.log("✅ User:", userEmail);
        console.log("✅ Timezone:", userTimezone);
      } catch (userError) {
        console.log("⚠️  Could not get user info:", userError.message);
        authResults.errors.push("User info: " + userError.message);
      }
      
      console.log("\n🔐 Testing OAuth Permissions:");
      console.log("=".repeat(50));
      
      // Test 1: Gmail Access
      console.log("\n1️⃣ Gmail API Access");
      try {
        var gmailQuota = GmailApp.getRemainingDailyQuota();
        console.log("   ✅ GRANTED - Daily quota remaining:", gmailQuota);
        authResults.permissions.gmail = {
          granted: true,
          quota: gmailQuota
        };
      } catch (gmailError) {
        console.log("   ❌ DENIED -", gmailError.message);
        authResults.permissions.gmail = {
          granted: false,
          error: gmailError.message
        };
        authResults.needsReauth = true;
        authResults.errors.push("Gmail: " + gmailError.message);
      }
      
      // Test 2: Drive Access
      console.log("\n2️⃣ Google Drive API Access");
      try {
        var driveQuota = DriveApp.getStorageUsed();
        var driveLimit = DriveApp.getStorageLimit();
        console.log("   ✅ GRANTED - Storage used:", driveQuota, "of", driveLimit, "bytes");
        authResults.permissions.drive = {
          granted: true,
          storageUsed: driveQuota,
          storageLimit: driveLimit
        };
      } catch (driveError) {
        console.log("   ❌ DENIED -", driveError.message);
        authResults.permissions.drive = {
          granted: false,
          error: driveError.message
        };
        authResults.needsReauth = true;
        authResults.errors.push("Drive: " + driveError.message);
      }
      
      // Test 3: External Requests (for API calls)
      console.log("\n3️⃣ External Requests (UrlFetch)");
      try {
        if (typeof UrlFetchApp === 'undefined') {
          throw new Error("UrlFetchApp not available");
        }
        // Test with a safe endpoint
        var testResponse = UrlFetchApp.fetch("https://www.google.com", {
          method: 'HEAD',
          muteHttpExceptions: true,
          timeout: 5000
        });
        console.log("   ✅ GRANTED - Test request successful");
        authResults.permissions.urlfetch = {
          granted: true,
          testResponseCode: testResponse.getResponseCode()
        };
      } catch (fetchError) {
        console.log("   ❌ DENIED -", fetchError.message);
        authResults.permissions.urlfetch = {
          granted: false,
          error: fetchError.message
        };
        authResults.needsReauth = true;
        authResults.errors.push("UrlFetch: " + fetchError.message);
      }
      
      // Test 4: Script Properties (for settings storage)
      console.log("\n4️⃣ Script Properties Access");
      try {
        var testValue = "test_" + new Date().getTime();
        var props = PropertiesService.getUserProperties();
        props.setProperty('_reauth_test', testValue);
        var retrieved = props.getProperty('_reauth_test');
        props.deleteProperty('_reauth_test');
        
        if (retrieved === testValue) {
          console.log("   ✅ GRANTED - Properties read/write working");
          authResults.permissions.properties = { granted: true };
        } else {
          throw new Error("Property read/write mismatch");
        }
      } catch (propError) {
        console.log("   ❌ DENIED -", propError.message);
        authResults.permissions.properties = {
          granted: false,
          error: propError.message
        };
        authResults.errors.push("Properties: " + propError.message);
      }
      
      // Summary
      console.log("\n" + "=".repeat(80));
      console.log("    REAUTHORIZATION SUMMARY");
      console.log("=".repeat(80));
      
      var grantedCount = 0;
      var totalCount = 0;
      
      for (var perm in authResults.permissions) {
        totalCount++;
        if (authResults.permissions[perm].granted) {
          grantedCount++;
        }
      }
      
      console.log("\n📊 RESULTS:");
      console.log("  Permissions Granted:", grantedCount, "of", totalCount);
      console.log("  Gmail:", authResults.permissions.gmail ? (authResults.permissions.gmail.granted ? "✅" : "❌") : "⚠️");
      console.log("  Drive:", authResults.permissions.drive ? (authResults.permissions.drive.granted ? "✅" : "❌") : "⚠️");
      console.log("  UrlFetch:", authResults.permissions.urlfetch ? (authResults.permissions.urlfetch.granted ? "✅" : "❌") : "⚠️");
      console.log("  Properties:", authResults.permissions.properties ? (authResults.permissions.properties.granted ? "✅" : "❌") : "⚠️");
      
      if (grantedCount === totalCount) {
        console.log("\n🎉 SUCCESS: All permissions granted!");
        console.log("  ✅ Your add-on is fully authorized");
        console.log("  ✅ Ready to save attachments to Google Drive");
        console.log("  ✅ No reauthorization needed");
      } else {
        console.log("\n⚠️ WARNING: Some permissions missing");
        console.log("\n📝 TO FIX:");
        console.log("  1. Check that appsscript.json contains all required scopes");
        console.log("  2. Redeploy the add-on");
        console.log("  3. Users will be prompted to reauthorize on next use");
        console.log("\n  Required scopes:");
        console.log("    - https://www.googleapis.com/auth/gmail.readonly");
        console.log("    - https://www.googleapis.com/auth/gmail.addons.execute");
        console.log("    - https://www.googleapis.com/auth/gmail.addons.current.message.readonly");
        console.log("    - https://www.googleapis.com/auth/drive");
        console.log("    - https://www.googleapis.com/auth/script.external_request");
        console.log("    - https://www.googleapis.com/auth/userinfo.email");
      }
      
      if (authResults.errors.length > 0) {
        console.log("\n❌ ERRORS ENCOUNTERED:");
        for (var i = 0; i < authResults.errors.length; i++) {
          console.log("  " + (i + 1) + ". " + authResults.errors[i]);
        }
      }
      
      console.log("\n💡 NOTE:");
      console.log("  • If you just updated OAuth scopes, redeploy the add-on");
      console.log("  • Users will be automatically prompted to reauthorize");
      console.log("  • This happens when they next open Gmail and use the add-on");
      console.log("  • No manual user action needed beyond accepting the new permissions");
      
      return authResults;
      
    } catch (error) {
      console.error("\n❌ REAUTHORIZATION FAILED:");
      console.error("  Error:", error.message);
      console.error("  Stack:", error.stack);
      
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  /**
   * Force reauthorization by accessing all required scopes
   * This triggers the OAuth prompt if permissions are missing
   * 
   * Usage: Run this function, then accept the authorization prompt
   */
  function forceReauthorization() {
    console.log("=".repeat(80));
    console.log("    FORCE REAUTHORIZATION");
    console.log("=".repeat(80));
    console.log("\n🔄 This will trigger OAuth reauthorization if needed");
    console.log("   You may see an authorization prompt - please accept it\n");
    
    try {
      // Access each service to trigger OAuth
      console.log("1. Accessing Gmail API...");
      GmailApp.getRemainingDailyQuota();
      console.log("   ✅ Gmail access confirmed");
      
      console.log("\n2. Accessing Drive API...");
      DriveApp.getStorageUsed();
      console.log("   ✅ Drive access confirmed");
      
      console.log("\n3. Testing external requests...");
      UrlFetchApp.fetch("https://www.google.com", {
        method: 'HEAD',
        muteHttpExceptions: true,
        timeout: 5000
      });
      console.log("   ✅ External requests confirmed");
      
      console.log("\n4. Accessing user info...");
      Session.getActiveUser().getEmail();
      console.log("   ✅ User info access confirmed");
      
      console.log("\n" + "=".repeat(80));
      console.log("🎉 REAUTHORIZATION COMPLETE");
      console.log("=".repeat(80));
      console.log("\n✅ All permissions have been verified");
      console.log("✅ Your add-on is fully authorized");
      console.log("✅ You can now use the Gmail Attachment Saver");
      
      return { success: true, message: "Reauthorization successful" };
      
    } catch (error) {
      console.error("\n❌ REAUTHORIZATION ERROR:");
      console.error("  Message:", error.message);
      console.error("\n💡 SOLUTION:");
      console.error("  1. Accept the authorization prompt when it appears");
      console.error("  2. Make sure all required scopes are in appsscript.json");
      console.error("  3. Redeploy the add-on if scopes were changed");
      
      return { success: false, error: error.message };
    }
  }
  
  // ===== HELPER FUNCTIONS =====
  
  function formatCompactTicketDisplay(issue) {
    var summary = issue.summary || 'No summary';
    var statusEmoji = getStatusEmoji(issue.status);
    
    // Extract client info from summary
    var clientMatch = summary.match(/^(\d+)\s*-\s*([^|]+)/);
    
    if (clientMatch && clientMatch[2]) {
      var clientName = clientMatch[2].trim();
      
      // Shorten client name to fit dropdown (keep it reasonable for dropdown width)
      if (clientName.length > 20) {
        clientName = clientName.substring(0, 17) + "...";
      }
      
      // Format: [Emoji] [Full Ticket Key] - [Client Name]
      return statusEmoji + " " + issue.key + " - " + clientName;
      
    } else {
      // Fallback: just show ticket key with emoji
      return statusEmoji + " " + issue.key;
    }
  }
  
  function getStatusEmoji(status) {
    var statusEmojis = {
      'HYPERCARE': '🔧',
      'HYPERCARE (WITH CHECK)': '🔍',
      'Order received': '📋',
      'Test system available': '🧪',
      'Test System Available (Implementation Order)': '🧪',
      'Project go-live/productive start': '🚀',
      'System Design Assigned': '👤',          // User assigned to design
      'System Design Order Received': '📨',
      'System Design Started': '🎨',
      'Implementation Assigned': '👨‍💻',        // User assigned to implement
      'Implementation Order Assigned': '🧑‍💼',  // User assigned to manage order
      'Implementation Started': '🔨',
      'Requirements Clarified': '✅',
      'Handover Check Needed': '🔍',
      'LIVE SYSTEM AVAILABLE': '🟢'
    };
    
    return statusEmojis[status] || '📝';
  }
  
  function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    var k = 1024;
    var sizes = ['B', 'KB', 'MB', 'GB'];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }
  
  /**
   * Get appropriate icon for Google Docs type
   * @param {string} docType - The type of Google Doc (docs, sheets, slides, etc.)
   * @returns {string} Emoji icon for the doc type
   */
  function getGoogleDocsIcon(docType) {
    var icons = {
      'docs': '📄',      // Document icon
      'sheets': '📊',    // Spreadsheet icon  
      'slides': '📽️',    // Presentation icon
      'forms': '📝',     // Form icon
      'drive': '📁',     // Generic drive file icon
      'drawings': '🎨'   // Drawing icon
    };
    
    return icons[docType] || '🔗'; // Default link icon for unknown types
  }
  
  function getOrCreateFolder(folderName, parentFolder) {
    var parent = parentFolder || DriveApp.getRootFolder();
    var folders = parent.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return parent.createFolder(folderName);
    }
  }
  
  // ===== SUBFOLDER MANAGEMENT FOR PROJECTS =====
  
  /**
   * Get standardized project subfolder configuration
   * @returns {Array} Array of subfolder definitions
   */
  function getProjectSubfolderConfig() {
    return [
      {
        name: "01_System_Design",
        icon: "📁",
        description: "System architecture, diagrams, technical specifications"
      },
      {
        name: "02_Meet_Recordings", 
        icon: "📁",
        description: "Meeting recordings, session notes, call summaries"
      },
      {
        name: "03_Correspondence",
        icon: "📁", 
        description: "Email threads, communications, external correspondence"
      },
      {
        name: "04_Project_Documentation",
        icon: "📁",
        description: "General project documents, reports, deliverables"
      }
    ];
  }
  
  /**
   * Create or access project subfolder with support for nested paths
   * @param {Folder} parentFolder - The project root folder
   * @param {string} subfolderPath - Path like "01_System_Design" or "04_Project_Documentation/Project_Management"
   * @returns {Object} Result object with folder info
   */
function getOrCreateProjectSubfolder(parentFolder, subfolderPath) {
  try {
    if (!subfolderPath || subfolderPath.trim() === "") {
      return {
        success: true,
        folder: parentFolder,
        path: "Project Root",
        created: false
      };
    }
    
    // Split path for nested folders (e.g., "04_Project_Documentation/Project_Management")
    var pathParts = subfolderPath.split('/');
    var currentFolder = parentFolder;
    var createdFolders = [];
    var fullPath = "";
    
    for (var i = 0; i < pathParts.length; i++) {
      var folderName = pathParts[i].trim();
      if (!folderName) continue;
      
      fullPath += (fullPath ? "/" : "") + folderName;
      
      // Check if folder already exists
      var existingFolders = currentFolder.getFoldersByName(folderName);
      
      if (existingFolders.hasNext()) {
        currentFolder = existingFolders.next();
      } else {
        // Create new folder
        currentFolder = currentFolder.createFolder(folderName);
        createdFolders.push(folderName);
      }
    }
    
    if (createdFolders.length > 0) {
      console.log("📁 Created subfolder(s): " + createdFolders.join(", "));
    }
    
    return {
      success: true,
      folder: currentFolder,
      path: fullPath,
      created: createdFolders.length > 0,
      createdFolders: createdFolders
    };
    
  } catch (error) {
    console.error("❌ Subfolder error: " + error.message);
    
    return {
      success: false,
      error: "Failed to create/access subfolder: " + error.message,
      path: subfolderPath
    };
  }
}
  
  /**
   * Validate subfolder selection against allowed paths
   * @param {string} subfolderPath - The subfolder path to validate
   * @returns {Object} Validation result
   */
  function validateSubfolderSelection(subfolderPath) {
    // Define allowed subfolder patterns
    var allowedPaths = [
      "", // Project root
      "01_System_Design",
      "02_Meet_Recordings", 
      "03_Correspondence",
      "04_Project_Documentation",
      "04_Project_Documentation/Project_Management",
      "04_Project_Documentation/Carrier_Onboarding"
    ];
    
    if (!subfolderPath || subfolderPath.trim() === "") {
      return { valid: true, path: "", sanitized: "" };
    }
    
    // Check against allowed paths
    if (allowedPaths.indexOf(subfolderPath) !== -1) {
      return { 
        valid: true, 
        path: subfolderPath,
        sanitized: subfolderPath
      };
    }
    
    // Sanitize path - remove invalid characters
    var sanitized = subfolderPath
      .replace(/[<>:"|?*]/g, '_') // Replace invalid chars
      .replace(/\.\./g, '_') // Prevent directory traversal
      .replace(/^\/+|\/+$/g, '') // Remove leading/trailing slashes
      .substring(0, 100); // Limit length
    
    console.log("⚠ Subfolder path sanitized:", subfolderPath, "→", sanitized);
    
    return {
      valid: false,
      path: subfolderPath,
      sanitized: sanitized,
      warning: "Path was not in predefined list and was sanitized"
    };
  }
  
  /**
   * Handle subfolder creation failure with graceful degradation
   * @param {Folder} projectFolder - The project root folder
   * @param {string} subfolderPath - The failed subfolder path
   * @param {Object} errorDetails - Error information
   * @returns {Object} Fallback result
   */
  function handleSubfolderCreationFailure(projectFolder, subfolderPath, errorDetails) {
    console.log("=== SUBFOLDER FALLBACK STRATEGY ===");
    console.log("Original path:", subfolderPath);
    console.log("Error:", errorDetails.error);
    
    // Strategy 1: Try simplified folder name (remove nesting, use only first level)
    if (subfolderPath.indexOf("/") !== -1) {
      var simplifiedPath = subfolderPath.split("/")[0]; // Use only first level
      console.log("Trying simplified path:", simplifiedPath);
      
      var fallbackResult = getOrCreateProjectSubfolder(projectFolder, simplifiedPath);
      if (fallbackResult.success) {
        console.log("✓ Fallback successful with simplified path");
        return {
          success: true,
          folder: fallbackResult.folder,
          fallbackUsed: "simplified_path",
          originalPath: subfolderPath,
          actualPath: simplifiedPath,
          warning: "Used simplified folder path due to nesting issue"
        };
      }
    }
    
    // Strategy 2: Fall back to project root
    console.log("All subfolder creation failed, using project root");
    return {
      success: true,
      folder: projectFolder,
      fallbackUsed: "project_root",
      originalPath: subfolderPath,
      actualPath: "Project Root",
      warning: "Saved to project root due to subfolder creation error"
    };
  }
  
  /**
   * Get user preferences for subfolder feature
   * @returns {Object} User preferences
   */
  function getUserSubfolderPreferences() {
    var userProperties = PropertiesService.getUserProperties();
    var preferences = {
      enableSubfolderSelection: userProperties.getProperty('ENABLE_SUBFOLDER_SELECTION') !== 'false', // Default true
      defaultSubfolder: userProperties.getProperty('DEFAULT_SUBFOLDER') || '',
      showSubfolderTooltips: userProperties.getProperty('SHOW_SUBFOLDER_TOOLTIPS') !== 'false'
    };
    
    return preferences;
  }
  
  /**
   * Check if enhanced UI should be used
   * @returns {boolean} True if enhanced UI should be shown
   */
  function shouldUseEnhancedUI() {
    try {
      var preferences = getUserSubfolderPreferences();
      return preferences.enableSubfolderSelection;
    } catch (error) {
      console.log("Error loading preferences, defaulting to enhanced UI:", error);
      return true; // Default to enhanced UI
    }
  }
  
  // ===== JIRA INTEGRATION (Updated to use user settings) =====
  
  function getMyJiraProjects() {
    try {
      var settings = getUserSettings();
      
      if (!settings.jiraUrl || !settings.jiraToken) {
        console.error('Jira settings not configured');
        return { projects: {}, issues: [], totalIssues: 0 };
      }
      
      var searchUrl = settings.jiraUrl + '/rest/api/2/search';
      var jqlQuery = settings.customJql || getDefaultJQL();
      
      var payload = {
        jql: jqlQuery,
        fields: ['project', 'key', 'summary', 'status', 'issuetype'],
        maxResults: 1000,
        startAt: 0
      };
      
      console.log('Using JQL filter:', jqlQuery);
      
      var response = UrlFetchApp.fetch(searchUrl, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + settings.jiraToken,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.error('Jira API failed:', response.getResponseCode(), response.getContentText());
        return { projects: {}, issues: [], totalIssues: 0 };
      }
      
      var data = JSON.parse(response.getContentText());
      var projects = {};
      var issues = [];
      
      console.log('JQL Filter returned', data.total || 0, 'issues');
      
      if (data.issues && data.issues.length > 0) {
        data.issues.forEach(function(issue) {
          var project = issue.fields.project;
          if (project) {
            projects[project.key] = {
              key: project.key,
              name: project.name,
              id: project.id
            };
            
            issues.push({
              key: issue.key,
              summary: issue.fields.summary || 'No summary',
              status: issue.fields.status ? issue.fields.status.name : 'Unknown',
              issueType: issue.fields.issuetype ? issue.fields.issuetype.name : 'Unknown',
              projectKey: project.key,
              projectName: project.name
            });
          }
        });
      }
      
      console.log('Projects found:', Object.keys(projects));
      console.log('Issues found:', issues.length);
      
      return { 
        projects: projects, 
        issues: issues,
        totalIssues: data.total || 0 
      };
      
    } catch (error) {
      console.error('Error getting Jira projects:', error);
      return { projects: {}, issues: [], totalIssues: 0 };
    }
  }
  
  // ===== GOOGLE DOCS LINK EXTRACTION FUNCTIONS =====
  
  /**
   * Extract Google Docs links from email message content
   * @param {GmailMessage} message - The Gmail message to analyze
   * @returns {Array} Array of Google Docs link objects
   */
  function extractGoogleDocsLinks(message) {
    try {
      console.log("Extracting Google Docs links from message...");
      
      var links = [];
      var htmlBody = '';
      var plainBody = '';
      
      // Get message content
      try {
        htmlBody = message.getBody() || '';
        plainBody = message.getPlainBody() || '';
      } catch (e) {
        console.log("Could not get message body:", e.message);
        return links;
      }
      
      // Define Google service URL patterns
      var patterns = [
        // Google Docs
        {
          regex: /https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'docs',
          defaultTitle: 'Google Doc'
        },
        // Google Sheets  
        {
          regex: /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'sheets',
          defaultTitle: 'Google Sheet'
        },
        // Google Slides
        {
          regex: /https:\/\/docs\.google\.com\/presentation\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'slides', 
          defaultTitle: 'Google Slides'
        },
        // Google Forms
        {
          regex: /https:\/\/docs\.google\.com\/forms\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'forms',
          defaultTitle: 'Google Form'
        },
        // Google Drive files (generic)
        {
          regex: /https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'drive',
          defaultTitle: 'Google Drive File'
        }
      ];
      
      // Search in both HTML and plain text
      var contentToSearch = [
        { content: htmlBody, isHtml: true },
        { content: plainBody, isHtml: false }
      ];
      
      var foundUrls = new Set(); // Prevent duplicates
      
      contentToSearch.forEach(function(contentObj) {
        patterns.forEach(function(pattern) {
          var matches;
          pattern.regex.lastIndex = 0; // Reset regex state
          
          while ((matches = pattern.regex.exec(contentObj.content)) !== null) {
            var fullUrl = matches[0];
            var fileId = matches[1];
            
            // Skip if we already found this URL
            if (foundUrls.has(fullUrl)) {
              continue;
            }
            foundUrls.add(fullUrl);
            
            // Try to extract title from HTML context
            var title = extractTitleFromContext(contentObj.content, fullUrl, contentObj.isHtml);
            if (!title) {
              title = pattern.defaultTitle;
            }
            
            // Ensure unique title by adding file ID if needed
            var finalTitle = title;
            var titleExists = links.some(function(link) { return link.title === finalTitle; });
            if (titleExists) {
              finalTitle = title + " (" + fileId.substring(0, 8) + "...)";
            }
            
            links.push({
              url: fullUrl,
              fileId: fileId,
              title: finalTitle,
              type: pattern.type
            });
            
            console.log("Found Google " + pattern.type + " link:", finalTitle, "->", fullUrl);
          }
        });
      });
      
      console.log("Extracted", links.length, "Google Docs links from message");
      return links;
      
    } catch (error) {
      console.error("Error extracting Google Docs links:", error.message);
      return [];
    }
  }
  
  /**
   * Try to extract title from the context around a Google Docs URL
   * @param {string} content - The content to search in
   * @param {string} url - The URL to find context for
   * @param {boolean} isHtml - Whether the content is HTML
   * @returns {string|null} Extracted title or null
   */
  function extractTitleFromContext(content, url, isHtml) {
    try {
      var urlIndex = content.indexOf(url);
      if (urlIndex === -1) return null;
      
      if (isHtml) {
        // For HTML content, look for link text or nearby text
        var beforeUrl = content.substring(Math.max(0, urlIndex - 200), urlIndex);
        var afterUrl = content.substring(urlIndex + url.length, Math.min(content.length, urlIndex + url.length + 200));
        
        // Look for <a> tag with text
        var linkRegex = /<a[^>]*href[^>]*>([^<]+)<\/a>/i;
        var contextRegex = />([^<]{2,50})<.*?href.*?google\.com/i;
        var reverseContextRegex = /href.*?google\.com.*?>([^<]{2,50})</i;
        
        var match = beforeUrl.match(contextRegex) || afterUrl.match(reverseContextRegex) || 
                   (beforeUrl + url + afterUrl).match(linkRegex);
        
        if (match && match[1]) {
          var title = match[1].trim();
          // Clean up common HTML artifacts
          title = title.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
          if (title.length > 3 && title.length < 100) {
            return title;
          }
        }
      } else {
        // For plain text, look for text near the URL
        var beforeUrl = content.substring(Math.max(0, urlIndex - 100), urlIndex);
        var afterUrl = content.substring(urlIndex + url.length, Math.min(content.length, urlIndex + url.length + 100));
        
        // Look for text patterns that might be titles
        var patterns = [
          /([^\n\r]{5,80})\s*$/,  // Text before URL on same line
          /^\s*([^\n\r]{5,80})/   // Text after URL on same line
        ];
        
        for (var i = 0; i < patterns.length; i++) {
          var match = (i === 0 ? beforeUrl : afterUrl).match(patterns[i]);
          if (match && match[1]) {
            var title = match[1].trim();
            // Skip if it looks like part of a URL or email
            if (!title.includes('http') && !title.includes('@') && title.length < 80) {
              return title;
            }
          }
        }
      }
      
      return null;
      
    } catch (error) {
      console.error("Error extracting title from context:", error.message);
      return null;
    }
  }
  
  /**
   * Copy a Google Doc to the target folder while preserving its format
   * @param {Object} docLink - The Google Docs link object
   * @param {Folder} targetFolder - The folder to copy the Google Doc to
   * @returns {File} The copied Google Doc file
   */
  function copyGoogleDocToFolder(docLink, targetFolder) {
    try {
      // Extract document ID from Google Docs URL
      var docId = extractGoogleDocId(docLink.url);
      if (!docId) {
        return createFallbackShortcut(docLink, targetFolder);
      }
      
      // Get the Google Doc file from Drive
      var driveFile;
      try {
        driveFile = DriveApp.getFileById(docId);
      } catch (driveError) {
        return createFallbackShortcut(docLink, targetFolder);
      }
      
      // Use email subject as filename if available, otherwise fall back to original name
      var fileName;
      if (docLink.emailSubject && docLink.emailSubject.trim() !== '') {
        fileName = sanitizeFileName(docLink.emailSubject);
      } else {
        fileName = docLink.title || driveFile.getName();
      }
      
      // Check if a file with the same name already exists in the target folder
      var existingFiles = targetFolder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        var existingFile = existingFiles.next();
        
        // Check if it's the same file (same ID)
        if (existingFile.getId() === docId) {
          verboseLog("Google Doc already exists: " + fileName);
          return existingFile;
        }
        
        // Different file with same name - add timestamp to avoid conflict
        var timestamp = new Date().getTime();
        var nameParts = fileName.split('.');
        if (nameParts.length > 1) {
          fileName = nameParts.slice(0, -1).join('.') + '_' + timestamp + '.' + nameParts[nameParts.length - 1];
        } else {
          fileName = fileName + '_' + timestamp;
        }
      }
      
      // Create a copy of the Google Doc in the target folder
      try {
        var copiedFile = driveFile.makeCopy(fileName, targetFolder);
        return copiedFile;
        
      } catch (copyError) {
        return createFallbackShortcut(docLink, targetFolder);
      }
      
    } catch (error) {
      return createFallbackShortcut(docLink, targetFolder);
    }
  }

  // Helper function to sanitize email subject for use as filename
  function sanitizeFileName(subject) {
    try {
      if (!subject || typeof subject !== 'string') {
        return 'Google Doc';
      }
      
      // Remove or replace characters that are not allowed in filenames
      var sanitized = subject
        .replace(/[<>:"/\\|?*]/g, '_')  // Replace invalid filename characters
        .replace(/\s+/g, ' ')          // Replace multiple spaces with single space
        .trim()                        // Remove leading/trailing whitespace
        .substring(0, 100);            // Limit length
      
      return sanitized || 'Google Doc';
      
    } catch (error) {
      return 'Google Doc';
    }
  }

  // Helper function to extract Google Doc ID from URL
  function extractGoogleDocId(url) {
    try {
      // Handle different Google Docs URL formats
      var patterns = [
        /\/document\/d\/([a-zA-Z0-9-_]+)/,     // Documents
        /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/, // Sheets  
        /\/presentation\/d\/([a-zA-Z0-9-_]+)/, // Slides
        /\/file\/d\/([a-zA-Z0-9-_]+)/,         // Generic Drive files
        /id=([a-zA-Z0-9-_]+)/                  // ID parameter format
      ];
      
      for (var i = 0; i < patterns.length; i++) {
        var match = url.match(patterns[i]);
        if (match && match[1]) {
          return match[1];
        }
      }
      
      return null;
    } catch (error) {
      return null;
    }
  }


  // Fallback function to create shortcut if download fails
  function createFallbackShortcut(docLink, targetFolder) {
    try {
      // Create a shortcut file content pointing to the Google Doc
      var shortcutContent = '[InternetShortcut]\nURL=' + docLink.url + '\n';
      
      // Use email subject as filename if available, otherwise use original title
      var baseName = (docLink.emailSubject && docLink.emailSubject.trim() !== '') 
        ? sanitizeFileName(docLink.emailSubject) 
        : (docLink.title || 'Google Doc');
      
      var fileName = baseName + '.url';
      
      // Handle duplicate names
      var existingFiles = targetFolder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        fileName = baseName + '_' + new Date().getTime() + '.url';
      }
      
      // Create the shortcut file
      var blob = Utilities.newBlob(shortcutContent, 'text/plain', fileName);
      var file = targetFolder.createFile(blob);
      
      verboseLog("Created shortcut: " + fileName);
      return file;
      
    } catch (error) {
      console.error("❌ Shortcut creation failed: " + error.message);
      throw error;
    }
  }

  // ===== ATTACHMENT SELECTION MEMORY FUNCTIONS =====
  
  function storeAttachmentSelections(threadId, selectionState) {
    try {
      if (!threadId) {
        console.log("No threadId provided, skipping attachment selection storage");
        return;
      }
      
      var userProperties = PropertiesService.getUserProperties();
      var key = 'ATTACHMENT_SELECTIONS_' + threadId;
      
      console.log("=== STORING ATTACHMENT SELECTIONS ===");
      console.log("ThreadId:", threadId);
      console.log("Storage key:", key);
      console.log("Selection state to store:", JSON.stringify(selectionState));
      
      userProperties.setProperty(key, JSON.stringify(selectionState));
      
      // Verify storage worked
      var stored = userProperties.getProperty(key);
      console.log("Verification - stored value:", stored);
      console.log("Storage successful:", stored !== null);
      
    } catch (error) {
      console.error("Error storing attachment selections:", error);
    }
  }
  
  function getStoredAttachmentSelections(threadId) {
    try {
      if (!threadId) {
        console.log("No threadId provided, returning null");
        return null;
      }
      
      var userProperties = PropertiesService.getUserProperties();
      var key = 'ATTACHMENT_SELECTIONS_' + threadId;
      
      console.log("=== RETRIEVING ATTACHMENT SELECTIONS ===");
      console.log("ThreadId:", threadId);
      console.log("Storage key:", key);
      
      var storedSelectionsJson = userProperties.getProperty(key);
      console.log("Raw stored value:", storedSelectionsJson);
      
      if (storedSelectionsJson) {
        var selections = JSON.parse(storedSelectionsJson);
        console.log("Parsed selections:", selections);
        console.log("Selection keys:", Object.keys(selections));
        return selections;
      } else {
        console.log("No stored attachment selections found for thread:", threadId);
        
        // Debug: List all stored keys to see what's actually stored
        var allProperties = userProperties.getProperties();
        var attachmentKeys = [];
        for (var prop in allProperties) {
          if (prop.startsWith('ATTACHMENT_SELECTIONS_')) {
            attachmentKeys.push(prop);
          }
        }
        console.log("All attachment selection keys found:", attachmentKeys);
        
        return null;
      }
      
    } catch (error) {
      console.error("Error retrieving attachment selections:", error);
      return null;
    }
  }

/**
 * Save attachment using Advanced Drive API v3 with Shared Drive support
 * This version uses Drive.Files.create() with supportsAllDrives: true
 * 
 * @param {Blob} attachment - The attachment blob to save
 * @param {string} folderId - The Google Drive folder ID
 * @param {string} fileName - Optional custom file name
 * @returns {Object} - File object or error
 */
function saveAttachmentWithAdvancedDrive(attachment, folderId, fileName) {
  try {
    var finalFileName = fileName || attachment.getName();
    
    // Create file metadata
    var fileMetadata = {
      name: finalFileName,
      parents: [folderId]
    };
    
    // Create blob
    var blob = attachment.copyBlob();
    
    // Use Advanced Drive Service with Shared Drive support
    var startTime = new Date().getTime();
    var file = Drive.Files.create(fileMetadata, blob, {
      supportsAllDrives: true,  // Critical for Shared Drives!
      fields: 'id,name,size,mimeType,driveId'
    });
    var duration = new Date().getTime() - startTime;
    
    verboseLog("✅ Saved: " + file.name + " (" + duration + "ms)");
    
    return {
      success: true,
      file: file,
      fileId: file.id,
      fileName: file.name,
      duration: duration
    };
    
  } catch (error) {
    console.error("❌ Drive API save failed: " + error.message);
    
    return {
      success: false,
      error: error.message
    };
  }
}