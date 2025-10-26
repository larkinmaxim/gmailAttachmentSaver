# Settings Management Process Flow

## Overview
The Settings Management feature provides secure storage and management of Jira credentials, custom JQL queries, and application preferences with robust validation and error handling.

## Process Flow

### 1. Settings Architecture
```
User Interface → Validation → Secure Storage → Runtime Retrieval → Application Use
```

**Storage Layer:**
```javascript
// Google Apps Script Properties Service
PropertiesService.getUserProperties()
  ├── JIRA_SETTINGS (JSON object)
  └── ATTACHMENT_SELECTIONS_{threadId} (Selection memory)
```

**Settings Object Structure:**
```javascript
{
  jiraUrl: "https://support.transporeon.com",
  jiraToken: "encrypted_api_token_here", 
  customJql: "project = CXPRODELIVERY AND...",
  savedAt: "2024-10-26T14:30:22.000Z"
}
```

### 2. Initial Setup Flow
```
First Launch → Detect Missing Settings → Show Setup Card → Guide Configuration → Test & Save
```

**Detection Logic:**
```javascript
function buildAddOn(e) {
  var settings = getUserSettings();
  if (!settings.jiraToken || !settings.jiraUrl) {
    console.log("Settings not configured, showing setup");
    return [buildSettingsCard(true)]; // true = first time setup
  }
  // Continue with normal operation...
}
```

**First-Time Setup Card:**
- **Header:** "⚙️ Initial Setup Required"
- **Message:** "Welcome! Please configure your Jira connection to get started."
- **Required Fields:** Jira URL, API Token
- **Optional Fields:** Custom JQL Query
- **Actions:** Save Settings, Test Connection

### 3. Settings Input & Validation
```
User Input → Field Validation → Format Checking → Error Handling → Secure Storage
```

**Input Processing:**
```javascript
function saveSettings(e) {
  var jiraUrl = e.formInput.jiraUrl;
  var jiraToken = e.formInput.jiraToken;  
  var customJql = e.formInput.customJql;
  
  // Get existing settings for token masking logic
  var existingSettings = getUserSettings();
  
  // Handle masked token preservation
  var finalToken = handleTokenMasking(jiraToken, existingSettings);
  
  // Validate inputs
  var validation = validateSettings(jiraUrl, finalToken, customJql);
  if (!validation.valid) {
    return errorResponse(validation.error);
  }
  
  // Save to secure storage
  saveToProperties(jiraUrl, finalToken, customJql);
}
```

### 4. Security Features

#### 4.1 Token Masking System
```
Display Security → Input Processing → Storage Protection → Retrieval Safety
```

**Masking Logic:**
```javascript
function maskToken(token) {
  if (!token || token.length < 8) return "****";
  return token.substring(0, 4) + "****" + token.substring(token.length - 4);
}

// Display: "abcd****xyz9" 
// Storage: Full token (encrypted by Google Apps Script)
```

**Input Handling:**
```javascript
// If user didn't change masked token, preserve existing
if (existingSettings.jiraToken && jiraToken && jiraToken.includes("****")) {
  finalToken = existingSettings.jiraToken; // Keep existing
} else {
  finalToken = jiraToken; // Use new token
}
```

#### 4.2 Secure Storage Implementation
```javascript
function setUserSettings(settings) {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('JIRA_SETTINGS', JSON.stringify(settings));
    console.log("Settings saved to user properties");
  } catch (error) {
    console.error("Error setting user settings:", error);
    throw error;
  }
}
```

**Security Features:**
- **User-Specific:** Settings isolated per Google account
- **Encrypted Storage:** Google Apps Script automatic encryption
- **No Client Storage:** No sensitive data in browser/client-side
- **Access Control:** Only user's scripts can access their properties

### 5. Validation Framework
```
Input Data → Format Checks → Connectivity Tests → Business Logic → Error Reporting
```

**URL Validation:**
```javascript
function validateJiraUrl(url) {
  // Clean up URL
  if (url.endsWith('/')) {
    url = url.slice(0, -1);
  }
  
  // Format validation
  if (!url.startsWith('http')) {
    return {
      valid: false,
      error: "Please provide a valid URL starting with http"
    };
  }
  
  return { valid: true, cleanUrl: url };
}
```

**Token Validation:**
```javascript
function validateToken(token) {
  if (!token || token.trim().length === 0) {
    return {
      valid: false,
      error: "API token is required"
    };
  }
  
  // Basic format check (Jira tokens are typically longer)
  if (token.length < 10) {
    return {
      valid: false, 
      error: "API token appears to be too short"
    };
  }
  
  return { valid: true };
}
```

### 6. Connection Testing
```
Settings Input → API Test Call → Response Analysis → Success/Error Feedback
```

**Test Implementation:**
```javascript
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
      return {
        success: false,
        error: "HTTP " + responseCode + ": " + response.getContentText()
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}
```

**Test Results:**
- **Success:** `"✅ Connection successful! Logged in as: John Smith"`
- **Auth Failure:** `"❌ Connection failed: HTTP 401: Unauthorized"`
- **Network Error:** `"❌ Test failed: Network connection timeout"`

### 7. JQL Query Management
```
Default Query → User Customization → Syntax Validation → Performance Testing → Storage
```

**Default JQL Configuration:**
```javascript
function getDefaultJQL() {
  return 'project = CXPRODELIVERY ' +
         'AND issuetype in (Project, "Project (Standard Solution)") ' +
         'AND status in (HYPERCARE, "Order received", "Test system available", ' +
                        '"Project go-live/productive start", "System Design Assigned", ' +
                        '"Implementation Assigned", "Implementation Order Assigned", ' +
                        '"Test System Available (Implementation Order)", ' +
                        '"Handover Check Needed", "HYPERCARE (WITH CHECK)", ' +
                        '"LIVE SYSTEM AVAILABLE", "System Design Order Received", ' +
                        '"System Design Started", "Requirements Clarified", ' +
                        '"Implementation Started") ' +
         'AND "Technical Project Manager" in (currentUser())';
}
```

**Custom JQL Features:**
- **Syntax Help:** Built-in examples and documentation
- **Validation:** Test query against Jira API during save
- **Fallback:** Revert to default if custom query fails
- **Performance:** Warn about slow queries (>1000 results)

### 8. Settings UI Components

#### 8.1 Settings Card Structure
```
Header Section → Input Fields → Help Text → Action Buttons → Current Status
```

**Card Layout:**
```javascript
function buildSettingsCard(isFirstTime) {
  var card = CardService.newCardBuilder();
  
  // Header with context-aware message
  var headerSection = createHeaderSection(isFirstTime);
  
  // Input fields with current values  
  var settingsSection = createSettingsInputs(getUserSettings());
  
  // Action buttons (Save, Test, Back)
  var buttonSection = createActionButtons(isFirstTime);
  
  // Current configuration display (if exists)
  var statusSection = createCurrentSettingsDisplay();
  
  card.addSection(headerSection)
      .addSection(settingsSection) 
      .addSection(buttonSection)
      .addSection(statusSection);
      
  return card.build();
}
```

#### 8.2 Dynamic Help System
```javascript
// Context-sensitive help based on field focus
var tokenHelp = CardService.newTextParagraph()
  .setText("📖 How to get API token:\n" +
           "1. Go to Jira → Profile → Manage Account → Security\n" +
           "2. Create API token\n" + 
           "3. Copy the generated token (NOT your password)\n\n" +
           "🔒 Security: Current token is masked for security. " +
           "Leave unchanged to keep current token.");
```

### 9. Error Handling & Recovery

#### 9.1 Storage Failures
```javascript
try {
  setUserSettings(settings);
  return successNotification("Settings saved successfully!");
} catch (error) {
  console.error("Settings save failed:", error);
  return errorNotification("Failed to save settings: " + error.message);
}
```

#### 9.2 Connectivity Issues
- **DNS Resolution:** Handle invalid Jira URLs
- **Firewall/Proxy:** Detect corporate network restrictions  
- **SSL/TLS:** Handle certificate validation errors
- **Rate Limiting:** Manage API quota exceeded responses

#### 9.3 Graceful Degradation
```javascript
// If settings load fails, provide fallback functionality
function getUserSettings() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    var settingsJson = userProperties.getProperty('JIRA_SETTINGS');
    
    if (settingsJson) {
      return JSON.parse(settingsJson);
    } else {
      return getDefaultSettings(); // Fallback to defaults
    }
  } catch (error) {
    console.error("Error getting user settings:", error);
    return getDefaultSettings(); // Fallback on any error
  }
}
```

## Advanced Features

### 1. Settings Migration
```javascript
function migrateSettingsIfNeeded(settings) {
  // Handle old format settings
  if (settings.version !== CURRENT_SETTINGS_VERSION) {
    return migrateToCurrentVersion(settings);
  }
  return settings;
}
```

### 2. Bulk Configuration
- **Team Deployment:** Export/import settings for team distribution
- **Environment Switching:** Dev/Test/Prod Jira instance switching
- **Backup/Restore:** Settings backup for disaster recovery

### 3. Usage Analytics
```javascript
// Track settings usage for optimization
var analyticsData = {
  lastUsed: new Date().toISOString(),
  jiraCallsCount: incrementCounter('jira_api_calls'),
  errorCount: getErrorCount(),
  settingsVersion: CURRENT_SETTINGS_VERSION
};
```

## Security Best Practices

### 1. Token Management
- **Rotation Policy:** Remind users to rotate tokens periodically
- **Expiration Detection:** Handle expired token scenarios
- **Minimal Permissions:** Recommend read-only Jira permissions

### 2. Data Protection
- **No Logging:** Never log sensitive credentials
- **Secure Transmission:** HTTPS-only API calls
- **Access Auditing:** Log settings access for security review

### 3. Compliance Features
- **Data Retention:** Configurable settings retention periods
- **Export Controls:** Secure settings export functionality
- **Audit Trails:** Track all settings modifications

## Performance Optimization

### 1. Caching Strategy
```javascript
// Cache settings during execution to avoid repeated property reads
var settingsCache = null;
function getCachedSettings() {
  if (!settingsCache) {
    settingsCache = getUserSettings();
  }
  return settingsCache;
}
```

### 2. Lazy Loading
- **On-Demand Validation:** Only validate settings when needed
- **Progressive Loading:** Load basic settings first, advanced later
- **Background Sync:** Refresh cached settings in background

### 3. Resource Management
- **Property Limits:** Respect Google Apps Script property size limits
- **Memory Usage:** Optimize JSON parsing for large settings
- **Network Efficiency:** Batch validation calls where possible
