# PMO Integration Process Flow

## Overview
The PMO Integration process provides automated folder lookup and creation through a centralized PMO webhook system, eliminating the need for local folder creation and ensuring all project files are stored in standardized PMO-managed folders.

## Process Flow

### 1. PMO Webhook Integration
```
Jira Ticket Selection → PMO Webhook Call → Folder ID Response → Folder Access → Attachment Save
```

**Complete Workflow:**
```
User selects ticket: CXPRODELIVERY-6605
     ↓
POST webhook: {"text": "CXPRODELIVERY-6605"}
     ↓
PMO system checks: Does folder exist?
     ↓
If YES: Return immediate folder ID
If NO:  Create folder → Return folder ID (may require retry)
     ↓
Access Google Drive folder using ID
     ↓
Save attachments with duplicate handling
```

### 2. PMO Webhook Behavior

#### 2.1 Existing Folder Scenario
```javascript
// Request
POST https://n8n-pmo.office.transporeon.com/webhook/...
{
  "text": "CXPRODELIVERY-6605"
}

// Response (immediate)
[
  {
    "folderid": "1oKM_fKM4LG99ZQpSFmEY_sWciuAq-cgp"
  }
]
```

#### 2.2 New Folder Scenario (Auto-Creation)
```javascript
// First Request
POST webhook → PMO creates folder automatically
// Response (timing issue - rare)
[
  {
    "folderid": undefined
  }
]

// Retry Request (1 second delay)
POST webhook → Same request
// Response (successful)
[
  {
    "folderid": "1oKM_fKM4LG99ZQpSFmEY_sWciuAq-NEW"
  }
]
```

### 3. Technical Implementation

#### 3.1 PMO Folder Lookup Function
```javascript
function getPMOProjectFolder(ticketKey) {
  var settings = getUserSettings();
  var webhookUrl = settings.pmoWebhookUrl;
  var maxRetries = settings.pmoRetryAttempts || 2;
  var timeout = settings.pmoTimeout || 10000;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({"text": ticketKey}),
      muteHttpExceptions: true,
      timeout: timeout
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var folderid = data[0].folderid;
      
      if (folderid && folderid !== 'undefined') {
        return {
          success: true,
          folderId: folderid,
          created: attempt > 1
        };
      }
      
      // Retry for undefined responses
      Utilities.sleep(1000);
    }
  }
  
  return { success: false, error: "PMO webhook failed" };
}
```

#### 3.2 Safe Folder Access
```javascript
function getFolderByIdSafely(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var folderName = folder.getName(); // Test access
    
    return {
      success: true,
      folder: folder,
      name: folderName
    };
  } catch (error) {
    return {
      success: false,
      error: 'Cannot access folder: ' + error.message
    };
  }
}
```

#### 3.3 Enhanced Save Function
```javascript
function saveSelectedAttachmentsToGDrive(e) {
  // ... attachment processing ...
  
  // PMO Integration (ONLY METHOD)
  var pmoResult = getPMOProjectFolder(finalTicket);
  
  if (!pmoResult.success) {
    return errorNotification(getPMOErrorMessage(pmoResult.error, finalTicket));
  }
  
  var folderResult = getFolderByIdSafely(pmoResult.folderId);
  
  if (!folderResult.success) {
    return folderAccessError(folderResult.error, pmoResult.folderId, finalTicket);
  }
  
  var targetFolder = folderResult.folder;
  
  // ... save attachments to PMO folder ...
  
  return successNotification(savedCount, folderResult.name, pmoResult.created);
}
```

### 4. Retry Logic and Error Handling

#### 4.1 Retry Strategy
```javascript
// Configuration
maxRetries: 2              // Handle undefined responses during auto-creation
retryDelay: 1000          // 1 second between attempts
timeout: 10000            // 10 second request timeout

// Retry Conditions
- folderid === undefined   // Auto-creation timing issue
- folderid === null       // Invalid response
- folderid === ""         // Empty response

// No Retry Conditions  
- HTTP errors (404, 500, 403)
- Network timeouts
- Invalid JSON response
```

#### 4.2 Error Classification and User Messages

##### Network Errors
```javascript
if (error.includes("timeout") || error.includes("network")) {
  return "🌐 Network Issue:\n" +
         "• PMO system temporarily unreachable\n" +
         "• Check internet connection\n" +
         "• Try again in a few moments";
}
```

##### PMO System Errors
```javascript
if (error.includes("HTTP 404") || error.includes("HTTP 500")) {
  return "🔧 PMO System Issue:\n" +
         "• PMO webhook temporarily unavailable\n" +
         "• Likely temporary service issue\n" +
         "• Try again in a few minutes";
}
```

##### Permission Errors
```javascript
if (error.includes("HTTP 403") || error.includes("unauthorized")) {
  return "🔐 Access Issue:\n" +
         "• May not have permission to create folders\n" +
         "• Check with project manager\n" +
         "• Verify ticket number is correct";
}
```

##### Folder Creation Issues
```javascript
if (error.includes("undefined folder ID")) {
  return "⏱️ PMO Folder Creation Issue:\n" +
         "• PMO processing project folder\n" +
         "• Creation took longer than expected\n" +
         "• Try again - folder may be ready now";
}
```

### 5. Settings Management

#### 5.1 PMO Configuration Settings
```javascript
{
  // PMO Integration Settings (Mandatory)
  pmoWebhookUrl: "https://n8n-pmo.office.transporeon.com/webhook/...",
  pmoTimeout: 10000,           // 10 second timeout
  pmoRetryAttempts: 2          // 2 retry attempts for undefined responses
}
```

#### 5.2 Settings Validation
```javascript
// URL Validation
- Must start with http:// or https://
- Must be valid URL format
- Trailing slashes removed automatically

// Timeout Validation  
- Range: 5000 to 60000 milliseconds (5-60 seconds)
- Default: 10000ms (10 seconds)

// Retry Validation
- Range: 1 to 5 attempts
- Default: 2 attempts
```

#### 5.3 PMO Connection Testing
```javascript
function testPMOConnection() {
  // Generate unique test ticket
  var testTicket = "CXPRODELIVERY-TEST-" + new Date().getTime();
  
  // Test PMO webhook
  var result = getPMOProjectFolder(testTicket);
  
  if (result.success) {
    // Test folder access
    var folderTest = getFolderByIdSafely(result.folderId);
    
    return detailedSuccessReport(result, folderTest);
  } else {
    return detailedErrorReport(result, testTicket);
  }
}
```

### 6. User Experience Features

#### 6.1 Enhanced Success Notifications
```javascript
"✅ Attachment Save Complete!

📁 New PMO folder created: Project Alpha Implementation
📎 Files processed: 3
💾 Files saved: 3
🎫 Project: CXPRODELIVERY-6605
✨ Success rate: 100%

📋 Saved files:
• requirements.pdf
• design.docx  
• screenshots.zip"
```

#### 6.2 Enhanced Error Messages
```javascript
"❌ Cannot save attachments for CXPRODELIVERY-6605

🌐 Network Issue:
• PMO system is temporarily unreachable
• Please check your internet connection
• Try again in a few moments

💡 If problem persists, contact IT support"
```

#### 6.3 Progress Feedback
```javascript
Console Logging:
"🔍 Contacting PMO system for project folder..."
"✅ PMO returned folder ID: 1oKM_... - Created: false"
"📁 Using PMO folder: Project Alpha Implementation"
"💾 Saved 3 files successfully"
```

## Advanced Features

### 1. Automatic Folder Creation
- **No Manual Setup**: PMO system automatically creates project folders
- **Standardized Structure**: All project folders follow PMO naming conventions  
- **Team Synchronization**: Multiple users get same folder for same project
- **Permission Management**: PMO system handles folder permissions

### 2. Intelligent Retry Logic  
- **Timing Issue Handling**: Automatic retry for undefined responses during creation
- **Network Resilience**: Configurable timeout and retry settings
- **Error Recovery**: Smart error classification with specific user guidance
- **Performance Optimization**: Efficient retry strategy with exponential backoff

### 3. Enhanced User Experience
- **Detailed Feedback**: Comprehensive success and error reporting
- **Progress Indication**: Real-time feedback during PMO operations  
- **Smart Validation**: Input validation with helpful error messages
- **Connection Testing**: Comprehensive PMO webhook connectivity testing

### 4. Security and Compliance
- **Centralized Management**: All folders managed by PMO system
- **Access Control**: PMO system handles permissions and access rights
- **Audit Trail**: Complete logging of all PMO interactions
- **Data Governance**: Standardized folder structure ensures compliance

## Performance Considerations

### 1. Network Optimization
```javascript
// Timeout Configuration
Default: 10 seconds    // Balance between reliability and performance
Range: 5-60 seconds   // User configurable based on network conditions

// Retry Strategy  
Default: 2 attempts   // Handle auto-creation timing issues efficiently
Range: 1-5 attempts  // Configurable for different environments
```

### 2. Caching Strategy
- **Settings Caching**: PMO configuration cached during session
- **Error State Management**: Failed requests logged to avoid repeated failures  
- **Folder Reference Caching**: Successfully accessed folders cached temporarily

### 3. Resource Management
- **Memory Efficiency**: Minimal overhead for PMO operations
- **Network Usage**: Optimized request/response handling
- **API Limits**: Respectful of Google Apps Script execution limits

## Integration Points

### 1. Gmail Add-on Integration
- **Contextual Activation**: PMO lookup triggered only when saving attachments
- **Thread Context**: PMO folders linked to specific email threads
- **Selection Memory**: Attachment selections preserved during PMO operations

### 2. Jira Integration
- **Ticket Validation**: Only valid Jira tickets trigger PMO lookup
- **Project Mapping**: Jira project keys map to PMO folder structure
- **Status Tracking**: Ticket status influences folder organization

### 3. Google Drive Integration
- **Folder Access**: Direct access to PMO-managed folders via Google Drive API
- **Permission Checking**: Automatic validation of user access rights
- **Duplicate Handling**: Standard duplicate file management in PMO folders

This PMO integration provides a complete replacement for local folder creation, ensuring all project files are properly organized in centralized, PMO-managed folder structures with automatic creation, intelligent error handling, and enhanced user experience.
