# PMO Integration Replacement Plan

## 🎯 Complete System Replacement

### Current Workflow (TO BE REMOVED)
```
Jira Ticket Selection → Create Local Folder Structure → Save Attachments
Example: "Jira Attachments/CXPRODELIVERY/CXPRODELIVERY-6605/"
```

### New Replacement Workflow (ONLY METHOD)
```
Jira Ticket Selection → PMO Webhook Lookup/Create → Use PMO Folder → Save Attachments
Example: Direct save to PMO-managed project folder (created if not exists)
```

### Benefits of Complete Replacement
- **Single Source of Truth**: PMO system is the only folder authority
- **Eliminated Folder Duplication**: No more local folder creation
- **Mandatory PMO Integration**: All projects use centralized PMO folders
- **Simplified Architecture**: Remove complex folder creation logic
- **Team Standardization**: Everyone uses the same PMO-managed folders

## 🔧 PMO Webhook Behavior

The PMO webhook provides both **lookup** and **auto-creation** functionality:

### Existing Folder Scenario
```
POST {"text": "CXPRODELIVERY-6605"} 
→ Response: [{"folderid": "1oKM_fKM4LG99ZQpSFmEY_sWciuAq-cgp"}]
→ Result: Immediate folder ID return
```

### Non-Existing Folder Scenario  
```
POST {"text": "CXPRODELIVERY-9999"} 
→ PMO creates folder automatically
→ First response: [{"folderid": undefined}] (rare timing issue)
→ Retry same POST request
→ Second response: [{"folderid": "1oKM_fKM4LG99ZQpSFmEY_sWciuAq-NEW"}]
→ Result: Newly created folder ID
```

### Key Behaviors
- **Automatic Creation**: PMO creates project folders if they don't exist
- **Timing Issue**: First request may return `undefined` folderid during creation
- **Retry Logic Required**: Second identical request returns valid folder ID
- **No Manual Creation**: System never needs to create folders manually

## 🔄 Technical Implementation Plan

### Phase 1: Core PMO Integration (High Priority)

#### 1.1 PMO Webhook Function with Auto-Creation and Retry Logic
```javascript
function getPMOProjectFolder(ticketKey) {
  try {
    var settings = getUserSettings();
    var webhookUrl = settings.pmoWebhookUrl || 'https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299';
    var maxRetries = settings.pmoRetryAttempts || 2;
    
    console.log("PMO webhook lookup/create for ticket:", ticketKey);
    
    for (var attempt = 1; attempt <= maxRetries; attempt++) {
      console.log("PMO request attempt", attempt, "of", maxRetries);
      
      var response = UrlFetchApp.fetch(webhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({"text": ticketKey}),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        
        if (data && data.length > 0) {
          var folderid = data[0].folderid;
          
          if (folderid && folderid !== 'undefined') {
            console.log("PMO webhook returned folder ID:", folderid);
            return {
              success: true,
              folderId: folderid,
              created: attempt > 1 // True if folder was created on this request
            };
          } else {
            console.log("PMO webhook returned undefined folder ID on attempt", attempt);
            
            if (attempt === maxRetries) {
              return {
                success: false,
                error: 'PMO webhook returned undefined folder ID after ' + maxRetries + ' attempts'
              };
            }
            
            // Wait briefly before retry (folder creation might need time)
            Utilities.sleep(1000); // 1 second delay
            continue;
          }
        } else {
          return {
            success: false,
            error: 'PMO webhook returned invalid response format'
          };
        }
      } else {
        return {
          success: false,
          error: 'PMO webhook HTTP error: ' + response.getResponseCode() + ' - ' + response.getContentText()
        };
      }
    }
    
  } catch (error) {
    return {
      success: false,
      error: 'PMO webhook network error: ' + error.message
    };
  }
}
```

#### 1.2 Safe Folder Access Function
```javascript
function getFolderByIdSafely(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    
    // Test folder access by trying to get name
    var folderName = folder.getName();
    
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

#### 1.3 Replaced Save Function with Auto-Creation Support
```javascript
function saveSelectedAttachmentsToGDrive(e) {
  // ... existing code for attachment processing ...
  
  // PMO Integration Logic (ONLY METHOD) - Lookup/Create folder
  console.log("PMO lookup/create for:", finalTicket);
  
  var pmoResult = getPMOProjectFolder(finalTicket);
  
  if (!pmoResult.success) {
    // PMO webhook failed - this is an error condition
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("❌ PMO folder lookup/creation failed for " + finalTicket + ": " + pmoResult.error))
      .build();
  }
  
  console.log("PMO returned folder ID:", pmoResult.folderId, "- Created:", pmoResult.created || false);
  
  var folderResult = getFolderByIdSafely(pmoResult.folderId);
  
  if (!folderResult.success) {
    // Folder access failed - this is an error condition  
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("❌ Cannot access PMO project folder (ID: " + pmoResult.folderId + "): " + folderResult.error))
      .build();
  }
  
  var targetFolder = folderResult.folder;
  console.log("Using PMO folder:", folderResult.name);
  
  // ... continue with existing attachment saving logic to targetFolder ...
  
  // Success notification with creation info
  var notificationText = "✅ Saved " + savedCount + " attachments to PMO project folder: " + folderResult.name;
  
  if (pmoResult.created) {
    notificationText = "✅ Created PMO folder and saved " + savedCount + " attachments: " + folderResult.name;
  }
  
  if (skippedCount > 0) {
    notificationText += " (⚠️ " + skippedCount + " duplicates skipped)";
  }
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(notificationText))
    .build();
}
```

### Phase 2: Settings Enhancement (Medium Priority)

#### 2.1 Simplified Settings Object
```javascript
// Updated settings structure - PMO is now mandatory
{
  // Existing Jira settings
  jiraUrl: "https://support.transporeon.com",
  jiraToken: "encrypted_api_token_here", 
  customJql: "project = CXPRODELIVERY AND...",
  savedAt: "2024-10-26T14:30:22.000Z",
  
  // PMO Integration Settings (MANDATORY)
  pmoWebhookUrl: "https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299",
  pmoTimeout: 10000,           // 10 second timeout (more generous since it's critical)
  pmoRetryAttempts: 2          // 2 attempts: handle undefined folderid during auto-creation
}
```

#### 2.2 Simplified Settings UI
```javascript
function buildSettingsCard(isFirstTime) {
  // ... existing Jira settings sections ...
  
  // PMO Integration Section (MANDATORY)
  var pmoSection = CardService.newCardSection()
    .setHeader("📁 PMO Project Folders (Required)");
  
  // PMO Webhook URL
  var pmoUrlInput = CardService.newTextInput()
    .setFieldName("pmoWebhookUrl")
    .setTitle("PMO Webhook URL")
    .setHint("URL for PMO project folder lookup (required)")
    .setValue(settings.pmoWebhookUrl || getDefaultPMOWebhookUrl());
  pmoSection.addWidget(pmoUrlInput);
  
  // PMO Timeout Setting
  var pmoTimeoutInput = CardService.newTextInput()
    .setFieldName("pmoTimeout")
    .setTitle("PMO Timeout (ms)")
    .setHint("Timeout for PMO webhook calls (default: 10000)")
    .setValue(String(settings.pmoTimeout || 10000));
  pmoSection.addWidget(pmoTimeoutInput);
  
  // PMO Help Text
  var pmoHelp = CardService.newTextParagraph()
    .setText("📋 PMO Integration (Mandatory):\n" +
             "• All attachments saved to PMO-managed project folders\n" +
             "• PMO webhook automatically creates folders if they don't exist\n" +
             "• Retry logic handles folder creation timing issues\n" +
             "• Attachment save fails only if PMO webhook is unreachable");
  pmoSection.addWidget(pmoHelp);
  
  // Test PMO Connection Button
  var testPMOButton = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("🧪 Test PMO Connection")
      .setOnClickAction(CardService.newAction().setFunctionName("testPMOConnection")));
  pmoSection.addWidget(testPMOButton);
  
  card.addSection(pmoSection);
  
  // ... rest of existing settings card ...
}
```

#### 2.3 PMO Connection Testing
```javascript
function testPMOConnection(e) {
  console.log("=== TEST PMO CONNECTION ===");
  
  try {
    var settings = getUserSettings();
    var webhookUrl = settings.pmoWebhookUrl || getDefaultPMOWebhookUrl();
    
    if (!webhookUrl) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ PMO webhook URL not configured. Please set URL first."))
        .build();
    }
    
    // Test with a known ticket (could be made configurable)
    var testTicket = "CXPRODELIVERY-TEST";
    var result = getPMOProjectFolder(testTicket);
    
    if (result.success) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("✅ PMO webhook connection successful! Found folder ID: " + result.folderId))
        .build();
    } else {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ PMO webhook test failed: " + result.error))
        .build();
    }
    
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("❌ PMO connection test failed: " + error.message))
      .build();
  }
}
```

### Phase 3: UX Improvements (Lower Priority)

#### 3.1 Loading Indicators
- Add loading message during PMO lookup: "🔍 Looking up PMO project folder..."
- Show folder source in success notification
- Display folder name/path after successful PMO lookup

#### 3.2 Enhanced Error Messages
```javascript
function getPMOErrorMessage(error, ticketKey) {
  var baseMessage = "PMO folder lookup failed for " + ticketKey + ": ";
  
  if (error.includes("timeout")) {
    return baseMessage + "Connection timeout. Using local folder instead.";
  } else if (error.includes("404")) {
    return baseMessage + "Project not found in PMO system. Using local folder instead.";
  } else if (error.includes("403")) {
    return baseMessage + "Access denied. Check permissions or use local folder.";
  } else {
    return baseMessage + error + ". Using local folder instead.";
  }
}
```

#### 3.3 Folder Confirmation Display
```javascript
// After successful PMO folder lookup, show confirmation
var folderInfoSection = CardService.newCardSection()
  .setHeader("📁 Target Folder");

var folderInfo = CardService.newTextParagraph()
  .setText("✅ Using PMO folder: " + folderResult.name + "\n" +
           "📊 Folder ID: " + pmoResult.folderId);
folderInfoSection.addWidget(folderInfo);
```

## 🔧 Implementation Details

### Modified Functions

#### Core Functions to Update:
1. **`saveSelectedAttachmentsToGDrive()`** - Add PMO lookup logic
2. **`buildSettingsCard()`** - Add PMO settings section  
3. **`saveSettings()`** - Handle PMO settings persistence
4. **`getUserSettings()`** - Include PMO defaults for new users

#### New Functions to Add:
1. **`getPMOProjectFolder(ticketKey)`** - PMO webhook integration
2. **`getFolderByIdSafely(folderId)`** - Safe folder access with error handling
3. **`testPMOConnection()`** - PMO connectivity testing
4. **`getDefaultPMOWebhookUrl()`** - Default PMO webhook URL provider

### Error Handling Strategy

#### Failure Scenarios & Responses:
| Failure Type | User Impact | System Response |
|--------------|-------------|-----------------|
| PMO webhook down | Save operation fails | Clear error message, ask user to retry later |
| Undefined folder ID | Save operation fails | Retry logic (2 attempts), then error if still undefined |
| Folder access denied | Save operation fails | Error: "Cannot access PMO project folder" |
| Network timeout | Save operation fails | Error: "PMO lookup timeout, please try again" |
| Invalid response format | Save operation fails | Error: "PMO webhook returned invalid response" |

#### Error Handling Logic:
```javascript
var pmoResult = getPMOProjectFolder(finalTicket);

if (!pmoResult.success) {
  // PMO lookup failed - operation cannot continue
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("❌ Cannot save attachments: " + pmoResult.error))
    .build();
}

// Continue only if PMO lookup succeeded
var folderResult = getFolderByIdSafely(pmoResult.folderId);
if (!folderResult.success) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("❌ Cannot access PMO folder: " + folderResult.error))
    .build();
}
```

## 🔄 System Replacement Strategy

### Complete Migration Required
- **Breaking Change**: Existing folder creation logic completely removed
- **PMO Dependency**: System now requires PMO webhook for all operations
- **Settings Update**: Users must configure PMO webhook URL
- **No Backwards Compatibility**: Old folder creation method no longer available

### User Migration Steps
1. **Configure PMO Settings**: Users must set PMO webhook URL
2. **Test PMO Connection**: Verify webhook connectivity before use
3. **Update Workflows**: All existing workflows now use PMO folders only
4. **Cleanup**: Existing "Jira Attachments" folders can be manually archived

### Migration Validation
```javascript
function validatePMOSettings() {
  var settings = getUserSettings();
  
  if (!settings.pmoWebhookUrl) {
    return {
      valid: false,
      error: "PMO webhook URL is required for system operation"
    };
  }
  
  return { valid: true };
}
```

## 📊 Testing Plan

### Test Scenarios

#### 1. PMO Integration Success Cases
- ✅ Existing folder: Immediate folder ID return and successful save
- ✅ New folder: Auto-creation with retry logic, successful save
- ✅ Multiple users accessing same PMO folder
- ✅ PMO folder with existing attachments (duplicate handling)
- ✅ Large attachments saved to PMO folders
- ✅ Various file types saved successfully
- ✅ Undefined folderid → successful retry → valid folder ID

#### 2. PMO Integration Failure Cases  
- ❌ PMO webhook unreachable → Save operation fails with retry suggestion
- ❌ Undefined folderid after max retries → Save fails with creation error
- ❌ Folder ID returned but folder inaccessible → Save fails with permission error
- ❌ Network timeout → Save fails with timeout error
- ❌ Invalid PMO response format → Save fails with format error
- ❌ Webhook returns HTTP error codes → Save fails with HTTP error

#### 3. Settings & Configuration
- ⚙️ PMO webhook URL configuration (required)
- ⚙️ PMO timeout settings
- ⚙️ PMO connection testing with real tickets
- ⚙️ Settings validation and error handling

#### 4. User Experience
- 🎭 Loading states during PMO lookup
- 🎭 Clear error messages when PMO fails
- 🎭 Success confirmations with PMO folder information
- 🎭 Retry mechanisms for failed operations

### Testing Data Requirements
```javascript
// Test tickets needed
var testTickets = {
  existingFolder: "CXPRODELIVERY-6605",        // Has existing PMO folder
  newFolder: "CXPRODELIVERY-9999",             // Will trigger PMO auto-creation
  invalidTicket: "INVALID-123",                // Invalid format
  longTicketName: "CXPRODELIVERY-VERY-LONG-NAME-TEST"  // Edge case naming
};

// Expected behaviors
var expectedBehaviors = {
  "CXPRODELIVERY-6605": {
    behavior: "immediate_return",
    folderid_defined: true,
    retry_needed: false
  },
  "CXPRODELIVERY-9999": {
    behavior: "auto_create", 
    folderid_defined: false,  // First request
    retry_needed: true        // Second request should succeed
  }
};
```

## 📚 Documentation Updates Required

### New Documentation
1. **`Documentation/pmo-integration-process.md`** - Complete PMO workflow documentation
2. **`Documentation/migration-guide.md`** - Migration guide for existing users
3. **`Documentation/troubleshooting-pmo.md`** - PMO-specific troubleshooting

### Updated Documentation
1. **`Documentation/smart-organization-process.md`** - Add PMO folder retrieval
2. **`Documentation/settings-management-process.md`** - Add PMO settings section
3. **`README.md`** - Update feature list with PMO integration
4. **`Documentation/README.md`** - Add PMO integration to index

## 🎯 Success Metrics

### Technical Success
- ✅ PMO lookup success rate > 95% for valid tickets
- ✅ Clear error handling for PMO failures (100% coverage)
- ✅ Response time < 10 seconds for PMO lookup (increased tolerance)
- ✅ Zero data loss on PMO failures (operation fails cleanly)
- ✅ Reliable retry mechanism for transient failures

### User Experience Success  
- ✅ Direct access to PMO-managed project folders
- ✅ Complete elimination of duplicate folder structures
- ✅ Clear feedback when PMO operations succeed/fail
- ✅ Intuitive PMO configuration and testing
- ✅ Helpful error messages with next steps

## 🚀 Implementation Timeline

### Week 1: Core Integration
- Implement PMO webhook functions
- Modify save function with PMO lookup
- Basic error handling and fallback

### Week 2: Settings Enhancement
- Add PMO settings to configuration UI
- Implement PMO connection testing
- Settings persistence and migration

### Week 3: UX & Polish
- Add loading indicators and better messaging
- Enhanced error handling and user feedback
- Comprehensive testing

### Week 4: Documentation & Deployment
- Update all documentation
- Final testing and validation
- Deployment preparation

This complete system replacement eliminates local folder creation in favor of mandatory PMO integration, ensuring all teams use centralized project folders managed by the PMO system.
