# PMO Integration Enhancement Plan

## 🎯 Enhancement Overview

### Current Workflow
```
Jira Ticket Selection → Create Local Folder Structure → Save Attachments
Example: "Jira Attachments/CXPRODELIVERY/CXPRODELIVERY-6605/"
```

### New Enhanced Workflow  
```
Jira Ticket Selection → PMO Webhook Lookup → Use Existing PMO Folder → Save Attachments
Example: Direct save to existing PMO-managed project folder
```

### Benefits
- **Centralized Folder Management**: Use existing PMO-managed project folders
- **Reduced Folder Proliferation**: Eliminate duplicate folder structures
- **Improved Organization**: Align with existing project management workflows
- **Team Collaboration**: Shared access to PMO-maintained project spaces

## 🔄 Technical Implementation Plan

### Phase 1: Core PMO Integration (High Priority)

#### 1.1 New PMO Webhook Function
```javascript
function getPMOProjectFolder(ticketKey) {
  try {
    var settings = getUserSettings();
    var webhookUrl = settings.pmoWebhookUrl || 'https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299';
    
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
      if (data && data.length > 0 && data[0].folderid) {
        return {
          success: true,
          folderId: data[0].folderid
        };
      }
    }
    
    return {
      success: false,
      error: 'No folder found for ticket: ' + ticketKey
    };
    
  } catch (error) {
    return {
      success: false,
      error: 'PMO webhook error: ' + error.message
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

#### 1.3 Enhanced Save Function
```javascript
function saveSelectedAttachmentsToGDrive(e) {
  // ... existing code for attachment processing ...
  
  // NEW: PMO Integration Logic
  var targetFolder = null;
  var folderSource = 'local'; // Track folder source for user feedback
  
  var settings = getUserSettings();
  
  if (settings.enablePMOIntegration) {
    console.log("Attempting PMO folder lookup for:", finalTicket);
    
    var pmoResult = getPMOProjectFolder(finalTicket);
    
    if (pmoResult.success) {
      var folderResult = getFolderByIdSafely(pmoResult.folderId);
      
      if (folderResult.success) {
        targetFolder = folderResult.folder;
        folderSource = 'pmo';
        console.log("Using PMO folder:", folderResult.name);
      } else {
        console.log("PMO folder access failed:", folderResult.error);
      }
    } else {
      console.log("PMO lookup failed:", pmoResult.error);
    }
  }
  
  // Fallback to local folder creation if PMO failed or disabled
  if (!targetFolder) {
    console.log("Falling back to local folder creation");
    var baseFolder = getOrCreateFolder("Jira Attachments");
    var projectFolder = getOrCreateFolder("CXPRODELIVERY", baseFolder);
    targetFolder = getOrCreateFolder(finalTicket, projectFolder);
    folderSource = 'local';
  }
  
  // ... continue with existing attachment saving logic ...
  
  // Enhanced notification with folder source info
  var notificationText = "✅ Saved " + savedCount + " attachments";
  if (folderSource === 'pmo') {
    notificationText += " to PMO project folder";
  } else {
    notificationText += " to CXPRODELIVERY/" + finalTicket;
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

#### 2.1 Extended Settings Object
```javascript
// Updated settings structure
{
  // Existing Jira settings
  jiraUrl: "https://support.transporeon.com",
  jiraToken: "encrypted_api_token_here", 
  customJql: "project = CXPRODELIVERY AND...",
  savedAt: "2024-10-26T14:30:22.000Z",
  
  // NEW: PMO Integration Settings
  enablePMOIntegration: false,  // Disabled by default for backwards compatibility
  pmoWebhookUrl: "https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299",
  pmoTimeout: 5000,            // 5 second timeout
  pmoFallbackToLocal: true,    // Always fallback to local folder creation
  pmoRetryAttempts: 2          // Number of retry attempts for PMO webhook
}
```

#### 2.2 Settings UI Enhancement
```javascript
function buildSettingsCard(isFirstTime) {
  // ... existing settings sections ...
  
  // NEW: PMO Integration Section
  var pmoSection = CardService.newCardSection()
    .setHeader("📁 PMO Folder Integration")
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(2);
  
  // Enable/Disable PMO Integration
  var enablePMOToggle = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("enablePMOIntegration")
    .addItem("Enable PMO folder lookup", "true", settings.enablePMOIntegration || false);
  pmoSection.addWidget(enablePMOToggle);
  
  // PMO Webhook URL
  var pmoUrlInput = CardService.newTextInput()
    .setFieldName("pmoWebhookUrl")
    .setTitle("PMO Webhook URL")
    .setHint("URL for PMO project folder lookup")
    .setValue(settings.pmoWebhookUrl || getDefaultPMOWebhookUrl());
  pmoSection.addWidget(pmoUrlInput);
  
  // PMO Help Text
  var pmoHelp = CardService.newTextParagraph()
    .setText("💡 PMO Integration:\n" +
             "• Uses existing PMO-managed project folders\n" +
             "• Falls back to local folder creation if PMO lookup fails\n" +
             "• Reduces folder duplication across teams");
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
    
    if (!settings.enablePMOIntegration) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ PMO integration is disabled. Enable it first."))
        .build();
    }
    
    // Test with a known ticket (could be made configurable)
    var testTicket = "CXPRODELIVERY-TEST";
    var result = getPMOProjectFolder(testTicket);
    
    if (result.success) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("✅ PMO webhook connection successful!"))
        .build();
    } else {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("⚠️ PMO webhook accessible but no test folder found: " + result.error))
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
| PMO webhook down | Transparent fallback | Use local folder creation, notify user |
| Ticket not in PMO | Transparent fallback | Create local folder, log for analysis |
| Folder access denied | Transparent fallback | Create local folder, warn user |
| Network timeout | Transparent fallback | Use cached folder or create local |

#### Fallback Logic:
```javascript
if (pmoLookupFailed && settings.pmoFallbackToLocal) {
  // Seamless fallback to existing folder creation logic
  console.log("PMO lookup failed, falling back to local folder creation");
  targetFolder = createLocalProjectFolder(finalTicket);
  folderSource = 'local_fallback';
}
```

## 🔄 Migration Strategy

### Backwards Compatibility
- **Default State**: PMO integration disabled for existing users
- **Opt-in Adoption**: Users can enable PMO integration through settings
- **Graceful Fallback**: Always maintain local folder creation as backup
- **Settings Migration**: Automatically add PMO settings with safe defaults

### User Communication
```javascript
// One-time notification for existing users about new PMO feature
if (isNewPMOFeature(settings)) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("🆕 New PMO Integration available! Check Settings to enable PMO project folder lookup."))
    .build();
}
```

## 📊 Testing Plan

### Test Scenarios

#### 1. PMO Integration Success Cases
- ✅ Ticket exists in PMO system, folder accessible
- ✅ Multiple users accessing same PMO folder
- ✅ PMO folder with existing attachments (duplicate handling)

#### 2. PMO Integration Failure Cases  
- ❌ Ticket not found in PMO system → Fallback to local
- ❌ PMO webhook unreachable → Fallback to local
- ❌ Folder ID returned but folder inaccessible → Fallback to local
- ❌ Network timeout → Fallback to local

#### 3. Settings & Configuration
- ⚙️ Enable/disable PMO integration
- ⚙️ Custom PMO webhook URL
- ⚙️ PMO connection testing
- ⚙️ Settings persistence and migration

#### 4. User Experience
- 🎭 Loading states during PMO lookup
- 🎭 Error messages and fallback notifications
- 🎭 Success confirmations with folder information

### Testing Data Requirements
```javascript
// Test tickets needed
var testTickets = {
  withPMOFolder: "CXPRODELIVERY-6605",     // Has PMO folder
  withoutPMOFolder: "CXPRODELIVERY-9999",  // No PMO folder
  invalidTicket: "INVALID-123"              // Invalid format
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
- ✅ PMO lookup success rate > 95%
- ✅ Fallback mechanism works 100% of time on PMO failure
- ✅ No breaking changes for existing users
- ✅ Response time < 5 seconds for PMO lookup

### User Experience Success  
- ✅ Seamless folder access to PMO-managed projects
- ✅ Reduced duplicate folder structures
- ✅ Clear feedback on folder source (PMO vs local)
- ✅ Easy PMO integration configuration

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

This enhancement maintains full backwards compatibility while adding powerful PMO integration that aligns with existing project management workflows.
