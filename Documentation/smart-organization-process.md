# PMO Smart Organization Process Flow

## Overview
The PMO Smart Organization feature automatically connects to centralized PMO-managed project folders, eliminating local folder creation and ensuring all attachments are stored in standardized PMO folder structures.

## Process Flow

### 1. PMO Folder Resolution
```
Ticket Selection → PMO Webhook Lookup → Retrieve Existing PMO Folder → Direct Access
```

**PMO Integration Structure:**
```
PMO System/
└── [PMO-Managed Project Folders]
    └── CXPRODELIVERY-{number}/          (PMO-created project folder)
        ├── document1.pdf
        ├── spreadsheet1.xlsx
        └── image1.png
```

**Technical Implementation:**
```javascript
var pmoResult = getPMOProjectFolder(finalTicket);    // PMO webhook call
var folderResult = getFolderByIdSafely(pmoResult.folderId);  // Direct folder access
var targetFolder = folderResult.folder;              // Use PMO folder
```

### 2. PMO Webhook Integration
```
Ticket Input → PMO API Call → Folder Lookup/Creation → Folder ID Response → Direct Access
```

**PMO Integration Process:**

#### 2.1 PMO Webhook Call
- **Endpoint:** Configurable PMO webhook URL
- **Method:** POST request with ticket key
- **Payload:** `{"text": "CXPRODELIVERY-6310"}`
- **Function:** Automatic folder lookup and creation

#### 2.2 PMO Response Handling
- **Existing Folder:** Immediate folder ID response
- **New Folder:** Auto-creation with potential retry for undefined responses
- **Validation:** Folder ID validation and access verification
- **Caching:** Folder reference stored for session duration

#### 2.3 Direct Folder Access
- **Method:** Google Drive API access using folder ID
- **Verification:** Folder name and permissions validation
- **Integration:** Seamless integration with attachment saving process
- **Error Handling:** Comprehensive error reporting for access issues

### 3. Dynamic Ticket Resolution
```
User Input → Ticket Format Validation → Ticket Key Generation → Folder Naming
```

**Input Sources:**
1. **Dropdown Selection:** Pre-formatted ticket key from Jira
2. **Manual Entry:** Numeric input converted to full key
3. **Direct Entry:** Full ticket key validation

**Ticket Resolution Logic:**
```javascript
var finalTicket;
if (selectedTicket && selectedTicket !== "manual") {
  finalTicket = selectedTicket; // From dropdown: "CXPRODELIVERY-6310"
} else if (manualTicketNumber) {
  finalTicket = "CXPRODELIVERY-" + manualTicketNumber; // From manual: "6310" → "CXPRODELIVERY-6310"
} else if (fullTicket) {
  finalTicket = fullTicket; // Direct entry: "CXPRODELIVERY-6310"
}
```

### 4. PMO Folder Validation and Access
```
PMO Response → Folder ID Validation → Google Drive Access → Folder Verification → Ready for Use
```

**Technical Implementation:**
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

**Validation Features:**
- **ID Verification:** Validates folder ID format and accessibility
- **Access Testing:** Verifies read/write permissions
- **Error Handling:** Detailed error reporting for access issues
- **Name Retrieval:** Provides folder name for user feedback

### 5. Path Generation Logic
```
Ticket Input → Normalize Format → Generate Full Path → Validate Characters
```

**Path Components:**
```javascript
// Base path construction
var basePath = "Jira Attachments/CXPRODELIVERY/" + finalTicket + "/";

// Example results:
// Manual input "6310" → "Jira Attachments/CXPRODELIVERY/CXPRODELIVERY-6310/"
// Dropdown selection → "Jira Attachments/CXPRODELIVERY/CXPRODELIVERY-6310/"
// Direct entry → "Jira Attachments/CXPRODELIVERY/CXPRODELIVERY-6310/"
```

**Character Validation:**
- Remove invalid filesystem characters
- Handle spaces and special characters
- Ensure cross-platform compatibility

## PMO Organization Benefits

### 1. Centralized Management
- **Single Source of Truth:** PMO system manages all project folders
- **Standardized Structure:** Consistent folder organization across all projects  
- **Automatic Creation:** Folders created automatically when needed
- **Permission Control:** PMO system handles access rights and permissions

### 2. Team Collaboration
- **Shared Access:** All team members access same PMO-managed folders
- **Real-time Sync:** Changes visible to all project stakeholders
- **Cross-platform Access:** Available through Google Drive, web, and mobile
- **Version Control:** Google Drive versioning built into PMO folders

### 3. Enterprise Integration
- **PMO Workflow:** Seamlessly integrates with existing PMO processes
- **Audit Trail:** Complete tracking of folder creation and access
- **Compliance Ready:** Meets organizational data governance requirements
- **Backup Included:** PMO folders included in organizational backup systems

## PMO Error Handling

### PMO Webhook Failures
```javascript
if (!pmoResult.success) {
  var userFriendlyError = getPMOErrorMessage(pmoResult.error, finalTicket);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(userFriendlyError))
    .build();
}
```

**PMO Error Categories:**
- **Network Issues:** Timeout, connectivity problems
- **System Errors:** PMO webhook unavailable (404, 500)
- **Access Errors:** Permission denied (403, unauthorized)
- **Timing Issues:** Undefined folder ID during auto-creation

### Folder Access Validation
```javascript
if (!folderResult.success) {
  var folderError = "❌ Cannot access PMO project folder for " + finalTicket;
  folderError += "\n🔗 Folder ID: " + pmoResult.folderId;
  folderError += "\n📋 Issue: " + folderResult.error;
  folderError += "\n💡 Contact project manager for folder permissions";
  return errorNotification(folderError);
}
```

**Folder Access Issues:**
- **Permission Denied:** User lacks access to PMO folder
- **Folder Deleted:** PMO folder removed or moved
- **Drive Limits:** Google Drive quota or API limits exceeded

### PMO System Recovery
- **Retry Logic:** Automatic retry for undefined folder responses
- **Timeout Handling:** Configurable timeout with fallback messaging
- **Error Classification:** User-friendly messages for different error types
- **No Local Fallback:** PMO integration is mandatory - no local folder creation

## PMO Performance Considerations

### PMO Webhook Optimization
```javascript
// Configurable timeout and retry settings
var settings = {
  pmoTimeout: 10000,        // 10 second timeout
  pmoRetryAttempts: 2       // 2 retry attempts for undefined responses
};

// Efficient retry logic with delay
for (var attempt = 1; attempt <= maxRetries; attempt++) {
  var result = callPMOWebhook(ticketKey);
  if (result.success && result.folderId !== 'undefined') {
    return result;
  }
  Utilities.sleep(1000); // 1 second delay between retries
}
```

### Folder Reference Caching
- **Session Caching:** PMO folder references cached during session
- **Access Validation:** Cached folders validated before use
- **Memory Management:** Cache cleared between email contexts

### PMO Integration Monitoring
- **Response Times:** Track PMO webhook performance
- **Success Rates:** Monitor PMO folder creation success
- **Error Patterns:** Analyze common PMO integration issues

## PMO Future Enhancements

### Multi-Project PMO Support
```javascript
// Enhanced project detection and PMO routing
var projectType = detectProjectType(ticketKey); // CXPRODELIVERY, ANOTHER, etc.
var pmoEndpoint = getPMOEndpointForProject(projectType);
var result = getPMOProjectFolder(ticketKey, pmoEndpoint);
```

### Advanced PMO Features
- **Bulk Operations:** PMO batch folder creation for multiple tickets
- **Folder Templates:** PMO-managed folder structure templates
- **Permission Inheritance:** Automatic permission assignment from PMO system
- **Folder Analytics:** Usage reporting and optimization suggestions

### PMO Integration Extensions
- **Real-time Notifications:** PMO folder creation/update alerts
- **Audit Integration:** Enhanced audit trail with PMO system logging
- **Mobile Optimization:** Improved PMO integration for mobile Gmail access
- **Offline Capability:** Queue PMO requests for later processing when offline
