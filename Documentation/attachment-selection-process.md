# Attachment Selection Process Flow

## Overview
The Attachment Selection feature provides an intelligent interface for users to choose specific attachments to save, with persistent memory across sessions and smart grouping by file type.

## Process Flow

### 1. Attachment Discovery
```
Email Thread → Message Iteration → Attachment Extraction → Metadata Collection
```

**Technical Steps:**
1. **Thread Analysis**
   ```javascript
   var thread = GmailApp.getThreadById(threadId);
   var messages = thread.getMessages();
   var allAttachments = [];
   ```

2. **Message Processing**
   - Iterate through each message in thread
   - Extract all attachments per message
   - Store message index for reference

3. **Metadata Extraction**
   ```javascript
   allAttachments.push({
     name: attachment.getName(),
     size: attachment.getSize(), 
     attachment: attachment,
     messageIndex: i
   });
   ```

### 2. Attachment Grouping & Organization
```
Raw Attachments → File Extension Analysis → Group by Type → Sort & Count
```

**Grouping Logic:**
```javascript
var attachmentsByExtension = {};
var extensionCounts = {};

for (var i = 0; i < allAttachments.length; i++) {
  var fileName = attachment.name;
  var extension = fileName.lastIndexOf('.') > -1 ? 
    fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
  
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
```

**Display Categories:**
- **PDF Files** → `.pdf` extensions
- **Image Files** → `.jpg`, `.png`, `.gif`, etc.  
- **Document Files** → `.docx`, `.xlsx`, `.pptx`, etc.
- **Archive Files** → `.zip`, `.rar`, `.7z`, etc.
- **Files without extension** → Special category for extensionless files

### 3. UI Section Generation
```
Grouped Attachments → Create Sections → Add Checkboxes → Apply Styling → Display
```

**Section Creation Process:**
1. **Sort Extensions**
   ```javascript
   var sortedExtensions = Object.keys(attachmentsByExtension).sort();
   ```

2. **Create Collapsible Sections**
   ```javascript
   var extensionSection = CardService.newCardSection()
     .setHeader(headerText)
     .setCollapsible(true)
     .setNumUncollapsibleWidgets(count > 5 ? 2 : count);
   ```

3. **Generate Checkboxes**
   ```javascript
   var checkbox = CardService.newSelectionInput()
     .setType(CardService.SelectionInputType.CHECK_BOX)
     .setFieldName("attachment_" + originalIndex)
     .addItem(displayText, originalIndex.toString(), isSelected);
   ```

### 4. Selection Memory System
```
User Selections → Store State → Retrieve on Reload → Apply Previous Selections
```

**Memory Storage Process:**

#### 4.1 Selection State Capture
```javascript
// During form submission or UI updates
var selectionState = {};
for (var i = 0; i < allAttachments.length; i++) {
  var checkboxValue = e.formInput["attachment_" + i];
  var isSelected = checkboxValue && checkboxValue.length > 0;
  var uniqueKey = allAttachments[i].name + "_" + i;
  selectionState[uniqueKey] = isSelected;
}
```

#### 4.2 Persistent Storage
```javascript
function storeAttachmentSelections(threadId, selectionState) {
  var userProperties = PropertiesService.getUserProperties();
  var key = 'ATTACHMENT_SELECTIONS_' + threadId;
  userProperties.setProperty(key, JSON.stringify(selectionState));
}
```

#### 4.3 Selection Retrieval
```javascript
function getStoredAttachmentSelections(threadId) {
  var userProperties = PropertiesService.getUserProperties();
  var key = 'ATTACHMENT_SELECTIONS_' + threadId;
  var storedSelectionsJson = userProperties.getProperty(key);
  
  if (storedSelectionsJson) {
    return JSON.parse(storedSelectionsJson);
  }
  return null;
}
```

### 5. Selection State Management
```
Form Changes → Capture State → Update Memory → Refresh UI → Maintain Context
```

**State Management Features:**

#### 5.1 Unique Key System
- **Purpose:** Handle duplicate filenames across messages
- **Format:** `filename + "_" + attachmentIndex`
- **Example:** `"document.pdf_0"`, `"document.pdf_1"` for duplicate names

#### 5.2 Default Behavior
- **Initial State:** All attachments unselected by default
- **User Choice:** Explicit selection required
- **Memory Override:** Stored selections take precedence

#### 5.3 Dynamic Updates
```javascript
// On ticket selection change
if (e.formInput) {
  // Preserve current checkbox states
  var currentSelections = {};
  for (var i = 0; i < allAttachments.length; i++) {
    var formFieldName = "attachment_" + i;
    var isCurrentlySelected = e.formInput.hasOwnProperty(formFieldName);
    var uniqueKey = allAttachments[i].name + "_" + i;
    currentSelections[uniqueKey] = isCurrentlySelected;
  }
  
  // Store for future use
  storeAttachmentSelections(threadId, currentSelections);
}
```

## Display Features

### 1. File Information Display
```
Filename + File Size + Selection Status
```

**Format:** 
```
☐ document.pdf (2.4 MB)
☑ spreadsheet.xlsx (1.1 MB)  
☐ presentation.pptx (5.7 MB)
```

**Size Formatting:**
```javascript
function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';
  var k = 1024;
  var sizes = ['B', 'KB', 'MB', 'GB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}
```

### 2. Section Headers
```
📎 Select Attachments: PDF files (3) (Total: 8)
📄 DOCX files (2)
🖼️ Image files (3)
```

**Header Logic:**
- First section shows total attachment count
- Each section shows file type and count
- Collapsible for better organization

### 3. Progressive Disclosure
- **Large Groups:** Show first 2 files, collapse rest
- **Small Groups:** Show all files expanded
- **User Control:** Manual expand/collapse available

## Error Handling

### Memory Storage Failures
```javascript
try {
  userProperties.setProperty(key, JSON.stringify(selectionState));
} catch (error) {
  console.error("Failed to store selections:", error);
  // Continue without memory - not critical failure
}
```

### Attachment Processing Errors
- **Large Files:** Handle memory limits gracefully
- **Corrupted Attachments:** Skip problematic files
- **Permission Issues:** Log and continue with accessible files

### UI State Corruption
- **Recovery:** Reset to default unselected state
- **Validation:** Verify form field names match attachments
- **Fallback:** Manual selection without memory

## Performance Optimizations

### Memory Management
1. **Cleanup Strategy:** Remove old thread selection data
2. **Size Limits:** Limit stored selection history
3. **Efficient Keys:** Minimize storage key length

### UI Rendering
1. **Lazy Loading:** Create UI sections on demand
2. **Batch Processing:** Group checkbox creation
3. **Memory Reuse:** Cache DOM elements where possible

### Selection Processing
1. **Debounced Updates:** Avoid excessive storage writes
2. **Batch Operations:** Process multiple selections together
3. **Smart Refresh:** Only update changed selections

## User Experience Features

### Visual Feedback
- **Selection Indicators:** Clear checkmark/empty checkbox states
- **File Type Icons:** Visual grouping cues
- **Size Information:** Help users make informed choices

### Accessibility
- **Keyboard Navigation:** Tab through checkboxes
- **Screen Reader Support:** Proper ARIA labels
- **Clear Labels:** Descriptive checkbox text

### Smart Defaults
- **Context Awareness:** Remember user preferences
- **Bulk Selection:** Future enhancement for select all/none
- **Filter Options:** Future enhancement for file type filtering

## Integration Points

### Gmail Context
- **Thread ID:** Links selections to specific email
- **Message Index:** Tracks attachment source message
- **Persistence:** Survives Gmail navigation

### Save Process
- **Validation:** Ensure at least one attachment selected
- **Progress:** Show selection count during save
- **Confirmation:** Report successful saves

### Settings Integration
- **Memory Duration:** Configurable selection memory retention
- **Default Behavior:** User-configurable selection defaults
- **Export Options:** Future bulk selection management
