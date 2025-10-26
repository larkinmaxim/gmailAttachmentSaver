# Duplicate Handling Process Flow

## Overview
The Duplicate Handling feature intelligently manages file duplicates by comparing file sizes, skipping identical files, and creating timestamped versions for different files with the same name.

## Process Flow

### 1. Duplicate Detection Strategy
```
File Save Request → Name Check → Size Comparison → Decision Logic → Action Execution
```

**Detection Algorithm:**
```javascript
// Check if file already exists in target folder
var existingFiles = ticketFolder.getFilesByName(fileName);

if (existingFiles.hasNext()) {
  var existingFile = existingFiles.next();
  var existingSize = existingFile.getSize();
  var newSize = attachmentData.size;
  
  if (existingSize === newSize) {
    // Likely duplicate - skip
    skipFile(fileName);
  } else {
    // Different content - create timestamped version
    createVersionedFile(attachmentData);
  }
} else {
  // New file - save normally
  saveNewFile(attachmentData);
}
```

### 2. Size-Based Duplicate Detection
```
Filename Match → Size Comparison → Duplicate Classification → Action Decision
```

**Classification Logic:**

#### 2.1 Identical Files (Same Size)
- **Assumption:** Files with identical names and sizes are duplicates
- **Action:** Skip saving, increment skip counter
- **Reason:** Avoid storage waste and confusion
- **User Feedback:** Report as "duplicate skipped"

#### 2.2 Different Content (Different Size)  
- **Assumption:** Same name but different size indicates updated/different content
- **Action:** Save with timestamp suffix
- **Reason:** Preserve both versions for user review
- **User Feedback:** Report as "versioned save"

#### 2.3 New Files (No Name Match)
- **Action:** Save with original filename
- **Reason:** No conflict exists
- **User Feedback:** Report as "successfully saved"

### 3. Timestamped Version Creation
```
File Conflict → Generate Timestamp → Parse Filename → Create New Name → Save File
```

**Timestamp Generation:**
```javascript
var timestamp = Utilities.formatDate(
  new Date(), 
  Session.getScriptTimeZone(), 
  "yyyyMMdd_HHmmss"
);
```

**Filename Processing:**
```javascript
var fileExtension = fileName.lastIndexOf('.') > -1 ? 
  fileName.substring(fileName.lastIndexOf('.')) : '';
var baseName = fileName.lastIndexOf('.') > -1 ? 
  fileName.substring(0, fileName.lastIndexOf('.')) : fileName;

var newFileName = baseName + "_" + timestamp + fileExtension;
```

**Examples:**
```
Original: "contract.pdf" → Versioned: "contract_20241026_143022.pdf"
Original: "image.jpg" → Versioned: "image_20241026_143023.jpg"  
Original: "document" → Versioned: "document_20241026_143024"
```

### 4. Processing Workflow
```
Selected Attachments → Process Each File → Apply Duplicate Logic → Track Results → Report Summary
```

**Batch Processing:**
```javascript
var savedCount = 0;
var skippedCount = 0;
var savedFiles = [];
var skippedFiles = [];

for (var i = 0; i < selectedAttachments.length; i++) {
  try {
    var result = processSingleAttachment(selectedAttachments[i], ticketFolder);
    
    if (result.action === 'saved') {
      savedCount++;
      savedFiles.push(result.filename);
    } else if (result.action === 'skipped') {
      skippedCount++;
      skippedFiles.push(result.filename);
    }
  } catch (error) {
    console.error("Processing error:", error);
    // Continue with next file
  }
}
```

### 5. User Notification System
```
Processing Results → Generate Summary → Create Notification → Display to User
```

**Notification Format:**
```javascript
var notificationText = "✅ Saved " + savedCount + " attachments to CXPRODELIVERY/" + finalTicket;
if (skippedCount > 0) {
  notificationText += " (⚠️ " + skippedCount + " duplicates skipped)";
}
```

**Example Messages:**
- Success: `"✅ Saved 3 attachments to CXPRODELIVERY/CXPRODELIVERY-6310"`
- With Duplicates: `"✅ Saved 2 attachments to CXPRODELIVERY/CXPRODELIVERY-6310 (⚠️ 1 duplicates skipped)"`
- Versions Created: `"✅ Saved 3 attachments (1 versioned) to CXPRODELIVERY/CXPRODELIVERY-6310"`

## Advanced Duplicate Scenarios

### 1. Multiple Version Handling
```
Existing Versions → Find Latest → Compare All Sizes → Create Next Version
```

**Version Chain Example:**
```
contract.pdf                  (Original)
contract_20241020_100000.pdf (Version 1) 
contract_20241025_140000.pdf (Version 2)
contract_20241026_143022.pdf (New Version)
```

### 2. Cross-Message Duplicates
- **Scenario:** Same file attached to multiple messages in thread
- **Detection:** Compare across all thread messages
- **Resolution:** Skip subsequent identical attachments
- **Tracking:** Note message source in logs

### 3. Partial Name Matches
```
Similar Names → Exact Match Check → Proceed with Normal Logic
```

**Examples of Non-Matches:**
- `contract.pdf` vs `contract_v2.pdf` → Different files
- `image.jpg` vs `image.jpeg` → Different extensions  
- `doc.docx` vs `doc.pdf` → Different formats

## Error Handling & Edge Cases

### File System Limitations
```javascript
try {
  var file = ticketFolder.createFile(attachmentData.attachment);
  file.setName(newFileName);
} catch (error) {
  if (error.message.includes('name too long')) {
    // Truncate filename and retry
    var truncatedName = truncateFilename(newFileName);
    file.setName(truncatedName);
  } else {
    throw error; // Re-throw other errors
  }
}
```

### Drive API Limits
- **Rate Limiting:** Handle API quota exceeded
- **Storage Limits:** Detect and report storage full errors
- **Permission Issues:** Handle folder access denied

### Filename Character Issues
```javascript
function sanitizeFilename(filename) {
  // Remove invalid characters for Google Drive
  return filename.replace(/[<>:"/\\|?*]/g, '_');
}
```

## Performance Considerations

### Optimization Strategies

#### 1. Batch File Operations
- **Single Folder Lookup:** Reuse folder reference across files
- **Cached Existence Checks:** Remember which files exist
- **Parallel Processing:** Process non-conflicting files simultaneously

#### 2. Efficient Size Checking
```javascript
// Cache file sizes to avoid repeated API calls
var existingFileSizes = {};
var files = ticketFolder.getFiles();
while (files.hasNext()) {
  var file = files.next();
  existingFileSizes[file.getName()] = file.getSize();
}
```

#### 3. Smart Conflict Resolution
- **Early Detection:** Check for conflicts before processing attachment data
- **Selective Processing:** Only process files that will actually be saved
- **Memory Efficiency:** Process large files individually

### Resource Management
- **Memory Usage:** Stream large files instead of loading into memory
- **API Quotas:** Respect Google Drive API rate limits
- **Execution Time:** Handle Google Apps Script timeout limits

## Logging & Debugging

### Comprehensive Logging
```javascript
console.log("=== DUPLICATE HANDLING RESULTS ===");
console.log("Total files processed:", selectedAttachments.length);
console.log("Files saved:", savedCount);
console.log("Files skipped (duplicates):", skippedCount);
console.log("Saved files:", savedFiles);
console.log("Skipped files:", skippedFiles);
```

### Debug Information
- **File Size Comparisons:** Log existing vs new file sizes
- **Timestamp Generation:** Verify timestamp format and uniqueness
- **Filename Transformations:** Track original → versioned name changes

### Error Tracking
```javascript
var processingErrors = [];
// ... during processing ...
if (error) {
  processingErrors.push({
    filename: attachmentData.name,
    error: error.message,
    timestamp: new Date()
  });
}
```

## User Experience Benefits

### 1. Intelligent Behavior
- **No User Intervention:** Automatic duplicate detection
- **Preserved Content:** Never lose different versions
- **Clear Communication:** Transparent reporting of actions

### 2. Storage Efficiency
- **Duplicate Prevention:** Avoid unnecessary storage usage  
- **Version Control:** Maintain file history when needed
- **Clean Organization:** Predictable filename patterns

### 3. Conflict Resolution
- **Non-Destructive:** Original files never overwritten
- **Traceable Changes:** Timestamp shows when versions created
- **User Control:** Can manually clean up versions later

## Future Enhancements

### Advanced Duplicate Detection
- **Content Hashing:** MD5/SHA256 comparison for true duplicate detection
- **Fuzzy Matching:** Detect similar filenames with minor differences
- **Size Threshold:** Ignore tiny size differences (e.g., metadata changes)

### User Configuration Options
- **Duplicate Policy:** User choice of skip/version/ask behavior
- **Timestamp Format:** Customizable version naming patterns
- **Automatic Cleanup:** Option to remove old versions automatically

### Integration Features
- **Version History:** Link to Google Drive version history
- **Merge Detection:** Identify when different versions can be consolidated
- **Conflict Reporting:** Detailed duplicate analysis reports
