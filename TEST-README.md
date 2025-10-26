# Google Drive Integration Test

This is a standalone test script to verify Google Drive folder access and file saving functionality without PMO webhook dependency.

## Purpose

Tests the Google Drive components of the Gmail Attachment Saver to ensure:
- ✅ Folder access works with the provided folder ID
- ✅ File creation and saving functionality operates correctly  
- ✅ Duplicate file handling works as expected
- ✅ Proper error handling and logging

## Setup

### 1. Create New Apps Script Project
1. Go to [Google Apps Script](https://script.google.com/)
2. Create a **New Project**
3. Replace the default `Code.gs` content with `test-gdrive-save.js`
4. Replace the default `appsscript.json` with `test-appsscript.json`

### 2. Configure OAuth Permissions
The test only requires Google Drive access:
- `https://www.googleapis.com/auth/drive`

### 3. Test Configuration
The test uses hard-coded values:
```javascript
var TEST_CONFIG = {
  folderId: "1ES9qkAU7m5oVkDckJVjrYqaxmxdNi8Xr", // Your provided folder ID
  testFileName: "gmail-attachment-saver-test.txt",
  // ... test content
};
```

## Running Tests

### Quick Test
Run the main test function:
```javascript
testGDriveSave()
```

**Expected Output:**
```
=== GOOGLE DRIVE SAVE TEST ===
Step 1: Testing folder access...
✅ FOLDER ACCESS SUCCESS
- Folder name: [Your Folder Name]

Step 2: Creating test file...
✅ FILE CREATION SUCCESS
- File name: gmail-attachment-saver-test.txt

Step 3: Testing duplicate file handling...
✅ DUPLICATE HANDLING SUCCESS

=== TEST COMPLETED SUCCESSFULLY ===
🎉 Google Drive integration is working correctly!
```

### Comprehensive Test
Run all tests:
```javascript
runFullGDriveTest()
```

### Individual Tests
- `testFolderPermissions()` - Test folder access and permissions
- `testMultipleFiles()` - Test creating multiple files
- `cleanupTestFiles()` - Remove test files

## Test Results Interpretation

### ✅ SUCCESS - All Tests Pass
- Google Drive API access works
- Folder permissions are correct
- File creation functionality works
- The Gmail Attachment Saver will work once PMO connectivity is resolved

### ❌ FAILURE Scenarios

#### "Cannot access folder" Error
- **Cause**: Folder ID incorrect or no access permissions
- **Solution**: 
  - Verify folder ID: `1ES9qkAU7m5oVkDckJVjrYqaxmxdNi8Xr`
  - Check if you have edit permissions to the folder
  - Make sure folder wasn't deleted or moved

#### "Permission denied" Error  
- **Cause**: Insufficient OAuth scopes or Drive API access
- **Solution**:
  - Ensure `https://www.googleapis.com/auth/drive` scope is configured
  - Re-authorize the script when prompted
  - Check if Drive API is enabled

#### "File creation failed" Error
- **Cause**: Folder is read-only or Drive quota exceeded
- **Solution**:
  - Verify you have write permissions to the folder
  - Check Google Drive storage quota
  - Ensure folder isn't restricted

## What This Tests

### Core Gmail Attachment Saver Functions
1. **`getFolderByIdSafely()`** - Safe folder access with error handling
2. **File creation with `DriveApp.createFile()`**
3. **Duplicate file handling** with timestamp suffixes
4. **Blob creation** from text content
5. **Error handling and logging**

### Simulates Real Usage
- Creates files similar to email attachments
- Tests duplicate scenarios (multiple files with same name)
- Uses same folder access patterns as main application
- Includes proper cleanup functionality

## Integration with Main App

Once this test passes, you can be confident that:
- The Google Drive integration components work correctly
- The issue is isolated to PMO webhook connectivity
- File saving will work once PMO DNS/firewall issues are resolved
- No changes needed to the Google Drive code

## Troubleshooting

### Script Authorization
If prompted for authorization:
1. Click "Review permissions"
2. Choose your Google account
3. Click "Advanced" → "Go to [Project Name] (unsafe)"
4. Click "Allow"

### Folder Access Issues
1. Open the folder URL directly: `https://drive.google.com/drive/folders/1ES9qkAU7m5oVkDckJVjrYqaxmxdNi8Xr`
2. Verify you can see and edit the folder
3. Check if folder sharing settings allow your account access

### Console Logging
- View detailed logs in Apps Script Editor → "Executions"
- All test functions provide comprehensive console output
- Error messages include specific guidance for resolution

## Next Steps

✅ **If tests pass**: Google Drive integration is working - focus on resolving PMO connectivity  
❌ **If tests fail**: Fix Google Drive access issues before addressing PMO integration
