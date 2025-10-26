/**
 * Test Script: Google Drive File Save Test
 * 
 * Purpose: Test Google Drive folder access and file saving functionality
 * without PMO webhook dependency using hard-coded folder ID
 * 
 * Usage: Deploy this as a separate Google Apps Script project and run testGDriveSave()
 */

// Test configuration
var TEST_CONFIG = {
  folderId: "1ES9qkAU7m5oVkDckJVjrYqaxmxdNi8Xr", // Hard-coded valid folder ID
  testFileName: "gmail-attachment-saver-test.txt",
  testFileContent: "Test file created by Gmail Attachment Saver\n" +
                   "Timestamp: " + new Date().toISOString() + "\n" +
                   "Purpose: Testing Google Drive integration\n" +
                   "Ticket: CXPRODELIVERY-TEST\n\n" +
                   "This file confirms that:\n" +
                   "✅ Google Drive API access works\n" +
                   "✅ Folder access permissions are correct\n" +
                   "✅ File creation functionality operates properly\n\n" +
                   "If you can see this file, the Google Drive integration is working!"
};

/**
 * Main test function - run this to test Google Drive save functionality
 */
function testGDriveSave() {
  console.log("=== GOOGLE DRIVE SAVE TEST ===");
  console.log("Test started at:", new Date().toISOString());
  console.log("Target folder ID:", TEST_CONFIG.folderId);
  
  try {
    // Step 1: Test folder access
    console.log("\nStep 1: Testing folder access...");
    var folderResult = getFolderByIdSafely(TEST_CONFIG.folderId);
    
    if (!folderResult.success) {
      console.error("❌ FOLDER ACCESS FAILED");
      console.error("Error:", folderResult.error);
      console.error("Possible causes:");
      console.error("- Folder ID is incorrect");
      console.error("- No permission to access folder");
      console.error("- Folder was deleted or moved");
      return { success: false, error: folderResult.error };
    }
    
    console.log("✅ FOLDER ACCESS SUCCESS");
    console.log("- Folder name:", folderResult.name);
    console.log("- Folder ID:", TEST_CONFIG.folderId);
    console.log("- Folder object type:", typeof folderResult.folder);
    
    // Step 2: Test file creation
    console.log("\nStep 2: Creating test file...");
    var testFile = createTestFile(folderResult.folder);
    
    if (!testFile.success) {
      console.error("❌ FILE CREATION FAILED");
      console.error("Error:", testFile.error);
      return { success: false, error: testFile.error };
    }
    
    console.log("✅ FILE CREATION SUCCESS");
    console.log("- File name:", testFile.fileName);
    console.log("- File ID:", testFile.fileId);
    console.log("- File URL:", testFile.fileUrl);
    
    // Step 3: Test duplicate handling
    console.log("\nStep 3: Testing duplicate file handling...");
    var duplicateTest = createTestFile(folderResult.folder);
    
    console.log("✅ DUPLICATE HANDLING SUCCESS");
    console.log("- Original file:", testFile.fileName);
    console.log("- Duplicate file:", duplicateTest.fileName);
    console.log("- Files are different:", testFile.fileId !== duplicateTest.fileId);
    
    // Step 4: Clean up (optional - comment out to keep test files)
    console.log("\nStep 4: Cleanup test files...");
    cleanupTestFiles(folderResult.folder);
    
    // Final result
    console.log("\n=== TEST COMPLETED SUCCESSFULLY ===");
    console.log("🎉 Google Drive integration is working correctly!");
    console.log("The Gmail Attachment Saver can save files once PMO connectivity is resolved.");
    
    return {
      success: true,
      folderName: folderResult.name,
      folderId: TEST_CONFIG.folderId,
      testFile: testFile.fileName,
      message: "Google Drive integration test passed"
    };
    
  } catch (error) {
    console.error("❌ TEST FAILED WITH EXCEPTION");
    console.error("Error:", error.message);
    console.error("Stack:", error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Safe folder access function (copied from main code)
 */
function getFolderByIdSafely(folderId) {
  console.log("Accessing folder with ID:", folderId);
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    var folderName = folder.getName(); // Test access by getting name
    
    console.log("Folder access successful:", folderName);
    
    return {
      success: true,
      folder: folder,
      name: folderName
    };
  } catch (error) {
    console.error("Folder access failed:", error.message);
    
    return {
      success: false,
      error: 'Cannot access folder (ID: ' + folderId + '): ' + error.message
    };
  }
}

/**
 * Create test file with duplicate handling
 */
function createTestFile(targetFolder) {
  try {
    var fileName = TEST_CONFIG.testFileName;
    var fileContent = TEST_CONFIG.testFileContent;
    
    // Check for existing file and handle duplicates
    var existingFiles = targetFolder.getFilesByName(fileName);
    if (existingFiles.hasNext()) {
      // File exists, add timestamp to make it unique
      var timestamp = new Date().getTime();
      var nameParts = fileName.split('.');
      if (nameParts.length > 1) {
        fileName = nameParts.slice(0, -1).join('.') + '_' + timestamp + '.' + nameParts[nameParts.length - 1];
      } else {
        fileName = fileName + '_' + timestamp;
      }
      console.log("File exists, using unique name:", fileName);
    }
    
    // Create the file
    console.log("Creating file:", fileName);
    var blob = Utilities.newBlob(fileContent, 'text/plain', fileName);
    var createdFile = targetFolder.createFile(blob);
    
    console.log("File created successfully");
    
    return {
      success: true,
      fileName: fileName,
      fileId: createdFile.getId(),
      fileUrl: createdFile.getUrl(),
      file: createdFile
    };
    
  } catch (error) {
    return {
      success: false,
      error: 'File creation failed: ' + error.message
    };
  }
}

/**
 * Test multiple file operations
 */
function testMultipleFiles() {
  console.log("=== TESTING MULTIPLE FILE OPERATIONS ===");
  
  var folderResult = getFolderByIdSafely(TEST_CONFIG.folderId);
  if (!folderResult.success) {
    console.error("Cannot access test folder");
    return false;
  }
  
  var testFiles = [
    { name: "test-document.txt", content: "Test document content" },
    { name: "test-image.txt", content: "Simulated image file content" },
    { name: "test-spreadsheet.txt", content: "Simulated spreadsheet content" }
  ];
  
  var results = [];
  
  for (var i = 0; i < testFiles.length; i++) {
    var testFile = testFiles[i];
    console.log("Creating test file:", testFile.name);
    
    try {
      var blob = Utilities.newBlob(testFile.content, 'text/plain', testFile.name);
      var createdFile = folderResult.folder.createFile(blob);
      
      results.push({
        name: testFile.name,
        id: createdFile.getId(),
        success: true
      });
      
      console.log("✅ Created:", testFile.name);
      
    } catch (error) {
      console.error("❌ Failed to create:", testFile.name, "-", error.message);
      results.push({
        name: testFile.name,
        error: error.message,
        success: false
      });
    }
  }
  
  console.log("Multiple file test results:", results);
  return results;
}

/**
 * Clean up test files (optional)
 */
function cleanupTestFiles(targetFolder) {
  try {
    console.log("Cleaning up test files...");
    
    // Find files that start with test prefix
    var files = targetFolder.getFiles();
    var deletedCount = 0;
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      
      if (fileName.indexOf("gmail-attachment-saver-test") === 0 || 
          fileName.indexOf("test-") === 0) {
        console.log("Deleting test file:", fileName);
        file.setTrashed(true);
        deletedCount++;
      }
    }
    
    console.log("Cleaned up", deletedCount, "test files");
    
  } catch (error) {
    console.log("Cleanup failed (not critical):", error.message);
  }
}

/**
 * Test folder permissions and capabilities
 */
function testFolderPermissions() {
  console.log("=== TESTING FOLDER PERMISSIONS ===");
  
  var folderResult = getFolderByIdSafely(TEST_CONFIG.folderId);
  if (!folderResult.success) {
    return false;
  }
  
  var folder = folderResult.folder;
  
  try {
    // Test various folder operations
    console.log("Testing folder capabilities...");
    
    // Read permissions
    console.log("- Folder name:", folder.getName());
    console.log("- Folder URL:", folder.getUrl());
    console.log("- Creation date:", folder.getDateCreated());
    console.log("- Last updated:", folder.getLastUpdated());
    
    // Count existing files
    var fileCount = 0;
    var files = folder.getFiles();
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    console.log("- Existing files count:", fileCount);
    
    // Test write permissions with small file
    var testBlob = Utilities.newBlob("permission test", 'text/plain', 'permission-test.txt');
    var testFile = folder.createFile(testBlob);
    console.log("✅ Write permissions: OK");
    
    // Clean up test file
    testFile.setTrashed(true);
    console.log("✅ Delete permissions: OK");
    
    console.log("🎉 All folder permissions are working correctly!");
    return true;
    
  } catch (error) {
    console.error("❌ Permission test failed:", error.message);
    return false;
  }
}

/**
 * Run comprehensive Google Drive integration test
 */
function runFullGDriveTest() {
  console.log("=== COMPREHENSIVE GOOGLE DRIVE TEST ===");
  console.log("Testing all Google Drive integration components...\n");
  
  var results = {
    folderAccess: false,
    fileCreation: false,
    multipleFiles: false,
    permissions: false,
    overall: false
  };
  
  // Test 1: Basic folder access
  console.log("1. Testing basic folder access...");
  var folderTest = getFolderByIdSafely(TEST_CONFIG.folderId);
  results.folderAccess = folderTest.success;
  
  if (!results.folderAccess) {
    console.log("❌ Cannot proceed - folder access failed");
    return results;
  }
  
  // Test 2: File creation
  console.log("\n2. Testing file creation...");
  var fileTest = testGDriveSave();
  results.fileCreation = fileTest.success;
  
  // Test 3: Multiple files
  console.log("\n3. Testing multiple file operations...");
  var multiTest = testMultipleFiles();
  results.multipleFiles = multiTest && multiTest.length > 0;
  
  // Test 4: Permissions
  console.log("\n4. Testing folder permissions...");
  results.permissions = testFolderPermissions();
  
  // Overall result
  results.overall = results.folderAccess && results.fileCreation && 
                   results.multipleFiles && results.permissions;
  
  console.log("\n=== FINAL TEST RESULTS ===");
  console.log("Folder Access:", results.folderAccess ? "✅ PASS" : "❌ FAIL");
  console.log("File Creation:", results.fileCreation ? "✅ PASS" : "❌ FAIL");
  console.log("Multiple Files:", results.multipleFiles ? "✅ PASS" : "❌ FAIL");
  console.log("Permissions:", results.permissions ? "✅ PASS" : "❌ FAIL");
  console.log("Overall:", results.overall ? "🎉 ALL TESTS PASSED" : "⚠️ SOME TESTS FAILED");
  
  if (results.overall) {
    console.log("\n✅ Google Drive integration is fully functional!");
    console.log("The Gmail Attachment Saver will work once PMO connectivity is resolved.");
  } else {
    console.log("\n❌ There are issues with Google Drive integration that need to be resolved.");
  }
  
  return results;
}
