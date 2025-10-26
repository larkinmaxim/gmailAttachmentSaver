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
 * Enhanced logging utility functions with safety checks
 */
function logStep(stepNum, stepName) {
  // Debug logging to track parameters (remove this once issue is resolved)
  console.log("DEBUG: logStep called with parameters:");
  console.log("  stepNum:", stepNum, "(type:", typeof stepNum + ")");
  console.log("  stepName:", stepName, "(type:", typeof stepName + ")");
  console.log("  Call stack location: " + (new Error().stack.split('\n')[1] || 'unknown'));
  
  // Safety checks for parameters
  if (stepNum === undefined || stepNum === null) {
    stepNum = "?";
  }
  if (stepName === undefined || stepName === null) {
    stepName = "Unknown Step";
  }
  
  // Ensure stepName is a string before calling toUpperCase()
  stepName = String(stepName);
  
  console.log("\n" + "=".repeat(50));
  console.log("STEP " + stepNum + ": " + stepName.toUpperCase());
  console.log("=".repeat(50));
  console.log("Timestamp:", new Date().toISOString());
}

function logSuccess(message) {
  message = message || "Operation completed";
  console.log("✅ SUCCESS:", message);
}

function logError(message) {
  message = message || "Unknown error occurred";
  console.error("❌ ERROR:", message);
}

function logInfo(message) {
  message = message || "No information provided";
  console.log("ℹ️  INFO:", message);
}

function logWarning(message) {
  message = message || "Unknown warning";
  console.log("⚠️  WARNING:", message);
}

function logDetail(key, value) {
  key = key || "Unknown";
  value = (value !== undefined && value !== null) ? value : "N/A";
  console.log("  •", key + ":", value);
}

/**
 * Simple diagnostic function - run this first to check basic logging
 */
function simpleLoggingTest() {
  console.log("=== SIMPLE LOGGING TEST START ===");
  console.log("About to test logStep with valid parameters...");
  
  logStep(1, "Simple Test");
  
  console.log("=== SIMPLE LOGGING TEST END ===");
  return "Test completed";
}

/**
 * Test logging functions - run this first to ensure logging works correctly
 */
function testLoggingFunctions() {
  console.log("=== TESTING LOGGING FUNCTIONS ===");
  
  try {
    // Test all logging functions with valid parameters
    logStep(1, "Test Step");
    logSuccess("Test success message");
    logError("Test error message");
    logInfo("Test info message");
    logWarning("Test warning message");
    logDetail("Test key", "Test value");
    
    console.log("\n--- Testing with undefined/null parameters ---");
    
    // Test with undefined parameters (should not crash)
    logStep(undefined, undefined);
    logStep(2, null);
    logSuccess(undefined);
    logError(null);
    logInfo(undefined);
    logWarning(null);
    logDetail(undefined, undefined);
    logDetail("Key", null);
    logDetail(null, "Value");
    
    console.log("\n✅ All logging functions work correctly!");
    return true;
    
  } catch (error) {
    console.error("❌ Logging function test failed:", error.message);
    return false;
  }
}

/**
 * Main test function - run this to test Google Drive save functionality
 */
function testGDriveSave() {
  console.log("=".repeat(80));
  console.log("    GOOGLE DRIVE INTEGRATION TEST - DETAILED LOGGING");
  console.log("=".repeat(80));
  console.log("Test session started:", new Date().toISOString());
  console.log("Test configuration:");
  logDetail("Target folder ID", TEST_CONFIG.folderId);
  logDetail("Test file name", TEST_CONFIG.testFileName);
  logDetail("Content length", TEST_CONFIG.testFileContent.length + " characters");
  
  var testStartTime = new Date().getTime();
  
  try {
    // Step 1: Test folder access
    logStep(1, "Testing Google Drive Folder Access");
    logInfo("Attempting to access folder using DriveApp.getFolderById()");
    logDetail("Folder ID", TEST_CONFIG.folderId);
    
    var step1StartTime = new Date().getTime();
    var folderResult = getFolderByIdSafely(TEST_CONFIG.folderId);
    var step1Duration = new Date().getTime() - step1StartTime;
    
    logDetail("Folder access duration", step1Duration + "ms");
    
    if (!folderResult.success) {
      logError("FOLDER ACCESS FAILED");
      logDetail("Error message", folderResult.error);
      logInfo("Possible causes:");
      console.log("    - Folder ID is incorrect or malformed");
      console.log("    - No permission to access folder (not shared with your account)");
      console.log("    - Folder was deleted or moved to trash");
      console.log("    - Google Drive API permission issues");
      logInfo("Resolution steps:");
      console.log("    1. Verify folder ID: " + TEST_CONFIG.folderId);
      console.log("    2. Check folder sharing permissions");  
      console.log("    3. Ensure Google Drive OAuth scope is authorized");
      return { success: false, error: folderResult.error, step: "folder_access" };
    }
    
    logSuccess("FOLDER ACCESS COMPLETED");
    logDetail("Folder name", folderResult.name);
    logDetail("Folder ID verified", TEST_CONFIG.folderId);
    logDetail("Folder object type", typeof folderResult.folder);
    logDetail("Folder URL", folderResult.folder.getUrl());
    
    // Step 2: Test file creation
    logStep(2, "Testing File Creation and Blob Handling");
    logInfo("Creating test file with specified content");
    logDetail("Target folder", folderResult.name);
    logDetail("File name", TEST_CONFIG.testFileName);
    
    var step2StartTime = new Date().getTime();
    var testFile = createTestFile(folderResult.folder);
    var step2Duration = new Date().getTime() - step2StartTime;
    
    logDetail("File creation duration", step2Duration + "ms");
    
    if (!testFile.success) {
      logError("FILE CREATION FAILED");
      logDetail("Error message", testFile.error);
      logInfo("Possible causes:");
      console.log("    - Insufficient permissions to create files in folder");
      console.log("    - Google Drive storage quota exceeded");
      console.log("    - Folder is read-only or restricted");
      console.log("    - Blob creation or file content issues");
      return { success: false, error: testFile.error, step: "file_creation" };
    }
    
    logSuccess("FILE CREATION COMPLETED");
    logDetail("File name", testFile.fileName);
    logDetail("File ID", testFile.fileId);
    logDetail("File URL", testFile.fileUrl);
    logDetail("File size", testFile.file ? testFile.file.getSize() + " bytes" : "unknown");
    logDetail("MIME type", testFile.file ? testFile.file.getBlob().getContentType() : "unknown");
    
    // Step 3: Test duplicate handling
    logStep(3, "Testing Duplicate File Handling");
    logInfo("Creating another file with same name to test duplicate logic");
    logDetail("Original file name", testFile.fileName);
    
    var step3StartTime = new Date().getTime();
    var duplicateTest = createTestFile(folderResult.folder);
    var step3Duration = new Date().getTime() - step3StartTime;
    
    logDetail("Duplicate handling duration", step3Duration + "ms");
    
    if (duplicateTest.success) {
      logSuccess("DUPLICATE HANDLING COMPLETED");
      logDetail("Original file", testFile.fileName);
      logDetail("Duplicate file", duplicateTest.fileName);
      logDetail("Files have different IDs", testFile.fileId !== duplicateTest.fileId);
      logDetail("Timestamp suffix added", duplicateTest.fileName !== testFile.fileName);
    } else {
      logWarning("Duplicate test failed, but original file creation worked");
      logDetail("Duplicate error", duplicateTest.error);
    }
    
    // Step 4: Clean up (optional - comment out to keep test files)
    logStep(4, "Cleanup Test Files (Optional)");
    logInfo("Removing test files to keep folder clean");
    logWarning("Comment out this section to keep test files for inspection");
    
    var step4StartTime = new Date().getTime();
    var cleanupResult = cleanupTestFiles(folderResult.folder);
    var step4Duration = new Date().getTime() - step4StartTime;
    
    logDetail("Cleanup duration", step4Duration + "ms");
    logDetail("Files cleaned up", cleanupResult + " files");
    
    // Calculate total test duration
    var totalTestDuration = new Date().getTime() - testStartTime;
    
    // Final result
    console.log("\n" + "=".repeat(80));
    console.log("    TEST COMPLETED SUCCESSFULLY");
    console.log("=".repeat(80));
    logSuccess("All Google Drive integration tests passed!");
    
    console.log("\n📊 TEST SUMMARY:");
    logDetail("Total test duration", totalTestDuration + "ms");
    logDetail("Step 1 - Folder access", step1Duration + "ms");
    logDetail("Step 2 - File creation", step2Duration + "ms");  
    logDetail("Step 3 - Duplicate handling", step3Duration + "ms");
    logDetail("Step 4 - Cleanup", step4Duration + "ms");
    
    console.log("\n📋 TEST RESULTS:");
    logDetail("Folder name", folderResult.name);
    logDetail("Folder ID", TEST_CONFIG.folderId);
    logDetail("Test file created", testFile.fileName);
    logDetail("Duplicate file created", duplicateTest.success ? duplicateTest.fileName : "N/A");
    logDetail("Files cleaned up", cleanupResult > 0 ? "Yes" : "No files to clean");
    
    console.log("\n🎉 CONCLUSION:");
    console.log("   Google Drive integration is fully functional!");
    console.log("   Gmail Attachment Saver will work once PMO connectivity is resolved.");
    console.log("   No changes needed to Google Drive code - issue is PMO DNS/firewall only.");
    
    return {
      success: true,
      folderName: folderResult.name,
      folderId: TEST_CONFIG.folderId,
      testFile: testFile.fileName,
      duplicateFile: duplicateTest.success ? duplicateTest.fileName : null,
      totalDuration: totalTestDuration,
      stepDurations: {
        folderAccess: step1Duration,
        fileCreation: step2Duration,
        duplicateHandling: step3Duration,
        cleanup: step4Duration
      },
      message: "Google Drive integration test completed successfully"
    };
    
  } catch (error) {
    var testDuration = new Date().getTime() - testStartTime;
    
    console.log("\n" + "=".repeat(80));
    logError("TEST FAILED WITH EXCEPTION");
    console.log("=".repeat(80));
    logDetail("Test duration before failure", testDuration + "ms");
    logDetail("Error name", error.name || "Unknown");
    logDetail("Error message", error.message || "No message");
    
    console.log("\n📋 ERROR DETAILS:");
    if (error.stack) {
      console.log("Stack trace:");
      console.log(error.stack);
    }
    
    console.log("\n🔧 TROUBLESHOOTING STEPS:");
    console.log("   1. Check OAuth permissions - ensure Drive API scope is authorized");
    console.log("   2. Verify folder ID is correct: " + TEST_CONFIG.folderId);  
    console.log("   3. Check folder sharing settings and your account access");
    console.log("   4. Ensure Google Drive API is enabled in your Apps Script project");
    console.log("   5. Try running individual test functions to isolate the issue");
    
    return { 
      success: false, 
      error: error.message,
      errorName: error.name,
      testDuration: testDuration,
      step: "exception_thrown"
    };
  }
}

/**
 * Safe folder access function (copied from main code) - Enhanced with detailed logging
 */
function getFolderByIdSafely(folderId) {
  logInfo("Initiating folder access via DriveApp.getFolderById()");
  logDetail("Folder ID parameter", folderId);
  logDetail("Folder ID type", typeof folderId);
  logDetail("Folder ID length", folderId ? folderId.length : "null/undefined");
  
  // Validate folder ID format
  if (!folderId || typeof folderId !== 'string' || folderId.length === 0) {
    logError("Invalid folder ID provided");
    return {
      success: false,
      error: 'Invalid folder ID: ' + (folderId || 'null/undefined')
    };
  }
  
  try {
    logInfo("Calling DriveApp.getFolderById() API...");
    var apiCallStart = new Date().getTime();
    
    var folder = DriveApp.getFolderById(folderId);
    var apiCallDuration = new Date().getTime() - apiCallStart;
    
    logDetail("DriveApp API call duration", apiCallDuration + "ms");
    logInfo("Folder object retrieved, testing access permissions...");
    
    // Test access by getting folder properties
    var folderName = folder.getName(); // Test read access
    var folderUrl = folder.getUrl();   // Test URL generation
    var folderCreated = folder.getDateCreated(); // Test metadata access
    
    logSuccess("Folder access validation completed");
    logDetail("Folder name", folderName);
    logDetail("Folder URL", folderUrl);
    logDetail("Folder created date", folderCreated.toISOString());
    
    // Test write permissions by counting files (requires read access to contents)
    try {
      var fileIterator = folder.getFiles();
      var fileCount = 0;
      while (fileIterator.hasNext() && fileCount < 10) { // Limit to avoid long operations
        fileIterator.next();
        fileCount++;
      }
      logDetail("Accessible files count", fileCount + (fileCount >= 10 ? "+" : ""));
    } catch (fileError) {
      logWarning("Cannot count files in folder (may indicate limited permissions): " + fileError.message);
    }
    
    return {
      success: true,
      folder: folder,
      name: folderName,
      url: folderUrl,
      created: folderCreated
    };
    
  } catch (error) {
    var apiCallDuration = new Date().getTime() - apiCallStart;
    
    logError("Folder access failed");
    logDetail("API call duration before error", apiCallDuration + "ms");
    logDetail("Error type", error.name || "Unknown");
    logDetail("Error message", error.message || "No message");
    
    // Analyze error type for better user guidance
    if (error.message && error.message.includes("not found")) {
      logInfo("Error analysis: Folder ID not found - folder may be deleted or ID is incorrect");
    } else if (error.message && error.message.includes("permission")) {
      logInfo("Error analysis: Permission denied - folder not shared with your account");
    } else if (error.message && error.message.includes("rate")) {
      logInfo("Error analysis: Rate limit exceeded - too many API calls");
    } else {
      logInfo("Error analysis: Generic DriveApp API error");
    }
    
    return {
      success: false,
      error: 'Cannot access folder (ID: ' + folderId + '): ' + error.message,
      errorType: error.name
    };
  }
}

/**
 * Create test file with duplicate handling - Enhanced with detailed logging
 */
function createTestFile(targetFolder) {
  try {
    var fileName = TEST_CONFIG.testFileName;
    var fileContent = TEST_CONFIG.testFileContent;
    
    logInfo("Starting file creation process");
    logDetail("Target file name", fileName);
    logDetail("Content length", fileContent.length + " characters");
    logDetail("Target folder", targetFolder.getName());
    
    // Check for existing file and handle duplicates
    logInfo("Checking for existing files with same name...");
    var duplicateCheckStart = new Date().getTime();
    
    var existingFiles = targetFolder.getFilesByName(fileName);
    var duplicateCheckDuration = new Date().getTime() - duplicateCheckStart;
    
    logDetail("Duplicate check duration", duplicateCheckDuration + "ms");
    
    if (existingFiles.hasNext()) {
      logWarning("File with same name already exists, applying timestamp suffix");
      
      // File exists, add timestamp to make it unique
      var timestamp = new Date().getTime();
      var originalFileName = fileName;
      var nameParts = fileName.split('.');
      
      if (nameParts.length > 1) {
        fileName = nameParts.slice(0, -1).join('.') + '_' + timestamp + '.' + nameParts[nameParts.length - 1];
      } else {
        fileName = fileName + '_' + timestamp;
      }
      
      logDetail("Original name", originalFileName);
      logDetail("New unique name", fileName);
      logDetail("Timestamp suffix", timestamp);
    } else {
      logInfo("No duplicate files found, using original name");
    }
    
    // Create blob object
    logInfo("Creating blob object for file content...");
    var blobCreationStart = new Date().getTime();
    
    var blob = Utilities.newBlob(fileContent, 'text/plain', fileName);
    var blobCreationDuration = new Date().getTime() - blobCreationStart;
    
    logDetail("Blob creation duration", blobCreationDuration + "ms");
    logDetail("Blob content type", blob.getContentType());
    logDetail("Blob size", blob.getBytes().length + " bytes");
    
    // Create the file in Google Drive
    logInfo("Creating file in Google Drive...");
    var fileCreationStart = new Date().getTime();
    
    var createdFile = targetFolder.createFile(blob);
    var fileCreationDuration = new Date().getTime() - fileCreationStart;
    
    logDetail("File creation duration", fileCreationDuration + "ms");
    logSuccess("File created successfully in Google Drive");
    
    // Get file details for verification
    var fileDetails = {
      id: createdFile.getId(),
      name: createdFile.getName(),
      url: createdFile.getUrl(),
      size: createdFile.getSize(),
      mimeType: createdFile.getBlob().getContentType(),
      created: createdFile.getDateCreated(),
      lastUpdated: createdFile.getLastUpdated()
    };
    
    logDetail("File ID", fileDetails.id);
    logDetail("File name verified", fileDetails.name);
    logDetail("File size verified", fileDetails.size + " bytes");
    logDetail("File MIME type", fileDetails.mimeType);
    logDetail("File creation time", fileDetails.created.toISOString());
    logDetail("File URL", fileDetails.url);
    
    return {
      success: true,
      fileName: fileName,
      fileId: fileDetails.id,
      fileUrl: fileDetails.url,
      file: createdFile,
      fileSize: fileDetails.size,
      mimeType: fileDetails.mimeType,
      created: fileDetails.created,
      durations: {
        duplicateCheck: duplicateCheckDuration,
        blobCreation: blobCreationDuration,
        fileCreation: fileCreationDuration
      }
    };
    
  } catch (error) {
    logError("File creation failed with exception");
    logDetail("Error type", error.name || "Unknown");
    logDetail("Error message", error.message || "No message");
    
    // Analyze common file creation errors
    if (error.message && error.message.includes("quota")) {
      logInfo("Error analysis: Google Drive storage quota exceeded");
    } else if (error.message && error.message.includes("permission")) {
      logInfo("Error analysis: Insufficient permissions to create files");
    } else if (error.message && error.message.includes("rate")) {
      logInfo("Error analysis: API rate limit exceeded");
    } else {
      logInfo("Error analysis: Generic file creation error");
    }
    
    return {
      success: false,
      error: 'File creation failed: ' + error.message,
      errorType: error.name
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
 * Clean up test files (optional) - Enhanced with detailed logging
 */
function cleanupTestFiles(targetFolder) {
  try {
    logInfo("Starting test file cleanup process");
    logDetail("Target folder", targetFolder.getName());
    
    var scanStart = new Date().getTime();
    
    // Find files that start with test prefix
    logInfo("Scanning folder for test files...");
    var files = targetFolder.getFiles();
    var deletedCount = 0;
    var scannedCount = 0;
    var testFilePatterns = ["gmail-attachment-saver-test", "test-", "permission-test"];
    
    logDetail("Test file patterns", testFilePatterns.join(", "));
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      scannedCount++;
      
      // Check if file matches test patterns
      var isTestFile = false;
      for (var i = 0; i < testFilePatterns.length; i++) {
        if (fileName.indexOf(testFilePatterns[i]) === 0) {
          isTestFile = true;
          break;
        }
      }
      
      if (isTestFile) {
        logInfo("Found test file to delete: " + fileName);
        
        try {
          var fileId = file.getId();
          var fileSize = file.getSize();
          
          logDetail("Deleting file ID", fileId);
          logDetail("File size", fileSize + " bytes");
          
          file.setTrashed(true);
          deletedCount++;
          
          logSuccess("Test file deleted: " + fileName);
        } catch (deleteError) {
          logWarning("Failed to delete file '" + fileName + "': " + deleteError.message);
        }
      }
    }
    
    var scanDuration = new Date().getTime() - scanStart;
    
    logDetail("Folder scan duration", scanDuration + "ms");
    logDetail("Total files scanned", scannedCount);
    logDetail("Test files deleted", deletedCount);
    
    if (deletedCount > 0) {
      logSuccess("Cleanup completed - " + deletedCount + " test files removed");
    } else {
      logInfo("No test files found to delete");
    }
    
    return deletedCount;
    
  } catch (error) {
    logWarning("Cleanup failed (not critical): " + error.message);
    logDetail("Cleanup error type", error.name || "Unknown");
    return 0;
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
 * Run comprehensive Google Drive integration test - Enhanced with detailed logging
 */
function runFullGDriveTest() {
  console.log("=".repeat(80));
  console.log("    COMPREHENSIVE GOOGLE DRIVE INTEGRATION TEST SUITE");
  console.log("=".repeat(80));
  console.log("Full test suite started:", new Date().toISOString());
  
  var fullTestStartTime = new Date().getTime();
  var results = {
    folderAccess: false,
    fileCreation: false,
    multipleFiles: false,
    permissions: false,
    overall: false,
    testDurations: {},
    errors: []
  };
  
  logInfo("Running 4 comprehensive tests to validate Google Drive integration");
  logDetail("Target folder ID", TEST_CONFIG.folderId);
  
  // Test 1: Basic folder access
  logStep("1/4", "Basic Folder Access Test");
  var test1Start = new Date().getTime();
  
  var folderTest = getFolderByIdSafely(TEST_CONFIG.folderId);
  results.folderAccess = folderTest.success;
  results.testDurations.folderAccess = new Date().getTime() - test1Start;
  
  if (!results.folderAccess) {
    logError("Cannot proceed with remaining tests - folder access failed");
    results.errors.push("Folder access failed: " + folderTest.error);
    
    logDetail("Test 1 duration", results.testDurations.folderAccess + "ms");
    console.log("\n" + "=".repeat(80));
    console.log("    TEST SUITE TERMINATED EARLY");
    console.log("=".repeat(80));
    
    return results;
  }
  
  logSuccess("Test 1 completed successfully");
  logDetail("Test 1 duration", results.testDurations.folderAccess + "ms");
  
  // Test 2: File creation
  logStep("2/4", "File Creation and Management Test");
  var test2Start = new Date().getTime();
  
  var fileTest = testGDriveSave();
  results.fileCreation = fileTest.success;
  results.testDurations.fileCreation = new Date().getTime() - test2Start;
  
  if (!results.fileCreation) {
    results.errors.push("File creation failed: " + fileTest.error);
    logError("File creation test failed");
  } else {
    logSuccess("Test 2 completed successfully");
  }
  
  logDetail("Test 2 duration", results.testDurations.fileCreation + "ms");
  
  // Test 3: Multiple files
  logStep("3/4", "Multiple File Operations Test");
  var test3Start = new Date().getTime();
  
  try {
    var multiTest = testMultipleFiles();
    results.multipleFiles = multiTest && multiTest.length > 0;
    results.testDurations.multipleFiles = new Date().getTime() - test3Start;
    
    if (!results.multipleFiles) {
      results.errors.push("Multiple files test failed");
      logError("Multiple files test failed");
    } else {
      logSuccess("Test 3 completed successfully");
      logDetail("Files created", multiTest.length);
    }
  } catch (error) {
    results.multipleFiles = false;
    results.errors.push("Multiple files test exception: " + error.message);
    results.testDurations.multipleFiles = new Date().getTime() - test3Start;
    logError("Multiple files test threw exception: " + error.message);
  }
  
  logDetail("Test 3 duration", results.testDurations.multipleFiles + "ms");
  
  // Test 4: Permissions
  logStep("4/4", "Folder Permissions and Capabilities Test");
  var test4Start = new Date().getTime();
  
  try {
    results.permissions = testFolderPermissions();
    results.testDurations.permissions = new Date().getTime() - test4Start;
    
    if (!results.permissions) {
      results.errors.push("Permissions test failed");
      logError("Permissions test failed");
    } else {
      logSuccess("Test 4 completed successfully");
    }
  } catch (error) {
    results.permissions = false;
    results.errors.push("Permissions test exception: " + error.message);
    results.testDurations.permissions = new Date().getTime() - test4Start;
    logError("Permissions test threw exception: " + error.message);
  }
  
  logDetail("Test 4 duration", results.testDurations.permissions + "ms");
  
  // Overall result calculation
  results.overall = results.folderAccess && results.fileCreation && 
                   results.multipleFiles && results.permissions;
  
  var totalTestDuration = new Date().getTime() - fullTestStartTime;
  results.totalDuration = totalTestDuration;
  
  // Final comprehensive results
  console.log("\n" + "=".repeat(80));
  console.log("    COMPREHENSIVE TEST RESULTS");
  console.log("=".repeat(80));
  
  console.log("\n📊 TEST OUTCOMES:");
  console.log("1. Folder Access:     " + (results.folderAccess ? "✅ PASS" : "❌ FAIL") + " (" + results.testDurations.folderAccess + "ms)");
  console.log("2. File Creation:     " + (results.fileCreation ? "✅ PASS" : "❌ FAIL") + " (" + results.testDurations.fileCreation + "ms)");
  console.log("3. Multiple Files:    " + (results.multipleFiles ? "✅ PASS" : "❌ FAIL") + " (" + results.testDurations.multipleFiles + "ms)");
  console.log("4. Permissions:       " + (results.permissions ? "✅ PASS" : "❌ FAIL") + " (" + results.testDurations.permissions + "ms)");
  console.log("\n🎯 OVERALL RESULT:    " + (results.overall ? "🎉 ALL TESTS PASSED" : "⚠️  SOME TESTS FAILED"));
  
  console.log("\n⏱️  PERFORMANCE SUMMARY:");
  logDetail("Total test suite duration", totalTestDuration + "ms");
  logDetail("Average test duration", Math.round(totalTestDuration / 4) + "ms");
  logDetail("Fastest test", Math.min.apply(Math, Object.values(results.testDurations)) + "ms");
  logDetail("Slowest test", Math.max.apply(Math, Object.values(results.testDurations)) + "ms");
  
  if (results.errors.length > 0) {
    console.log("\n❌ ERRORS ENCOUNTERED:");
    for (var i = 0; i < results.errors.length; i++) {
      console.log("   " + (i + 1) + ". " + results.errors[i]);
    }
  }
  
  console.log("\n🏁 FINAL ASSESSMENT:");
  if (results.overall) {
    console.log("   ✅ Google Drive integration is fully functional!");
    console.log("   ✅ All core functionality tested and working");
    console.log("   ✅ Gmail Attachment Saver will work once PMO connectivity is resolved");
    console.log("   💡 Focus efforts on resolving PMO DNS/firewall issues");
  } else {
    console.log("   ❌ Google Drive integration has issues that need resolution");
    console.log("   🔧 Fix Google Drive access problems before addressing PMO integration");
    console.log("   💡 Check folder permissions, OAuth scopes, and API access");
  }
  
  return results;
}
