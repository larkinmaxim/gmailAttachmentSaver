# Customer Folder Link in Selected Ticket Details

## Implementation Date
December 3, 2025

## Overview
Added a clickable customer folder link to the "📋 Selected Ticket Details" section in the Gmail Add-on. When a user selects a Jira ticket, they can now see and access the associated customer folder directly from the ticket details card.

## Changes Made

### Location
**File:** `Code.js`  
**Function:** `showTicketDetails()`  
**Lines:** 2044-2084

### Implementation Details

#### 1. Folder Lookup
- Calls `completeJiraToFolderWorkflowWithCreate(selectedTicket.key, false)` 
- The `false` parameter means **search only, don't create**
- Uses existing optimized search with caching
- Extracts customer folder info if it exists

#### 2. Display Logic

**If Customer Folder EXISTS:**
```
─────────────────────
Customer Folder:
[Folder Name - e.g., "C12345 - Acme Corp"]

[📁 Open Customer Folder] ← Clickable button
```

**If Customer Folder DOESN'T EXIST:**
```
─────────────────────
Customer Folder:
C12345 - Acme Corp
(will be created on first upload)
```

**If Lookup FAILS:**
- Logs error to console
- Continues without showing customer folder
- Doesn't break the card display

#### 3. Link Behavior
- Opens in **full size** browser window/tab
- Direct link to Google Drive folder
- No action on close (user stays where they are)

## Code Structure

```javascript
// After ticket details section created (line 2042)
console.log("Looking up customer folder for ticket:", selectedTicket.key);

try {
  // Search for existing customer folder (don't create if missing)
  var folderResult = completeJiraToFolderWorkflowWithCreate(selectedTicket.key, false);
  
  if (folderResult.success && folderResult.customerFolder && folderResult.customerFolder.exists) {
    // Show clickable link to existing folder
    // - Visual separator
    // - Folder name display
    // - Button with Google Drive URL
  } else {
    // Show expected folder name with creation message
    // Uses folderResult.customerFolder.suggestedName
  }
} catch (folderError) {
  // Error handling - logs but doesn't break the card
  console.error("❌ Error looking up customer folder:", folderError.message);
}
```

## Technical Benefits

### ✅ Performance
- Leverages existing optimized folder search with caching
- Doesn't create unnecessary API calls
- Fast response using cached folder IDs

### ✅ User Experience
- Immediate access to customer folder
- Visual feedback if folder doesn't exist yet
- Graceful error handling
- Consistent with existing UI patterns

### ✅ Error Handling
- Try-catch block prevents card failure
- Informative console logging
- Falls back gracefully if lookup fails

## Folder Naming Pattern

The customer folder follows this pattern (as defined in line 536):
```
Customer ID - Customer Name
```

Example:
```
C12345 - Acme Corporation
```

This name is parsed from the Jira ticket summary using the `parseJiraSummaryForFolders()` function.

## Related Functions

### Used in Implementation
- `completeJiraToFolderWorkflowWithCreate()` (line 772)
  - Main workflow function
  - Handles Jira parsing and folder search
  - Returns folder info with URL

- `getCachedOrSearchFolder()` (called internally)
  - Fast cached folder lookup
  - Avoids redundant Drive API calls

- `parseJiraSummaryForFolders()` (line 468)
  - Extracts customer ID and name
  - Generates folder name pattern

### Related Functions
- `createCustomerFolder()` (line 651)
  - Creates new customer folder
  - Only called when uploading attachments with `createIfMissing=true`

- `findFolderByNamePattern()` (line 316)
  - Low-level folder search
  - Used by the workflow function

## Testing Scenarios

### Scenario 1: Existing Customer Folder
**Expected:** 
- ✅ Displays folder name
- ✅ Shows clickable "📁 Open Customer Folder" button
- ✅ Button opens Google Drive to correct folder

### Scenario 2: New Customer (No Folder Yet)
**Expected:**
- ✅ Displays expected folder name
- ✅ Shows message "(will be created on first upload)"
- ✅ No clickable button (folder doesn't exist yet)

### Scenario 3: Jira API Error
**Expected:**
- ✅ Logs error to console
- ✅ Card still displays other ticket details
- ✅ Customer folder section not shown

### Scenario 4: Manual Ticket Entry
**Expected:**
- ✅ Same behavior as dropdown selection
- ✅ Folder lookup works with manually entered ticket numbers

## User Workflow

1. User opens Gmail Add-on
2. User selects a Jira ticket from dropdown (or enters manually)
3. `showTicketDetails()` is triggered
4. Ticket details card displays with:
   - Ticket key and status
   - Summary
   - Status and Type
   - **NEW:** Customer folder link (if exists)
5. User clicks "📁 Open Customer Folder"
6. Google Drive opens in new tab showing customer folder

## Future Enhancements (Optional)

### Potential Additions
1. **Project Folder Link**: Add link to project folder as well
2. **Folder Stats**: Show number of files/subfolders
3. **Recent Activity**: Show last upload date
4. **Create Now Button**: Allow manual folder creation before upload

### Not Implemented (By Design)
- ❌ Automatic folder creation on ticket selection
  - Reason: Folders should only be created when needed (on upload)
- ❌ Folder contents preview
  - Reason: Would require additional API calls and slow down card display

## Maintenance Notes

### If Customer Folder Logic Changes
- Update `parseJiraSummaryForFolders()` for new naming patterns
- Update `completeJiraToFolderWorkflowWithCreate()` for new search logic
- This implementation will automatically use updated logic

### If UI Needs Changes
- Modify lines 2054-2078 in `showTicketDetails()`
- Separator: Line 2055-2056
- Button text: Line 2060
- Folder name display: Line 2067-2068
- "Will be created" message: Line 2077-2078

## Rollback Instructions

If needed, revert to previous behavior by removing lines 2044-2084 and replacing with:
```javascript
console.log("Ticket details section created");

// Handle Gmail context for attachments (threadId already determined above)
```

---

**Status:** ✅ Implemented and Tested  
**No Breaking Changes**  
**No Linter Errors**

