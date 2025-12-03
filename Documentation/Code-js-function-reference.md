# Code.js Function Reference

A comprehensive summary of all functions in the Gmail Attachment Saver add-on.

---

## 📋 Table of Contents

1. [Script Properties & Configuration](#script-properties--configuration)
2. [Jira Summary Parsing & Folder Lookup](#jira-summary-parsing--folder-lookup)
3. [Gmail Add-on UI](#gmail-add-on-ui)
4. [Settings Management](#settings-management)
5. [Google Drive Operations](#google-drive-operations)
6. [Jira Integration](#jira-integration)
7. [Google Docs Link Handling](#google-docs-link-handling)
8. [Attachment Management](#attachment-management)
9. [Helper & Utility Functions](#helper--utility-functions)
10. [Debug & Testing Functions](#debug--testing-functions)

---

## Script Properties & Configuration

| Function | Summary |
|----------|---------|
| `setupScriptProperties()` | One-time setup function to configure authentication credentials and environment variables in Script Properties. |
| `getDefaultJiraURL()` | Returns the default Jira URL from Script Properties or fallback value. |
| `getSettingsStorageKey()` | Returns the key used for storing user settings (default: `JIRA_SETTINGS`). |
| `getCustomersFolderId()` | Returns the customers folder ID from Script Properties or fallback value. |

---

## Jira Summary Parsing & Folder Lookup

| Function | Summary |
|----------|---------|
| `getJiraTicketFolderNames(ticketKey)` | Fetches Jira ticket and parses summary to extract customer ID, name, and project type. |
| `parseJiraSummaryForFolders(summary, ticketKey)` | Parses Jira summary string to extract folder names for customer and project. |
| `sanitizeFolderName(folderName)` | Cleans folder names by removing invalid characters and trimming whitespace. |
| `completeJiraToFolderWorkflowWithCreate(ticketKey, createIfMissing)` | Complete workflow: parses ticket, searches for folders, creates if needed. |
| `getCachedOrSearchFolder(cacheKey, parentFolderId, searchPattern)` | Searches for folder with caching to improve performance. |
| `createProjectFolderStructure(parentFolderId, projectFolderName)` | Creates complete project folder structure with standard subfolders. |
| `createCustomerFolder(customerFolderName)` | Creates a new customer folder in the Customers root folder. |

---

## Gmail Add-on UI

| Function | Summary |
|----------|---------|
| `buildAddOn(e)` | Main entry point that builds the Gmail add-on card with ticket selection, attachments list, and save button. |
| `buildSettingsCard(isFirstTime)` | Creates the settings configuration card with Jira URL, token, and JQL inputs. |
| `showSettings(e)` | Action handler that displays the settings card when user clicks settings button. |
| `showTicketDetails(e)` | Updates the card to show full ticket details when user selects a ticket from dropdown. |
| `buildRequestAccessCard(ticketKey, folderId, errorMessage)` | Creates a card prompting user to request access to a Google Drive folder they can't access. |

---

## Settings Management

| Function | Summary |
|----------|---------|
| `saveSettings(e)` | Saves user's Jira URL, API token, and custom JQL query to user properties. |
| `getUserSettings()` | Retrieves user settings from storage with PMO defaults for new users. |
| `setUserSettings(settings)` | Stores user settings JSON to user properties. |
| `getDefaultJQL()` | Returns the default JQL query for filtering TPM project tickets. |
| `maskToken(token)` | Returns a masked version of API token for display (shows first/last 4 chars). |
| `getUserSubfolderPreferences()` | Returns user preferences for subfolder selection feature. |
| `shouldUseEnhancedUI()` | Checks if enhanced UI with subfolder selection should be displayed. |

---

## Google Drive Operations

| Function | Summary |
|----------|---------|
| `getFolderByIdSafely(folderId)` | Safely accesses a folder by ID with Advanced Drive API, detecting Shared Drive vs My Drive. |
| `getOrCreateFolder(folderName, parentFolder)` | Gets existing folder or creates new one in the specified parent folder. |
| `getProjectSubfolderConfig()` | Returns standardized project subfolder structure definitions. |
| `getOrCreateProjectSubfolder(parentFolder, subfolderPath)` | Creates nested subfolder structure (e.g., `04_Project_Documentation/Project_Management`). |
| `validateSubfolderSelection(subfolderPath)` | Validates and sanitizes subfolder path against allowed patterns. |
| `handleSubfolderCreationFailure(projectFolder, subfolderPath, errorDetails)` | Fallback strategy when subfolder creation fails (uses simplified path or project root). |
| `saveAttachmentWithAdvancedDrive(attachment, folderId, fileName)` | Saves file using Advanced Drive API v3 with `supportsAllDrives: true` for Shared Drive support. |

---

## Jira Integration

| Function | Summary |
|----------|---------|
| `getMyJiraProjects()` | Fetches user's TPM project tickets from Jira using configured JQL query. |
| `testJiraAPI(settings)` | Tests Jira API connection by calling `/rest/api/2/myself` endpoint. |
| `testJiraConnection(e)` | Action handler for "Test Jira Connection" button in settings. |

---

## Google Docs Link Handling

| Function | Summary |
|----------|---------|
| `extractGoogleDocsLinks(message)` | Scans email content for Google Docs/Sheets/Slides/Forms/Drive links. |
| `extractTitleFromContext(content, url, isHtml)` | Attempts to extract document title from surrounding HTML/text context. |
| `copyGoogleDocToFolder(docLink, targetFolder)` | Copies a Google Doc to target folder using the email subject as filename. |
| `sanitizeFileName(subject)` | Cleans email subject for use as filename (removes invalid characters). |
| `extractGoogleDocId(url)` | Extracts document ID from various Google Docs URL formats. |
| `createFallbackShortcut(docLink, targetFolder)` | Creates a `.url` shortcut file when Google Doc copy fails. |
| `getGoogleDocsIcon(docType)` | Returns appropriate emoji icon for each Google Docs type. |

---

## Attachment Management

| Function | Summary |
|----------|---------|
| `saveSelectedAttachmentsToGDrive(e)` | Main save function: processes selected attachments and saves to PMO project folder. |
| `storeAttachmentSelections(threadId, selectionState)` | Persists user's attachment checkbox selections for a specific email thread. |
| `getStoredAttachmentSelections(threadId)` | Retrieves previously stored attachment selections for an email thread. |

---

## Helper & Utility Functions

| Function | Summary |
|----------|---------|
| `formatCompactTicketDisplay(issue)` | Formats ticket for dropdown display: emoji + key + client name. |
| `getStatusEmoji(status)` | Returns appropriate emoji for each Jira issue status. |
| `formatFileSize(bytes)` | Converts bytes to human-readable format (B, KB, MB, GB). |
| `backToMain(e)` | Navigation handler to return from settings to main card. |
| `getLogLevelName(level)` | Returns human-readable name for log level number. |
| `verboseLog(message)` | Conditional logging based on configured log level. |

---

## Debug & Testing Functions

| Function | Summary |
|----------|---------|
| `testJiraConnection(e)` | Action handler for "Test Jira Connection" button in settings UI. |
| `testFolderLookup(e)` | Action handler for "Test Folder Lookup" button to verify customer folder access. |
| `reauthorizeUser()` | Comprehensive OAuth reauthorization test with detailed reporting. |
| `forceReauthorization()` | Triggers OAuth reauthorization by accessing all required services. |
| `logDiagnosticInfo()` | Outputs comprehensive diagnostic information to console for troubleshooting. |

---

## Function Call Flow

### Main Save Operation Flow

```
buildAddOn() 
  → showTicketDetails()
    → saveSelectedAttachmentsToGDrive()
      → completeJiraToFolderWorkflowWithCreate()
        → getJiraTicketFolderNames()
        → parseJiraSummaryForFolders()
        → getCachedOrSearchFolder()
        → [if not found] createProjectFolderStructure()
      → getFolderByIdSafely()
      → getOrCreateProjectSubfolder()
      → saveAttachmentWithAdvancedDrive()
```

### Jira Summary Parsing Flow

```
completeJiraToFolderWorkflowWithCreate(ticketKey, createIfMissing)
  → getJiraTicketFolderNames(ticketKey)
  → parseJiraSummaryForFolders(summary)
  → getCachedOrSearchFolder(customerFolder)
  → getCachedOrSearchFolder(projectFolder)
  → [if createIfMissing=true] createProjectFolderStructure()
  → return { customerFolder, projectFolder, success }
```

### Settings Flow

```
buildAddOn()
  → getUserSettings()
  → [if not configured] → buildSettingsCard(true)
showSettings()
  → buildSettingsCard(false)
saveSettings()
  → setUserSettings()
```

---

## Quick Reference: Key Functions

| Action | Function |
|--------|----------|
| Save attachments | `saveSelectedAttachmentsToGDrive()` |
| Parse ticket summary | `getJiraTicketFolderNames()` |
| Find/create folder | `completeJiraToFolderWorkflowWithCreate()` |
| Save to Shared Drive | `saveAttachmentWithAdvancedDrive()` |
| Load settings | `getUserSettings()` |
| Save settings | `saveSettings()` |
| Test Jira connection | `testJiraConnection()` |
| Test folder lookup | `testFolderLookup()` |

---

*Generated from Code.js - Gmail Attachment Saver Add-on*

