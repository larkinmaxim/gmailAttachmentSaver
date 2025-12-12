# Gmail Attachment Saver

A Google Apps Script Gmail add-on that integrates with Jira to save email attachments directly to Google Drive.

## Features

- **Gmail Integration**: Contextual Gmail add-on that appears when viewing emails with attachments
- **Jira Integration**: Connects to Jira to fetch your active Technical Project Manager tickets
- **Smart Folder Lookup**: Automatically parses Jira ticket summaries to find/create customer and project folders
- **Subfolder Organization**: Save attachments to organized subfolders
- **Attachment Selection**: Select specific attachments to save, with memory across sessions
- **Duplicate Handling**: Intelligently handles duplicate files by checking size and adding timestamps when needed
- **Shared Drive Support**: Full support for Google Shared Drives using Advanced Drive API v3
- **Settings Management**: Secure storage of Jira credentials and custom JQL queries

## Setup

### Prerequisites
- Google Apps Script environment
- Jira instance with API access
- Jira API token

### Configuration

1. Instalation steps as internal app needs to be added

2. **Configure Integration Settings**:
   - Open the add-on in Gmail
   - Go to Settings (⚙️)
   - **Jira Configuration**:
     - Enter your Jira URL (e.g., `https://support.transporeon.com`)
     - Enter your [Jira API token]
     - Customize JQL query if needed (default focuses on TPM tickets)
   - **Test Connection**: Click "Test Jira Connection" to verify settings
   - **Test Folder Lookup**: Click "Test Folder Lookup" to verify folder access

## Usage

1. **Open an email with attachments** in Gmail
2. **Click the add-on** in the sidebar (labeled "Jira Project")
3. **Select your Jira ticket** from the dropdown or enter manually
4. **Choose attachments** to save (checkboxes grouped by file type)
5. **(Optional) Select project sub folder** for better organization
6. **Click "Save to Project Folder"** to save selected attachments

### Smart Folder Organization

The add-on automatically parses Jira ticket summaries to determine the correct folder location:

**Jira Summary Format:**
```
[Customer ID] Customer Name | Project Type | Description
Example: 620254 - Frosta AG | TP Essential | Ocean Visibility and TPE
```

**Folder Structure:**
```
📁 Customers Folder/
└── 📁 620254 - Frosta AG/
    └── 📁 CXPRODELIVERY-1234/
        ├── 📁 01_System_Design/
        ├── 📁 02_Meet_Recordings/
        ├── 📁 03_Correspondence/
        ├── 📁 04_Project_Documentation/
        │   ├── 📁 Project_Management/
        │   ├── 📁 Custom_Bundle/
        │   └── 📁 Carrier_Onboarding/
        ├── document1.pdf
        └── image1.png
```

**Benefits:**
- **Automatic Folder Detection**: Parses ticket summaries to find correct folders
- **Intelligent Creation**: Creates customer and project folders if they don't exist
- **Subfolder Organization**: Save to standardized subfolders for better structure
- **Team Access**: Folders can be shared across team members
- **Shared Drive Support**: Works with Google Shared Drives

## Features in Detail

### Ticket Selection
- **Dynamic Dropdown**: Shows active TPM tickets fetched from Jira
- **Ticket Details**: View full ticket information when selected
- **Manual Entry**: Enter ticket numbers directly if needed
- **Status Indicators**: Visual status emojis for different ticket states

### Attachment Management
- **Grouped by Type**: Attachments organized by file extension
- **Size Display**: File sizes shown for each attachment
- **Selection Memory**: Remembers your selections during the session
- **Duplicate Prevention**: Skips identical files, renames if different sizes

### Jira Integration
- **Custom JQL**: Configure custom queries to filter your tickets
- **TPM Focus**: Default query targets Technical Project Manager assignments
- **Connection Testing**: Verify Jira connectivity and credentials
- **Secure Storage**: API tokens stored securely in user properties
- **Summary Parsing**: Automatically extracts customer and project information from ticket summaries

### Smart Folder Management
- **Automatic Detection**: Parses Jira summary to find customer and project folders
- **Intelligent Search**: Caches folder searches for improved performance
- **Auto-Creation**: Creates folder structures when they don't exist
- **Subfolder Support**: Organize files into standardized project subfolders
- **Shared Drive Compatible**: Full support for Google Shared Drives via Advanced Drive API v3

## Default JQL Query

The default JQL query targets TPM tickets:
```jql
project = CXPRODELIVERY 
AND issuetype in (Project, "Project (Standard Solution)") 
AND status in (HYPERCARE, "Order received", "Test system available", ...) 
AND "Technical Project Manager" in (currentUser())
```

## Development

### File Structure
- `Code.js` - Main Google Apps Script code (4142 lines)
- `appsscript.json` - Project configuration and permissions
- `README.md` - Main documentation
- `Documentation/` - Detailed process and technical documentation
  - Process guides (Jira integration, folder lookup, attachment handling)
  - Function reference
  - Feature implementation summaries
  - `archive/` - Historical troubleshooting documentation

### Key Functions
- `buildAddOn()` - Main entry point for Gmail add-on
- `showTicketDetails()` - Dynamic ticket selection handler
- `saveSelectedAttachmentsToGDrive()` - Core save functionality with subfolder support
- `getMyJiraProjects()` - Jira API integration
- `completeJiraToFolderWorkflowWithCreate()` - Smart folder lookup/creation workflow
- `parseJiraSummaryForFolders()` - Jira summary parser for folder names
- `saveAttachmentWithAdvancedDrive()` - Shared Drive compatible file saving


