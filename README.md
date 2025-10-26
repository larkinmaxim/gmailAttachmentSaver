# Gmail Attachment Saver

A Google Apps Script Gmail add-on that integrates with Jira to save email attachments directly to Google Drive.

## Features

- **Gmail Integration**: Contextual Gmail add-on that appears when viewing emails with attachments
- **Jira Integration**: Connects to Jira to fetch your active Technical Project Manager tickets
- **PMO Integration**: Automatically connects to centralized PMO-managed project folders (replaces local folder creation)
- **Attachment Selection**: Select specific attachments to save, with memory across sessions
- **Duplicate Handling**: Intelligently handles duplicate files by checking size and adding timestamps when needed
- **Settings Management**: Secure storage of Jira credentials, PMO webhook configuration, and custom JQL queries

## Setup

### Prerequisites
- Google Apps Script environment
- Jira instance with API access
- Jira API token

### Configuration

1. **Deploy the Google Apps Script**:
   - Open [Google Apps Script](https://script.google.com/)
   - Create a new project or use an existing one
   - Copy the code from `Code.js` into your script
   - Copy the configuration from `appsscript.json`

2. **Configure Integration Settings**:
   - Open the add-on in Gmail
   - Go to Settings (⚙️)
   - **Jira Configuration**:
     - Enter your Jira URL (e.g., `https://your-company.atlassian.net`)
     - Enter your [Jira API token](https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/)
     - Customize JQL query if needed (default focuses on TPM tickets)
   - **PMO Configuration**:
     - PMO webhook URL (pre-configured with default)
     - Timeout and retry settings
   - Test both connections

### Required OAuth Scopes

The add-on requires the following permissions:
- `https://www.googleapis.com/auth/gmail.readonly` - Read Gmail messages
- `https://www.googleapis.com/auth/drive` - Create folders and files in Google Drive
- `https://www.googleapis.com/auth/script.locale` - Localization
- `https://www.googleapis.com/auth/gmail.addons.execute` - Gmail add-on functionality
- `https://www.googleapis.com/auth/script.external_request` - Jira API calls

## Usage

1. **Open an email with attachments** in Gmail
2. **Click the add-on** in the sidebar
3. **Select your Jira ticket** from the dropdown or enter manually
4. **Choose attachments** to save (checkboxes grouped by file type)
5. **Click "Save to PMO Folder"** to save selected attachments

### PMO Folder Integration

Attachments are automatically saved to PMO-managed project folders:
```
PMO System/
└── [PMO-Created Project Folders]/
    └── CXPRODELIVERY-{ticket-number}/     (Auto-created by PMO)
        ├── document1.pdf
        ├── image1.png
        └── ...
```

**Benefits:**
- **Centralized Management**: All project files in PMO-standardized locations
- **Automatic Creation**: PMO system creates folders as needed
- **Team Access**: Shared access across project team members
- **Enterprise Compliance**: Meets organizational data governance requirements

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

### PMO Integration
- **Centralized Folders**: Automatic connection to PMO-managed project folders
- **Webhook Integration**: Real-time folder lookup and creation via PMO webhook
- **Auto-Creation**: PMO system creates project folders automatically when needed
- **Retry Logic**: Intelligent retry handling for folder creation timing issues
- **Enhanced Error Handling**: User-friendly error messages for PMO system issues
- **Connection Testing**: Comprehensive PMO webhook connectivity testing

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
- `Code.js` - Main Google Apps Script code
- `appsscript.json` - Project configuration and permissions
- `README.md` - Documentation
- `.gitignore` - Git ignore patterns

### Key Functions
- `buildAddOn()` - Main entry point for Gmail add-on
- `showTicketDetails()` - Dynamic ticket selection handler
- `saveSelectedAttachmentsToGDrive()` - Core save functionality
- `getMyJiraProjects()` - Jira API integration

## License

This project is licensed under the MIT License.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Support

For issues related to:
- **Google Apps Script**: Check the [Apps Script documentation](https://developers.google.com/apps-script)
- **Jira API**: Refer to the [Jira REST API documentation](https://developer.atlassian.com/cloud/jira/platform/rest/v2/)
- **Gmail Add-ons**: See [Gmail Add-on guide](https://developers.google.com/gmail/add-ons)
