# Deployment Guide: Gmail Attachment Saver to GCP

This guide walks you through deploying the Gmail Attachment Saver add-on to a new Google Cloud Platform (GCP) project and publishing it as an Internal App on Google Workspace Marketplace.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Phase 1: Set Up GCP Project](#phase-1-set-up-gcp-project)
3. [Phase 2: Configure OAuth Consent Screen](#phase-2-configure-oauth-consent-screen)
4. [Phase 3: Deploy Apps Script Project](#phase-3-deploy-apps-script-project)
5. [Phase 4: Test the Add-on](#phase-4-test-the-add-on)
6. [Phase 5: Prepare for Marketplace Publishing](#phase-5-prepare-for-marketplace-publishing)
7. [Phase 6: Publish to Google Workspace Marketplace (Internal)](#phase-6-publish-to-google-workspace-marketplace-internal)
8. [Phase 7: Post-Deployment Configuration](#phase-7-post-deployment-configuration)
9. [Troubleshooting](#troubleshooting)

---

## Prerequisites

Before starting, ensure you have:

- ✅ **Google Workspace Admin Account** (for internal app publishing)
- ✅ **GCP Project Creation Rights** in your organization
- ✅ **Access to Google Cloud Console** (console.cloud.google.com)
- ✅ **clasp CLI installed** (optional but recommended): `npm install -g @google/clasp`
- ✅ **Jira Instance** with API access and credentials
- ✅ **Google Drive Customers Folder ID** (the parent folder for storing attachments)
- ✅ **Project Files**:
  - `Code.js` (main script)
  - `appsscript.json` (configuration)
  - `drive_2020q4_32dp.png` (logo)

---

## Phase 1: Set Up GCP Project

### Step 1.1: Create New GCP Project

1. **Navigate to Google Cloud Console**
   - Go to: https://console.cloud.google.com/
   - Sign in with your Google Workspace account

2. **Create New Project**
   - Click the project dropdown at the top
   - Click **"New Project"**
   - Enter project details:
     - **Project Name**: `Gmail Attachment Saver` (or your preferred name)
     - **Organization**: Select your organization
     - **Location**: Select appropriate parent folder
   - Click **"Create"**
   - Wait for project creation (takes ~30 seconds)
   - Note down your **Project ID** (e.g., `gmail-attachment-saver-123456`)

### Step 1.2: Enable Required APIs

1. **Navigate to APIs & Services**
   - In the GCP Console, go to: **Navigation Menu (☰) → APIs & Services → Library**

2. **Enable the following APIs** (search and enable each):
   - ✅ **Gmail API**
     - Search: "Gmail API"
     - Click the API
     - Click **"Enable"**
   
   - ✅ **Google Drive API**
     - Search: "Google Drive API"
     - Click the API
     - Click **"Enable"**
   
   - ✅ **Apps Script API**
     - Search: "Apps Script API"
     - Click the API
     - Click **"Enable"**
   
   - ✅ **Google Workspace Marketplace SDK** (for publishing)
     - Search: "Google Workspace Marketplace SDK"
     - Click the SDK
     - Click **"Enable"**

3. **Verify APIs are Enabled**
   - Go to: **APIs & Services → Dashboard**
   - Confirm all 4 APIs show as "Enabled"

---

## Phase 2: Configure OAuth Consent Screen

This is crucial for user authentication and app authorization.

### Step 2.1: Configure OAuth Consent Screen

1. **Navigate to OAuth Consent Screen**
   - Go to: **APIs & Services → OAuth consent screen**

2. **Select User Type**
   - Choose **"Internal"** (for Google Workspace organization only)
   - Click **"Create"**

3. **Fill App Information** (Page 1 of 4)
   - **App name**: `Jira Attachment Saver`
   - **User support email**: Your email address
   - **App logo**: Upload `drive_2020q4_32dp.png` (optional but recommended)
   - **Application home page**: (leave blank or add documentation URL)
   - **Application privacy policy link**: (recommended - add your organization's policy)
   - **Application terms of service link**: (optional)
   - **Authorized domains**: Add your domain (e.g., `transporeon.com`)
   - **Developer contact information**: Your email address
   - Click **"Save and Continue"**

4. **Configure Scopes** (Page 2 of 4)
   - Click **"Add or Remove Scopes"**
   - **Manually add these scopes** (paste in the "Manually add scopes" field):
     ```
     https://www.googleapis.com/auth/gmail.readonly
     https://www.googleapis.com/auth/drive
     https://www.googleapis.com/auth/script.locale
     https://www.googleapis.com/auth/gmail.addons.execute
     https://www.googleapis.com/auth/script.external_request
     https://www.googleapis.com/auth/gmail.addons.current.message.metadata
     https://www.googleapis.com/auth/gmail.addons.current.message.readonly
     https://www.googleapis.com/auth/userinfo.email
     ```
   - Click **"Add to Table"**
   - Verify all 8 scopes appear in the table
   - Click **"Update"**
   - Click **"Save and Continue"**

5. **Summary** (Page 3 of 4)
   - Review all information
   - Click **"Back to Dashboard"**

### Step 2.2: Note OAuth Client Details

- You'll create OAuth credentials automatically when linking Apps Script (next phase)
- No manual OAuth client creation needed for Apps Script add-ons

---

## Phase 3: Deploy Apps Script Project

### Step 3.1: Create Apps Script Project

**Option A: Using Apps Script Web IDE (Recommended for first deployment)**

1. **Create New Project**
   - Go to: https://script.google.com/
   - Click **"New Project"**
   - Name the project: `Jira Attachment Saver`

2. **Link to GCP Project**
   - In Apps Script editor, click **Project Settings** (⚙️) in left sidebar
   - Scroll to **"Google Cloud Platform (GCP) Project"**
   - Click **"Change project"**
   - Enter your **GCP Project Number** (not Project ID):
     - To find Project Number: Go to GCP Console → Dashboard → Project info card
     - Or go to: https://console.cloud.google.com/home/dashboard
     - Copy the **Project Number** (numeric value)
   - Paste the Project Number
   - Click **"Set Project"**

3. **Upload Project Files**
   - Click on **"Editor"** (<>) in left sidebar
   - Delete the default `Code.gs` file content
   - Copy entire contents of your local `Code.js` and paste it
   - Rename `Code.gs` to `Code.js` if desired (optional)
   
4. **Configure appsscript.json**
   - Click **Project Settings** (⚙️)
   - Check ✅ **"Show "appsscript.json" manifest file in editor"**
   - Go back to **Editor**
   - Click on `appsscript.json` in the file list
   - Replace its content with your local `appsscript.json` content:

   ```json
   {
     "timeZone": "Europe/Berlin",
     "runtimeVersion": "V8",
     "dependencies": {
       "enabledAdvancedServices": [
         {
           "userSymbol": "Gmail",
           "version": "v1",
           "serviceId": "gmail"
         },
         {
           "userSymbol": "Drive",
           "version": "v3",
           "serviceId": "drive"
         }
       ]
     },
     "addOns": {
       "common": {
         "name": "Jira Attachment Saver",
         "logoUrl": "https://www.gstatic.com/images/branding/product/1x/drive_2020q4_32dp.png",
         "useLocaleFromApp": true
       },
       "gmail": {
         "contextualTriggers": [
           {
             "unconditional": {},
             "onTriggerFunction": "buildAddOn"
           }
         ]
       }
     },
     "oauthScopes": [
       "https://www.googleapis.com/auth/gmail.readonly",
       "https://www.googleapis.com/auth/drive",
       "https://www.googleapis.com/auth/script.locale",
       "https://www.googleapis.com/auth/gmail.addons.execute",
       "https://www.googleapis.com/auth/script.external_request",
       "https://www.googleapis.com/auth/gmail.addons.current.message.metadata",
       "https://www.googleapis.com/auth/gmail.addons.current.message.readonly",
       "https://www.googleapis.com/auth/userinfo.email"
     ]
   }
   ```

5. **Save All Files**
   - Click **Save** icon (💾) or press `Ctrl+S`

**Option B: Using clasp CLI (Alternative method)**

```bash
# Login to clasp
clasp login

# Create new project
clasp create --type standalone --title "Jira Attachment Saver"

# Push your local files
clasp push

# Open in browser to link GCP project
clasp open
```

### Step 3.2: Configure Script Properties

1. **Run Setup Function**
   - In Apps Script editor, select function: `setupScriptProperties` from dropdown
   - Click **Run** ▶️
   - **First time**: You'll be prompted to authorize the script
   - Click **Review Permissions**
   - Select your account
   - Click **Advanced** → **Go to Jira Attachment Saver (unsafe)**
   - Click **Allow**

2. **Verify Script Properties**
   - After successful run, go to **Project Settings** (⚙️)
   - Scroll to **"Script Properties"**
   - Verify these properties exist:
     - `DEFAULT_JIRA_URL`: `https://support.transporeon.com`
     - `SETTINGS_STORAGE_KEY`: `JIRA_SETTINGS`
     - `CUSTOMERS_FOLDER_ID`: `1pi6fJzg5nncV7ADrcivI0gEAa9cjZLQp`
     - `DEFAULT_LOG_LEVEL`: `2`

3. **Update for Your Environment**
   - Click **Edit script properties**
   - Update `DEFAULT_JIRA_URL` if using different Jira instance
   - Update `CUSTOMERS_FOLDER_ID` with your Google Drive folder ID:
     - Open your Customers folder in Google Drive
     - Copy folder ID from URL: `https://drive.google.com/drive/folders/[FOLDER_ID]`
   - Click **Save script properties**

### Step 3.3: Create Deployment

1. **Create New Deployment**
   - Click **Deploy** → **New deployment**
   - Click **⚙️ Select type** → **Add-on**
   - Configure deployment:
     - **Description**: `Initial deployment` or version number
     - **Version**: Auto-generated (e.g., `Version 1`)
   - Click **Deploy**

2. **Note Deployment Details**
   - Copy the **Deployment ID** (e.g., `AKfycby...`)
   - Copy the **Head Deployment ID** (used for testing)
   - Click **Done**

---

## Phase 4: Test the Add-on

### Step 4.1: Install Test Version

1. **Get Test Deployment URL**
   - In Apps Script editor, click **Deploy** → **Test deployments**
   - Click **⚙️ Select type** → **Gmail Add-on**
   - Click **Install**
   - This installs the HEAD version (latest changes) for testing

2. **Alternative: Use Installation URL**
   - Click **Deploy** → **Manage deployments**
   - Find your deployment
   - Click **Get installation URL**
   - Open URL in browser while logged in to your Google account

### Step 4.2: Test in Gmail

1. **Open Gmail**
   - Go to: https://mail.google.com/
   - Open any email with attachments

2. **Locate Add-on**
   - Look for the add-on icon in the right sidebar
   - Icon should match the logo specified in `appsscript.json`
   - Click the icon

3. **Test Functionality**
   - ✅ Add-on loads without errors
   - ✅ Settings page appears
   - ✅ Can configure Jira credentials
   - ✅ Can test Jira connection
   - ✅ Can fetch Jira tickets
   - ✅ Can select attachments
   - ✅ Can save to Google Drive

4. **Review Logs**
   - In Apps Script editor, click **Executions** (⏱️) in left sidebar
   - Review execution logs for errors
   - Check for successful function completions

---

## Phase 5: Prepare for Marketplace Publishing

### Step 5.1: Create Store Listing Assets

Prepare these assets before starting the publishing process:

1. **Application Icon**
   - Size: 128x128 pixels (PNG format)
   - Current: `drive_2020q4_32dp.png` (may need resizing)
   - Tool: Use image editor or online resizer

2. **Screenshots** (Recommended: 3-5 screenshots)
   - Size: 1280x800 pixels or 16:9 ratio
   - Show:
     - Screenshot 1: Add-on in Gmail sidebar with attachment selection
     - Screenshot 2: Jira ticket selection dropdown
     - Screenshot 3: Settings page with configuration options
     - Screenshot 4: Successful save confirmation
     - Screenshot 5: Organized folder structure in Google Drive

3. **Marketing Assets** (Optional for internal)
   - Promotional images
   - Video demo (optional)

4. **Documentation**
   - You already have `README.md`
   - Consider creating:
     - User guide (how to use the add-on)
     - Admin installation guide
     - FAQ document

### Step 5.2: Prepare Description Content

Draft your marketplace listing description:

**Short Description** (80 characters max):
```
Save Gmail attachments to Google Drive, organized by Jira project tickets.
```

**Full Description** (4000 characters max):
```
# Jira Attachment Saver

Streamline your workflow by automatically saving Gmail attachments to Google Drive, organized by your Jira projects.

## Key Features

✅ **Gmail Integration**: Contextual add-on appears when viewing emails with attachments
✅ **Jira Integration**: Connects to your Jira instance to fetch active tickets
✅ **Smart Organization**: Automatically creates organized folder structures based on Jira ticket information
✅ **Selective Saving**: Choose which attachments to save with persistent selection memory
✅ **Duplicate Handling**: Intelligently manages duplicate files
✅ **Shared Drive Support**: Full compatibility with Google Shared Drives
✅ **Secure Configuration**: User and script-level settings for team deployments

## How It Works

1. Open any email with attachments in Gmail
2. Click the add-on icon in the right sidebar
3. Select your Jira project ticket from the dropdown
4. Choose which attachments to save
5. Select project subfolder for organization (optional)
6. Click "Save to Project Folder" - done!

## Smart Folder Organization

The add-on automatically parses Jira ticket summaries to create organized folder structures:

```
📁 Customers Folder/
└── 📁 [Customer ID] - Customer Name/
    └── 📁 [JIRA-TICKET]/
        ├── 📁 01_System_Design/
        ├── 📁 02_Meet_Recordings/
        ├── 📁 03_Correspondence/
        └── 📁 04_Project_Documentation/
```

## Setup Requirements

- Google Workspace account
- Jira instance with API access
- Jira API token
- Google Drive folder for customer files

## Security & Privacy

- All credentials stored securely using Google Apps Script Properties Service
- No data transmitted to third parties
- Runs entirely within your Google Workspace environment
- Internal app - available only to your organization

## Support

For issues or feature requests, contact your IT administrator.
```

---

## Phase 6: Publish to Google Workspace Marketplace (Internal)

### Step 6.1: Access Marketplace SDK

1. **Navigate to Marketplace SDK**
   - Go to GCP Console: https://console.cloud.google.com/
   - Select your project: `Gmail Attachment Saver`
   - Go to: **Navigation Menu (☰) → APIs & Services → Google Workspace Marketplace SDK**
   - If not enabled, click **"Enable"**

2. **Start Configuration**
   - Click **"App Configuration"** tab
   - Click **"Create new listing"** or **"Edit configuration"**

### Step 6.2: Configure Application

**Section 1: App Basics**

1. **Application Name**
   - Enter: `Jira Attachment Saver`

2. **Application Description**
   - Paste your prepared full description (from Step 5.2)

3. **Application Icon**
   - Upload: `drive_2020q4_32dp.png` (128x128px)

4. **Category**
   - Select: **"Productivity"**

5. **Application URL** (Optional)
   - Leave blank or add documentation URL

**Section 2: Extensions**

1. **Add Extension**
   - Click **"+ Add Extension"**
   - Select: **"Gmail Add-on"**

2. **Configure Gmail Add-on**
   - **Apps Script Project Key**:
     - Go to Apps Script editor → Project Settings
     - Copy **"Script ID"** (e.g., `1a2b3c...`)
     - Paste in this field
   
   - **Gmail Add-on Deployment ID**:
     - In Apps Script, click Deploy → Manage deployments
     - Copy the **Deployment ID** of your production deployment
     - Paste in this field

3. **Verify Extension**
   - System should show: ✅ "Extension verified"

**Section 3: Store Listing**

1. **Screenshots**
   - Upload 3-5 screenshots (1280x800px)
   - Add captions for each screenshot

2. **Promotional Images** (Optional)
   - Upload promotional banner if available

3. **Language**
   - Select: **English (United States)** or your primary language

**Section 4: Distribution**

1. **Visibility**
   - Select: **"Private"** → **"My domain only"**
   - This makes it an **Internal App** (only for your organization)

2. **Installation Type**
   - Select: **"Available to Install"** (users choose to install)
   - OR **"Install for Everyone"** (admin pushes to all users)

3. **Authorized Domains**
   - Enter your Google Workspace domain (e.g., `transporeon.com`)

**Section 5: OAuth Scopes**

1. **Verify Scopes**
   - System should auto-detect scopes from `appsscript.json`
   - Verify all 8 scopes are listed:
     - gmail.readonly
     - drive
     - script.locale
     - gmail.addons.execute
     - script.external_request
     - gmail.addons.current.message.metadata
     - gmail.addons.current.message.readonly
     - userinfo.email

2. **Scope Justifications** (Required)
   - For each scope, provide justification:
     - **gmail.readonly**: "Read email metadata and attachment information"
     - **drive**: "Save attachments to Google Drive folders"
     - **script.locale**: "Display add-on UI in user's language"
     - **gmail.addons.execute**: "Execute Gmail add-on functionality"
     - **script.external_request**: "Connect to Jira API for ticket information"
     - **gmail.addons.current.message.metadata**: "Access current email metadata"
     - **gmail.addons.current.message.readonly**: "Read current email content and attachments"
     - **userinfo.email**: "Identify user for settings and permissions"

**Section 6: Support & Contact**

1. **Developer Name**
   - Enter: Your name or team name

2. **Developer Email**
   - Enter: Your support email

3. **Support URL** (Optional)
   - Add support documentation URL or ticketing system

4. **Privacy Policy URL** (Recommended)
   - Add your organization's privacy policy URL

5. **Terms of Service URL** (Optional)
   - Add terms of service URL if applicable

### Step 6.3: Submit for Review

1. **Review All Sections**
   - Go through each section and verify all information is correct
   - Ensure all required fields are filled

2. **Save Configuration**
   - Click **"Save"** at the bottom

3. **Submit for Internal Publishing**
   - Click **"Publish"** or **"Submit"**
   - For **internal apps**, review is typically automatic or very fast

4. **Approval Status**
   - Internal apps: Usually approved within minutes to hours
   - Check status in **"Publishing Status"** section

---

## Phase 7: Post-Deployment Configuration

### Step 7.1: Admin Installation (If Auto-Install)

If you selected "Install for Everyone":

1. **Access Google Workspace Admin Console**
   - Go to: https://admin.google.com/
   - Sign in with admin account

2. **Navigate to Apps**
   - Go to: **Apps → Google Workspace Marketplace apps**

3. **Add App**
   - Click **"+ Add app"** → **"Add from Google Workspace Marketplace"**
   - Search: `Jira Attachment Saver`
   - Click on your app
   - Click **"Install"** or **"Domain Install"**

4. **Configure Installation**
   - Select: **"Everyone"** or specific organizational units
   - Review permissions
   - Click **"Continue"** → **"Finish"**

### Step 7.2: User Installation (If Available to Install)

Provide these instructions to users:

1. **Access Gmail**
   - Go to: https://mail.google.com/

2. **Install Add-on**
   - Click **Settings (⚙️)** → **"Get add-ons"**
   - Search: `Jira Attachment Saver`
   - Click **"Install"**
   - Review permissions
   - Click **"Continue"** → **"Allow"**

3. **Configure Add-on**
   - Open any email
   - Click add-on icon in right sidebar
   - Go to Settings (⚙️)
   - Enter Jira credentials:
     - Jira URL: `https://support.transporeon.com`
     - Jira Email: Your Jira login email
     - Jira API Token: [Generate from Jira](https://id.atlassian.com/manage-profile/security/api-tokens)
   - Click **"Save Settings"**
   - Click **"Test Jira Connection"** to verify

### Step 7.3: Share Documentation

1. **Create User Guide**
   - Share the `README.md` with users
   - Consider creating a wiki page or Google Doc

2. **Training Materials**
   - Optional: Create video tutorial
   - Host knowledge base articles

3. **Communication**
   - Send announcement email to organization
   - Include:
     - What the add-on does
     - How to install it
     - Link to user guide
     - Support contact

---

## Troubleshooting

### Common Issues During Deployment

#### Issue 1: "OAuth Consent Screen Error"
**Symptoms**: Cannot link GCP project to Apps Script

**Solution**:
1. Ensure OAuth Consent Screen is fully configured
2. Verify app type is "Internal"
3. Check that all required scopes are added
4. Ensure your account has proper permissions in GCP

#### Issue 2: "Script ID Not Found"
**Symptoms**: Cannot find Script ID when configuring Marketplace SDK

**Solution**:
1. In Apps Script editor, go to **Project Settings** (⚙️)
2. Look for **"Script ID"** (not Deployment ID)
3. Copy the entire ID string
4. If still not visible, ensure project is saved and linked to GCP

#### Issue 3: "API Not Enabled"
**Symptoms**: Error when running functions

**Solution**:
1. Go to GCP Console → APIs & Services → Dashboard
2. Enable missing APIs:
   - Gmail API
   - Google Drive API
   - Apps Script API
3. Wait a few minutes for propagation

#### Issue 4: "Permission Denied" Errors
**Symptoms**: Users cannot authorize the add-on

**Solution**:
1. Check OAuth scopes match exactly between:
   - `appsscript.json`
   - OAuth Consent Screen in GCP
   - Marketplace SDK configuration
2. Ensure users are in your Google Workspace domain
3. Check Google Workspace admin hasn't blocked the app

#### Issue 5: "Add-on Doesn't Appear in Gmail"
**Symptoms**: Add-on not showing in sidebar

**Solution**:
1. Verify deployment is active (not HEAD for production)
2. Check installation status in Google Workspace admin console
3. Refresh Gmail (Ctrl+R or Cmd+R)
4. Try in incognito mode
5. Verify user has installed the add-on
6. Check if user's Gmail is in correct mode (not Simplified view)

#### Issue 6: "Marketplace Listing Not Found"
**Symptoms**: Cannot find app in Workspace Marketplace

**Solution**:
1. Verify publishing status is "Published"
2. For internal apps, ensure user is signed in with domain account
3. Check visibility settings in Marketplace SDK
4. Wait up to 24 hours for full propagation

### Getting Help

**Google Workspace Admin Support**:
- For deployment and admin issues
- Contact: Google Workspace support portal

**Apps Script Documentation**:
- https://developers.google.com/apps-script/guides/gmail/add-ons

**GCP Support**:
- For GCP project and API issues
- Contact: GCP support console

**Marketplace Publishing Guide**:
- https://developers.google.com/workspace/marketplace/how-to-publish

---

## Checklist Summary

Use this checklist to track your deployment progress:

### GCP Setup
- [ ] Create new GCP project
- [ ] Note Project ID and Project Number
- [ ] Enable Gmail API
- [ ] Enable Google Drive API
- [ ] Enable Apps Script API
- [ ] Enable Google Workspace Marketplace SDK

### OAuth Configuration
- [ ] Configure OAuth Consent Screen (Internal)
- [ ] Add all 8 required scopes
- [ ] Upload app logo
- [ ] Add authorized domains

### Apps Script Deployment
- [ ] Create Apps Script project
- [ ] Link to GCP project (using Project Number)
- [ ] Upload Code.js
- [ ] Configure appsscript.json
- [ ] Run setupScriptProperties()
- [ ] Update Script Properties for your environment
- [ ] Create production deployment
- [ ] Note Deployment ID

### Testing
- [ ] Install test version
- [ ] Test in Gmail with real email
- [ ] Verify Jira connection
- [ ] Test attachment saving
- [ ] Check folder creation
- [ ] Review execution logs

### Marketplace Publishing
- [ ] Prepare screenshots (3-5 images)
- [ ] Prepare application description
- [ ] Configure Marketplace SDK
- [ ] Add Gmail Add-on extension
- [ ] Set visibility to "Private - My domain only"
- [ ] Add scope justifications
- [ ] Submit for publishing
- [ ] Verify published status

### Post-Deployment
- [ ] Install via Admin Console (if auto-install)
- [ ] OR share installation instructions with users
- [ ] Share user documentation
- [ ] Set up support process
- [ ] Monitor usage and errors

---

## Next Steps After Deployment

1. **Monitor Usage**
   - Check Apps Script Executions regularly
   - Review error logs
   - Gather user feedback

2. **Iterate and Improve**
   - Fix bugs reported by users
   - Add requested features
   - Update deployment with new versions

3. **Maintain**
   - Keep Jira API tokens updated
   - Monitor quota usage (Apps Script has daily quotas)
   - Update OAuth scopes if adding new features
   - Keep documentation current

4. **Scale**
   - If successful, consider expanding to other teams
   - Add more integrations (e.g., other ticketing systems)
   - Implement analytics for usage tracking

---

## Version History Template

Keep track of your deployments:

| Version | Date | Changes | Deployment ID |
|---------|------|---------|---------------|
| 1.0.0   | YYYY-MM-DD | Initial release | AKfycby... |
| 1.1.0   | YYYY-MM-DD | Bug fixes, added subfolder support | AKfycby... |

---

## Support Contacts

**Internal Team**:
- Developer: [Your name/email]
- Admin: [Admin name/email]

**External Resources**:
- Apps Script Community: https://support.google.com/code/community
- Stack Overflow: Tag with `google-apps-script` and `gmail-add-on`

---

**Document Version**: 1.0
**Last Updated**: 2025-12-12
**Author**: AI Assistant
**Status**: Ready for Use

