# Deployment Troubleshooting Guide

Comprehensive troubleshooting guide for Gmail Attachment Saver deployment issues.

---

## Table of Contents

1. [GCP Project Issues](#gcp-project-issues)
2. [OAuth Configuration Issues](#oauth-configuration-issues)
3. [Apps Script Deployment Issues](#apps-script-deployment-issues)
4. [Testing Issues](#testing-issues)
5. [Marketplace Publishing Issues](#marketplace-publishing-issues)
6. [User Installation Issues](#user-installation-issues)
7. [Runtime Issues](#runtime-issues)
8. [Performance Issues](#performance-issues)
9. [Getting Additional Help](#getting-additional-help)

---

## GCP Project Issues

### ❌ Issue: Cannot Create GCP Project

**Symptoms**:
- "You don't have permission to create projects"
- Create button is grayed out

**Causes**:
1. Insufficient permissions in your organization
2. Organization policy restrictions
3. Billing not enabled

**Solutions**:

✅ **Check Your Permissions**:
```
1. Contact your Google Workspace admin
2. Required roles:
   - Project Creator
   - OR Organization Admin
3. Ask admin to grant permissions or create project for you
```

✅ **Check Organization Policies**:
```
1. Admin needs to check: GCP Console → IAM & Admin → Organization Policies
2. Look for restrictions on project creation
3. May need to request exception
```

✅ **Enable Billing**:
```
1. GCP projects require billing account (even if free tier)
2. Go to: Billing → Create billing account
3. Link billing account to project
```

---

### ❌ Issue: Can't Find Project Number

**Symptoms**:
- Apps Script asking for Project Number
- Only see Project ID

**Solution**:

✅ **Locate Project Number**:
```
Method 1:
1. Go to: https://console.cloud.google.com/home/dashboard
2. Select your project from dropdown
3. Look at "Project Info" card on left
4. Project Number is shown below Project ID
   Example: Project Number: 123456789012

Method 2:
1. GCP Console → IAM & Admin → Settings
2. Project Number shown at top of page

Method 3:
1. Run this in Cloud Shell:
   gcloud projects describe YOUR-PROJECT-ID --format="value(projectNumber)"
```

**Important**: 
- Project ID: `gmail-attachment-saver-123` (alphanumeric with dashes)
- Project Number: `123456789012` (numeric only)
- Apps Script needs the **numeric Project Number**

---

### ❌ Issue: API Enablement Fails

**Symptoms**:
- "Failed to enable API"
- API shows as "Disabled" after enabling

**Solutions**:

✅ **Wait and Retry**:
```
1. API enablement can take 1-5 minutes
2. Refresh the page
3. Try enabling again
```

✅ **Check Quotas**:
```
1. Go to: APIs & Services → Quotas
2. Look for "API Enable" quota
3. May need to wait if quota exceeded
```

✅ **Use gcloud Command**:
```bash
# Enable all required APIs at once
gcloud services enable gmail.googleapis.com \
  drive.googleapis.com \
  script.googleapis.com \
  appsmarket-component.googleapis.com \
  --project=YOUR-PROJECT-ID
```

---

## OAuth Configuration Issues

### ❌ Issue: "OAuth Consent Screen Required"

**Symptoms**:
- Cannot link Apps Script to GCP project
- Error: "Configure OAuth consent screen first"

**Solution**:

✅ **Complete OAuth Setup First**:
```
MUST be done in this order:
1. ✓ Create GCP project
2. ✓ Enable APIs
3. ✓ Configure OAuth Consent Screen  ← You are here
4. Link Apps Script project
```

✅ **Minimum Required Fields**:
```
- App name: Required
- User support email: Required
- Developer contact email: Required
- At least one scope: Required
```

---

### ❌ Issue: Scope Not Found / Invalid Scope

**Symptoms**:
- "Scope is invalid"
- Cannot add specific scope to OAuth screen

**Solution**:

✅ **Use Exact Scope URLs**:
```
Copy these exactly (case-sensitive, no spaces):

✓ https://www.googleapis.com/auth/gmail.readonly
✓ https://www.googleapis.com/auth/drive
✓ https://www.googleapis.com/auth/script.locale
✓ https://www.googleapis.com/auth/gmail.addons.execute
✓ https://www.googleapis.com/auth/script.external_request
✓ https://www.googleapis.com/auth/gmail.addons.current.message.metadata
✓ https://www.googleapis.com/auth/gmail.addons.current.message.readonly
✓ https://www.googleapis.com/auth/userinfo.email
```

✅ **Add Scopes Manually**:
```
1. In OAuth Consent Screen → Scopes
2. Click "Add or Remove Scopes"
3. Scroll to bottom
4. Click "Manually add scopes"
5. Paste all scopes (one per line or comma-separated)
6. Click "Add to Table"
```

---

### ❌ Issue: "This app hasn't been verified"

**Symptoms**:
- Users see warning screen when authorizing
- "Google hasn't verified this app"

**Solution**:

✅ **For Internal Apps** (Recommended):
```
This is NORMAL for internal apps and can be safely bypassed:
1. User clicks "Advanced"
2. User clicks "Go to [App Name] (unsafe)"
3. This is expected behavior for domain-internal apps
```

✅ **For External Apps** (If needed):
```
Requires OAuth verification process:
1. Submit app for verification
2. Provide demo video
3. Provide test credentials
4. Wait 4-6 weeks for review
5. NOT recommended for internal tools
```

---

## Apps Script Deployment Issues

### ❌ Issue: Cannot Link Apps Script to GCP Project

**Symptoms**:
- "Invalid project ID or number"
- "Project not found"
- Link button not responding

**Solutions**:

✅ **Verify Project Number** (Not Project ID):
```
❌ Wrong: gmail-attachment-saver-123456 (Project ID)
✅ Correct: 123456789012 (Project Number)

To find: GCP Console → Home Dashboard → Project Info card
```

✅ **Check Apps Script API**:
```
1. GCP Console → APIs & Services → Dashboard
2. Verify "Apps Script API" is enabled
3. If not, enable it and wait 5 minutes
4. Try linking again
```

✅ **Check Permissions**:
```
Your account needs:
- Apps Script API User role
- OR Project Editor role
- OR Owner role

To check:
1. GCP Console → IAM & Admin → IAM
2. Find your email
3. Verify role includes Apps Script access
```

✅ **Clear and Retry**:
```
1. In Apps Script: Project Settings
2. If a project is already linked, click "Change project"
3. Enter the correct Project Number
4. Save and refresh the page
```

---

### ❌ Issue: "Script Properties Not Set"

**Symptoms**:
- Add-on runs but crashes
- Errors about undefined properties
- `getCustomersFolderId()` returns null

**Solution**:

✅ **Run Setup Function**:
```
1. Apps Script Editor
2. Select function: setupScriptProperties
3. Click Run ▶️
4. Check execution log for "✅ Script properties configured"
```

✅ **Manual Property Setup**:
```
If setupScriptProperties() fails:
1. Apps Script → Project Settings
2. Scroll to "Script Properties"
3. Click "Add script property"
4. Add these manually:

   Property Name              | Value
   ---------------------------|--------------------------------
   DEFAULT_JIRA_URL          | https://support.transporeon.com
   SETTINGS_STORAGE_KEY      | JIRA_SETTINGS
   CUSTOMERS_FOLDER_ID       | 1pi6fJzg5nncV7ADrcivI0gEAa9cjZLQp
   DEFAULT_LOG_LEVEL         | 2

5. Click "Save script properties"
```

✅ **Verify Properties**:
```javascript
// Run this function to test:
function testProperties() {
  var scriptProps = PropertiesService.getScriptProperties();
  var allProps = scriptProps.getProperties();
  
  Logger.log('All Script Properties:');
  for (var key in allProps) {
    Logger.log(key + ': ' + allProps[key]);
  }
}
```

---

### ❌ Issue: Deployment Creation Fails

**Symptoms**:
- "Failed to create deployment"
- Deployment button grayed out
- Error during deployment process

**Solutions**:

✅ **Check appsscript.json Syntax**:
```
Common JSON errors:
- Missing commas
- Trailing commas (not allowed in JSON)
- Mismatched brackets
- Wrong quotes (use double quotes ")

Validation:
1. Copy appsscript.json content
2. Paste into: https://jsonlint.com/
3. Fix any reported errors
```

✅ **Verify Trigger Function Exists**:
```javascript
// appsscript.json references this function:
"onTriggerFunction": "buildAddOn"

// Make sure this function exists in Code.js:
function buildAddOn(e) {
  // ... function code
}
```

✅ **Check for Code Errors**:
```
1. Apps Script Editor
2. Select any function
3. Click Run ▶️
4. Check for syntax errors
5. Fix all errors before deploying
```

✅ **Try Different Deployment Type**:
```
If Add-on deployment fails:
1. Try: Deploy → New deployment → Web app first
2. If that works, then try Add-on deployment
3. This helps identify if issue is with code or deployment type
```

---

### ❌ Issue: Advanced Services Not Working

**Symptoms**:
- "Gmail is not defined"
- "Drive is not defined"
- Advanced services errors

**Solutions**:

✅ **Enable Advanced Services in Apps Script**:
```
1. Apps Script Editor
2. Services (+ icon on left sidebar)
3. Add services:
   - Gmail API (v1)
   - Drive API (v3)
4. They should appear in Services list
```

✅ **Verify appsscript.json Configuration**:
```json
{
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
  }
}
```

✅ **Verify APIs Enabled in GCP**:
```
GCP Console → APIs & Services → Dashboard
Must show as enabled:
- Gmail API
- Google Drive API
```

---

## Testing Issues

### ❌ Issue: Add-on Not Appearing in Gmail

**Symptoms**:
- No add-on icon in Gmail sidebar
- Sidebar is empty
- Add-on installed but not visible

**Solutions**:

✅ **Check Gmail View**:
```
Gmail has multiple views:
1. Default view ✅ (add-ons work)
2. Simplified view ❌ (no add-ons)

To check:
1. Gmail Settings (⚙️) → See all settings
2. Top tabs → "Inbox"
3. Look for "Inbox type"
4. Make sure NOT using "Simplified" Gmail

To enable default view:
- Click ⚙️ → Try the new Gmail
- OR disable Gmail offline mode
```

✅ **Verify Installation**:
```
Method 1 - Check test installation:
1. Apps Script → Deploy → Test deployments
2. Click "Install"
3. Verify installation completed

Method 2 - Check in Gmail:
1. Gmail → Settings (⚙️) → See all settings
2. Add-ons tab
3. Look for "Jira Attachment Saver"
4. Should show as "Installed"

Method 3 - Reinstall:
1. Uninstall add-on
2. Clear browser cache
3. Reinstall test deployment
```

✅ **Hard Refresh Gmail**:
```
Windows/Linux: Ctrl + Shift + R
Mac: Cmd + Shift + R

Or:
1. Close all Gmail tabs
2. Clear browser cache for mail.google.com
3. Reopen Gmail
4. Wait 30 seconds for add-on to load
```

✅ **Check Browser Console**:
```
1. In Gmail, press F12 (Developer Tools)
2. Go to Console tab
3. Look for errors related to add-on
4. Common issues:
   - CORS errors → Check GCP project linking
   - 401/403 errors → Authorization issue
   - Timeout errors → Script taking too long
```

---

### ❌ Issue: "Authorization Required" Loop

**Symptoms**:
- Keeps asking for authorization
- Authorization doesn't persist
- Returns to auth screen repeatedly

**Solutions**:

✅ **Complete Authorization Properly**:
```
1. Click "Authorize access"
2. Select CORRECT Google account (domain account)
3. Review permissions carefully
4. Click "Allow" (NOT Cancel)
5. Wait for redirect back to Gmail
```

✅ **Clear Stored Authorizations**:
```
1. Go to: https://myaccount.google.com/permissions
2. Find "Jira Attachment Saver"
3. Click "Remove Access"
4. Go back to Gmail
5. Open add-on and re-authorize from scratch
```

✅ **Check for Scope Mismatches**:
```
If scopes changed after users authorized:
1. Users must re-authorize
2. Remove old authorization (steps above)
3. Fresh authorization with new scopes
```

✅ **Incognito Test**:
```
1. Open Gmail in incognito/private window
2. Sign in
3. Try to authorize add-on
4. If works in incognito:
   - Clear cookies/cache in normal browser
   - Try again
```

---

### ❌ Issue: Execution Timeouts

**Symptoms**:
- "Execution time limit exceeded"
- Script stops running
- Only partial results

**Solutions**:

✅ **Optimize Long-Running Operations**:
```javascript
// Bad: Processing too many items at once
function processAllAttachments(attachments) {
  attachments.forEach(att => saveAttachment(att));
}

// Good: Batch processing with timeouts
function processAttachmentsBatch(attachments, maxTime) {
  var startTime = new Date().getTime();
  var processedCount = 0;
  
  for (var i = 0; i < attachments.length; i++) {
    if (new Date().getTime() - startTime > maxTime) {
      Logger.log('Timeout approaching, processed ' + processedCount);
      break;
    }
    saveAttachment(attachments[i]);
    processedCount++;
  }
}
```

✅ **Use Time-Based Triggers for Large Operations**:
```javascript
// Instead of doing everything at once,
// use triggers for long operations
function setupBatchProcessing() {
  ScriptApp.newTrigger('continueProcessing')
    .timeBased()
    .after(1000) // 1 second
    .create();
}
```

✅ **Check Execution Logs**:
```
1. Apps Script → Executions (⏱️)
2. Click on failed execution
3. Note where it stopped
4. Optimize that section
```

**Gmail Add-on Limits**:
- Max execution time: 30 seconds per action
- Solution: Process fewer items per action

---

## Marketplace Publishing Issues

### ❌ Issue: Cannot Find Script ID

**Symptoms**:
- Marketplace SDK asking for Script ID
- Can't locate Script ID in Apps Script

**Solution**:

✅ **Locate Script ID**:
```
1. Apps Script Editor
2. Click: Project Settings (⚙️) in left sidebar
3. Scroll down to "IDs" section
4. Copy "Script ID"

Example Script ID format:
1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7q8r9s0t1u2v3w4x5y6z

Note: This is different from:
- Project Number (for linking GCP)
- Deployment ID (for deployments)
```

---

### ❌ Issue: "Extension Verification Failed"

**Symptoms**:
- Red X next to extension in Marketplace SDK
- "Could not verify extension"
- Extension not loading

**Solutions**:

✅ **Verify Deployment is Active**:
```
1. Apps Script → Deploy → Manage deployments
2. Check deployment status: "Active"
3. If not active, create new deployment
4. Use that Deployment ID in Marketplace SDK
```

✅ **Use Correct IDs**:
```
Marketplace SDK requires TWO IDs:

1. Script ID (NOT Project Number):
   - Get from: Apps Script → Project Settings
   - Example: 1a2b3c4d5e...

2. Deployment ID:
   - Get from: Apps Script → Deploy → Manage deployments
   - Example: AKfycby...
   
Common mistake: Using Project Number instead of Script ID
```

✅ **Wait for Propagation**:
```
After creating deployment:
1. Wait 5-10 minutes
2. Refresh Marketplace SDK page
3. Re-enter IDs and verify
```

✅ **Check Project Linking**:
```
1. Apps Script → Project Settings
2. "Google Cloud Platform (GCP) Project" section
3. Should show your GCP Project Number
4. If not linked, link it first
5. Then try Marketplace SDK again
```

---

### ❌ Issue: Publishing Fails / Stuck in Review

**Symptoms**:
- "Unable to publish"
- Status stuck at "In Review"
- No approval after long wait

**Solutions**:

✅ **For Internal Apps** (Should be instant):
```
If stuck in review:
1. Check visibility setting: MUST be "Private → My domain only"
2. Check all required fields filled:
   - App name ✓
   - Description ✓
   - Icon ✓
   - At least 1 screenshot ✓
   - Developer contact ✓
3. Save configuration
4. Try publishing again
```

✅ **Check for Validation Errors**:
```
1. Marketplace SDK → App Configuration
2. Look for red error messages
3. Common issues:
   - Missing required fields
   - Invalid image dimensions
   - Scope justifications missing
   - Description too short
4. Fix all errors
5. Save and publish again
```

✅ **Contact Support**:
```
If internal app not approved after 24 hours:
1. Check GCP Console notifications
2. Contact Google Workspace support
3. Provide:
   - GCP Project ID
   - App name
   - Screenshot of publishing status
```

---

## User Installation Issues

### ❌ Issue: Users Cannot Find App in Marketplace

**Symptoms**:
- Search returns no results
- App not visible to domain users
- "No apps found"

**Solutions**:

✅ **Verify Publishing Status**:
```
1. GCP Console → Marketplace SDK
2. Check "Publishing Status": Must say "Published"
3. If not published, complete publishing process
```

✅ **Verify Visibility Settings**:
```
1. Marketplace SDK → App Configuration → Distribution
2. Visibility: "Private → My domain only"
3. Authorized domains: your-domain.com (should match user's email)
4. Save if changed, wait 30 minutes for propagation
```

✅ **Direct Installation Link**:
```
Workaround while troubleshooting:
1. Apps Script → Deploy → Manage deployments
2. Click deployment → Get installation URL
3. Share URL directly with users
4. Users click link to install
```

✅ **Admin Installation**:
```
Admin can push app to users:
1. https://admin.google.com/
2. Apps → Google Workspace Marketplace apps
3. Add app → Search for your app
4. Domain Install → Select OUs/users
5. Install
```

---

### ❌ Issue: "App Not Available for Your Account"

**Symptoms**:
- User sees error when trying to install
- "This app is not available"
- Installation blocked

**Solutions**:

✅ **Check User Account Type**:
```
For internal apps:
- User MUST use Google Workspace account (@your-domain.com)
- Personal Gmail accounts (@gmail.com) CANNOT install
- User must be in authorized domain

Verification:
- User checks: What email am I signed in with?
- Must match authorized domain in Marketplace config
```

✅ **Check Admin Restrictions**:
```
Admin may have blocked marketplace installations:
1. Admin Console → Apps → Google Workspace Marketplace
2. Settings
3. Check: "Allow users to install apps"
4. If disabled, admin must:
   - Either enable user installations
   - OR admin install app for users
```

✅ **Check Organizational Unit**:
```
If admin installed for specific OUs:
1. User must be in included OU
2. Admin can check: Admin Console → Users → [User] → User information
3. See which OU user belongs to
4. Admin may need to expand installation to more OUs
```

---

## Runtime Issues

### ❌ Issue: "Failed to Connect to Jira"

**Symptoms**:
- Test Jira Connection fails
- Cannot fetch tickets
- 401 or 403 errors

**Solutions**:

✅ **Verify Jira Credentials**:
```
1. Jira URL format:
   ✅ https://your-jira.atlassian.com
   ✅ https://jira.your-company.com
   ❌ https://your-jira.atlassian.com/ (no trailing slash)
   ❌ your-jira.atlassian.com (must include https://)

2. Jira Email:
   - Must be email you use to login to Jira
   - Usually your company email

3. Jira API Token:
   - Generate new: https://id.atlassian.com/manage-profile/security/api-tokens
   - Click "Create API token"
   - Copy ENTIRE token (they're long!)
   - Paste into add-on settings
```

✅ **Test Jira Access Manually**:
```bash
# Test if Jira API is accessible
curl -u your-email@company.com:YOUR_API_TOKEN \
  https://your-jira.atlassian.com/rest/api/2/myself

Expected: JSON with your user info
Error: Check credentials or network access
```

✅ **Check Jira Permissions**:
```
Your Jira account needs:
- Access to projects in JQL query
- Permission to view issues
- API access enabled (some orgs disable this)

Test:
1. Log into Jira web interface
2. Try running the JQL query manually:
   - Jira → Filters → Advanced search
   - Paste JQL from add-on settings
   - Should return results
   - If no results, adjust JQL
```

✅ **Check External URL Access**:
```
Apps Script might be blocked from external URLs:
1. GCP Console → VPC Network → Firewall rules (rare)
2. More likely: Jira blocking Google's IPs
   - Contact Jira admin
   - Whitelist Google Apps Script IPs
```

---

### ❌ Issue: "Folder Not Found" or Permission Denied

**Symptoms**:
- Cannot save attachments
- "You need permission to access this folder"
- Folder ID errors

**Solutions**:

✅ **Verify Customers Folder ID**:
```
1. Open Google Drive
2. Navigate to your "Customers" folder
3. URL will look like:
   https://drive.google.com/drive/folders/1pi6fJzg5nncV7ADrcivI0gEAa9cjZLQp
                                            ^^^^^^^^^^^^^^^^^^^^^^^^^
                                            This is the Folder ID

4. Copy Folder ID
5. Update Script Property:
   Apps Script → Project Settings → Script Properties
   CUSTOMERS_FOLDER_ID = [paste folder ID]
```

✅ **Check Folder Permissions**:
```
Users need access to Customers folder:

Method 1 - Individual sharing:
1. Right-click Customers folder in Drive
2. Share
3. Add all users who will use add-on
4. Give "Editor" or "Content Manager" permission

Method 2 - Shared Drive (Recommended):
1. Create/use Shared Drive for customers
2. Add all users to Shared Drive
3. Use Shared Drive folder ID in settings
4. Benefits:
   - Automatic access for all members
   - Easier management
   - Better for teams
```

✅ **Test Folder Access**:
```javascript
// Run this in Apps Script to test:
function testFolderAccess() {
  try {
    var folderId = getCustomersFolderId();
    Logger.log('Folder ID: ' + folderId);
    
    var folder = DriveApp.getFolderById(folderId);
    Logger.log('Folder name: ' + folder.getName());
    Logger.log('Folder owner: ' + folder.getOwner().getEmail());
    Logger.log('✅ Access confirmed');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
  }
}
```

---

### ❌ Issue: Attachments Not Saving / Silent Failures

**Symptoms**:
- Success message appears but files not in Drive
- No error but no files either
- Inconsistent saves

**Solutions**:

✅ **Check Execution Logs**:
```
1. Apps Script → Executions (⏱️)
2. Find recent execution
3. Click to expand
4. Look for actual errors (even if user saw "success")
5. Common hidden errors:
   - Quota exceeded
   - Timeout after file download
   - Drive API errors
```

✅ **Check Drive Quotas**:
```
1. Check user's Google Drive storage
2. Check organization's Drive quotas
3. Check Apps Script quotas:
   - Go to: Apps Script → Project Settings → Quotas
   - Look for "Drive API calls" and "URL Fetch calls"
   - If near limit, wait or optimize code
```

✅ **Test with Small Attachment**:
```
1. Send yourself test email with small file (<1MB)
2. Try to save just that one file
3. If succeeds: Issue might be with large files
4. If fails: Issue is with save logic

For large files:
- Gmail Add-ons have memory limits
- Files >25MB may fail
- Consider warning users about file size limits
```

✅ **Add Better Error Handling**:
```javascript
// In Code.js, find saveAttachment functions
// Add try-catch with detailed logging:

function saveAttachment(attachment, folder) {
  try {
    Logger.log('Saving: ' + attachment.name + ' (' + attachment.size + ' bytes)');
    
    var blob = attachment.copyBlob();
    Logger.log('Blob created successfully');
    
    var file = folder.createFile(blob);
    Logger.log('File created: ' + file.getId());
    
    return file;
    
  } catch (error) {
    Logger.log('❌ Error saving ' + attachment.name + ': ' + error.message);
    throw error; // Re-throw to show user
  }
}
```

---

## Performance Issues

### ❌ Issue: Slow Loading / Timeouts

**Symptoms**:
- Add-on takes very long to load
- Jira tickets take forever to fetch
- Timeouts during operations

**Solutions**:

✅ **Optimize Jira Queries**:
```javascript
// Limit number of results
var jql = "project = CXPRO... ORDER BY updated DESC";
var maxResults = 50; // Don't fetch 1000s of tickets

// Add to Jira API call:
var url = jiraUrl + '/rest/api/2/search?jql=' + 
          encodeURIComponent(jql) + 
          '&maxResults=' + maxResults +
          '&fields=key,summary,status'; // Only fields you need
```

✅ **Implement Caching**:
```javascript
// Cache Jira results for short time
function getJiraTicketsWithCache() {
  var cache = CacheService.getUserCache();
  var cacheKey = 'jira_tickets';
  var cached = cache.get(cacheKey);
  
  if (cached) {
    Logger.log('Using cached tickets');
    return JSON.parse(cached);
  }
  
  // Fetch fresh data
  var tickets = fetchJiraTickets();
  
  // Cache for 5 minutes (300 seconds)
  cache.put(cacheKey, JSON.stringify(tickets), 300);
  
  return tickets;
}
```

✅ **Optimize Drive API Calls**:
```javascript
// Use Drive.Files.list (Advanced Drive API) instead of DriveApp for searches
// It's faster for large folder structures

function findFolderOptimized(parentId, folderName) {
  try {
    var query = "'" + parentId + "' in parents and " +
                "mimeType = 'application/vnd.google-apps.folder' and " +
                "name = '" + folderName.replace(/'/g, "\\'") + "' and " +
                "trashed = false";
    
    var results = Drive.Files.list({
      q: query,
      fields: 'files(id, name)',
      pageSize: 1
    });
    
    if (results.files && results.files.length > 0) {
      return results.files[0].id;
    }
    return null;
    
  } catch (error) {
    Logger.log('Error in optimized search: ' + error.message);
    return null;
  }
}
```

---

## Getting Additional Help

### Documentation Resources

📚 **Official Google Documentation**:
- Apps Script: https://developers.google.com/apps-script
- Gmail Add-ons: https://developers.google.com/apps-script/add-ons/gmail
- Drive API: https://developers.google.com/drive/api/guides/about-sdk
- Workspace Marketplace: https://developers.google.com/workspace/marketplace

🔧 **Tools**:
- Apps Script Community: https://support.google.com/code/community
- Stack Overflow: Tag [google-apps-script]
- Issue Tracker: https://issuetracker.google.com/issues?q=componentid:191640

### Support Channels

**For GCP/API Issues**:
1. GCP Console → Support
2. Create support ticket with:
   - Project ID
   - Error messages
   - Steps to reproduce

**For Workspace Admin Issues**:
1. https://admin.google.com/
2. Help → Contact support
3. Choose: "Google Workspace Marketplace"

**For Development Questions**:
1. Stack Overflow: https://stackoverflow.com/questions/tagged/google-apps-script
2. Tag your question with: `google-apps-script`, `gmail-addon`
3. Include:
   - Relevant code snippet
   - Error message
   - What you've tried
   - Expected vs actual behavior

### Debug Mode Execution

Enable detailed logging for troubleshooting:

```javascript
// Add to start of Code.js
var DEBUG_MODE = true; // Set to false for production

function debugLog(message) {
  if (DEBUG_MODE) {
    Logger.log('[DEBUG] ' + new Date().toISOString() + ' - ' + message);
  }
}

// Use throughout code:
function someFunction() {
  debugLog('Function started');
  // ... code ...
  debugLog('Operation completed');
}
```

### Collecting Debug Information

When reporting issues, provide:

```
1. Environment:
   - GCP Project ID: ___________
   - Apps Script ID: ___________
   - Deployment ID: ___________

2. Error Details:
   - Exact error message
   - Screenshot if applicable
   - When does it occur? (during what action)

3. Execution Log:
   - Apps Script → Executions
   - Find failed execution
   - Copy full log
   - Remove any sensitive data (API tokens, emails)

4. Steps to Reproduce:
   - List exact steps that cause the issue
   - Include any specific data (ticket numbers, etc)

5. Expected vs Actual:
   - What should happen?
   - What actually happens?
```

---

## Emergency Rollback Procedure

If deployment causes major issues:

```
1. Revert to Previous Deployment:
   - Apps Script → Deploy → Manage deployments
   - Find previous working deployment
   - Click ⋮ → Make active deployment
   - Users will automatically use old version

2. Disable Add-on Temporarily:
   - Admin Console → Google Workspace Marketplace apps
   - Find app → Disable
   - Fix issues in Apps Script
   - Re-enable when ready

3. Notify Users:
   - Send email about temporary issue
   - Provide ETA for fix
   - Suggest workarounds if available
```

---

## Preventive Measures

To avoid common issues:

```
✓ Version Control:
  - Keep backups of working versions
  - Document changes between versions
  - Test thoroughly before deploying

✓ Staged Rollout:
  - Test with small user group first
  - Monitor for 24-48 hours
  - Gradually expand to more users

✓ Monitoring:
  - Check execution logs daily (first week)
  - Set up alerts for high error rates
  - Maintain support channel for quick feedback

✓ Documentation:
  - Keep user docs updated
  - Document known issues
  - Maintain troubleshooting wiki
```

---

**Troubleshooting Guide Version**: 1.0  
**Last Updated**: 2025-12-12  
**Related Docs**: 
- [Deployment Guide](./deployment-guide.md)
- [Quick Checklist](./deployment-quick-checklist.md)
- [Workflow Diagram](./deployment-workflow-diagram.md)






