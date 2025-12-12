# Quick Deployment Checklist

A condensed checklist for deploying Gmail Attachment Saver to GCP and Google Workspace Marketplace.

> 📖 For detailed instructions, see: [deployment-guide.md](./deployment-guide.md)

---

## Pre-Flight Check ✈️

Before you begin:

- [ ] Google Workspace Admin access
- [ ] GCP project creation rights
- [ ] Jira instance URL and API credentials ready
- [ ] Google Drive "Customers Folder" ID ready
- [ ] All project files: `Code.js`, `appsscript.json`, logo image

---

## Phase 1: GCP Project (15 min) 🏗️

- [ ] Create new GCP project at https://console.cloud.google.com/
- [ ] **Note Project ID**: ____________________
- [ ] **Note Project Number**: ____________________
- [ ] Enable APIs:
  - [ ] Gmail API
  - [ ] Google Drive API  
  - [ ] Apps Script API
  - [ ] Google Workspace Marketplace SDK

---

## Phase 2: OAuth Consent Screen (10 min) 🔐

- [ ] Navigate to: APIs & Services → OAuth consent screen
- [ ] Select: **Internal** user type
- [ ] Fill app info:
  - [ ] App name: "Jira Attachment Saver"
  - [ ] User support email
  - [ ] Upload logo (128x128px)
  - [ ] Developer contact email
- [ ] Add all 8 scopes:
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
- [ ] Save configuration

---

## Phase 3: Apps Script Deployment (20 min) 📝

- [ ] Go to https://script.google.com/
- [ ] Create new project: "Jira Attachment Saver"
- [ ] Link to GCP:
  - [ ] Project Settings → Change project
  - [ ] Enter GCP **Project Number** (not ID!)
  - [ ] Click "Set Project"
- [ ] Upload files:
  - [ ] Copy-paste `Code.js` content
  - [ ] Enable & configure `appsscript.json` manifest
- [ ] Configure Script Properties:
  - [ ] Run function: `setupScriptProperties()`
  - [ ] Authorize the script
  - [ ] Verify properties created
  - [ ] Update `CUSTOMERS_FOLDER_ID` with your folder ID
  - [ ] Update `DEFAULT_JIRA_URL` if needed
- [ ] Create deployment:
  - [ ] Deploy → New deployment → Add-on
  - [ ] **Note Deployment ID**: ____________________
  - [ ] **Note Script ID**: ____________________

---

## Phase 4: Testing (15 min) 🧪

- [ ] Install test version: Deploy → Test deployments → Install
- [ ] Open Gmail: https://mail.google.com/
- [ ] Find add-on in right sidebar
- [ ] Test flow:
  - [ ] Add-on loads without errors
  - [ ] Configure Jira settings
  - [ ] Test Jira connection ✅
  - [ ] Fetch tickets successfully
  - [ ] Select attachments
  - [ ] Save to Drive successfully
  - [ ] Verify folder structure created
- [ ] Check execution logs: Apps Script → Executions

---

## Phase 5: Prepare Assets (30 min) 🎨

- [ ] Create 3-5 screenshots (1280x800px):
  - [ ] Add-on in Gmail sidebar
  - [ ] Jira ticket selection
  - [ ] Settings page
  - [ ] Success message
  - [ ] Drive folder structure
- [ ] Prepare descriptions (copy from deployment guide)
- [ ] Review documentation

---

## Phase 6: Marketplace Publishing (20 min) 🚀

- [ ] Go to: GCP Console → Google Workspace Marketplace SDK
- [ ] Click: App Configuration → Create new listing
- [ ] **App Basics**:
  - [ ] Name: "Jira Attachment Saver"
  - [ ] Description (from guide)
  - [ ] Icon (128x128px)
  - [ ] Category: Productivity
- [ ] **Extensions**:
  - [ ] Add Extension → Gmail Add-on
  - [ ] Enter Script ID
  - [ ] Enter Deployment ID
  - [ ] Verify extension ✅
- [ ] **Store Listing**:
  - [ ] Upload screenshots with captions
  - [ ] Set language
- [ ] **Distribution**:
  - [ ] Visibility: **Private → My domain only**
  - [ ] Installation: "Available to Install" or "Install for Everyone"
  - [ ] Authorized domain: your-domain.com
- [ ] **OAuth Scopes**:
  - [ ] Verify all 8 scopes detected
  - [ ] Add justification for each scope
- [ ] **Support & Contact**:
  - [ ] Developer name
  - [ ] Developer email
  - [ ] Support URL (optional)
  - [ ] Privacy policy URL (recommended)
- [ ] Save configuration
- [ ] Click **Publish** / **Submit**
- [ ] Wait for approval (usually instant for internal apps)

---

## Phase 7: Admin Installation (10 min) 👥

### Option A: Auto-Install for All Users

- [ ] Go to: https://admin.google.com/
- [ ] Navigate to: Apps → Google Workspace Marketplace apps
- [ ] Add app → Add from Google Workspace Marketplace
- [ ] Search: "Jira Attachment Saver"
- [ ] Install for: Everyone (or specific OUs)
- [ ] Finish installation

### Option B: Available for Users to Install

- [ ] Create user installation guide
- [ ] Share installation link
- [ ] Communicate to organization

---

## Phase 8: User Onboarding (5 min per user) 👤

Provide users with these steps:

- [ ] Open Gmail
- [ ] Find add-on in sidebar OR install from Marketplace
- [ ] Click add-on icon
- [ ] Go to Settings (⚙️)
- [ ] Configure:
  - [ ] Jira URL: `https://support.transporeon.com`
  - [ ] Jira Email: user's Jira email
  - [ ] Jira API Token: Generate at https://id.atlassian.com/manage-profile/security/api-tokens
- [ ] Save Settings
- [ ] Test Jira Connection ✅
- [ ] Test with real email

---

## Post-Deployment Monitoring 📊

### Week 1:
- [ ] Monitor execution logs daily
- [ ] Respond to user questions
- [ ] Document common issues

### Ongoing:
- [ ] Review weekly execution stats
- [ ] Gather user feedback
- [ ] Plan feature updates
- [ ] Keep documentation updated

---

## Troubleshooting Quick Fixes 🔧

| Problem | Quick Fix |
|---------|-----------|
| "OAuth error" | Verify all scopes match in appsscript.json, OAuth screen, and Marketplace |
| "Script ID not found" | Use Project Settings → Script ID (not Deployment ID) |
| "API not enabled" | Enable APIs in GCP Console, wait 5 minutes |
| "Add-on not visible" | Refresh Gmail, check installation status, verify domain |
| "Permission denied" | Re-authorize script, check user is in correct domain |

---

## Key URLs 🔗

- **GCP Console**: https://console.cloud.google.com/
- **Apps Script Editor**: https://script.google.com/
- **Gmail**: https://mail.google.com/
- **Admin Console**: https://admin.google.com/
- **Jira API Tokens**: https://id.atlassian.com/manage-profile/security/api-tokens

---

## Important IDs to Track 📝

| Item | Value | Where to Find |
|------|-------|---------------|
| GCP Project ID | _____________ | GCP Console Dashboard |
| GCP Project Number | _____________ | GCP Console Dashboard |
| Apps Script ID | _____________ | Apps Script → Project Settings |
| Deployment ID | _____________ | Apps Script → Deploy → Manage |
| Customers Folder ID | _____________ | Google Drive folder URL |

---

## Estimated Total Time ⏱️

- **First-time deployment**: ~2-3 hours
- **Subsequent updates**: ~30 minutes
- **User onboarding**: ~5 minutes per user

---

## Success Criteria ✅

Your deployment is successful when:

- [x] Add-on appears in Gmail sidebar
- [x] Users can configure Jira credentials
- [x] Jira connection test passes
- [x] Attachments save to correct Drive folders
- [x] Folder structure auto-creates properly
- [x] No errors in execution logs
- [x] Users report successful workflows

---

## Next Version Deployment 🔄

When you need to update the app:

1. [ ] Make changes in Apps Script editor
2. [ ] Test with HEAD deployment
3. [ ] Create new deployment: Deploy → New deployment
4. [ ] Update Marketplace SDK with new Deployment ID (optional)
5. [ ] Users automatically get update (no reinstall needed)

---

**Quick Reference Version**: 1.0  
**Last Updated**: 2025-12-12  
**Full Guide**: [deployment-guide.md](./deployment-guide.md)

