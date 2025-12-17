# How to Remove Existing Add-on Deployment

Guide for removing the Gmail Attachment Saver add-on from your personal Gmail account and cleaning up the previous deployment.

---

## Quick Removal Steps

### Step 1: Uninstall Add-on from Gmail (2 min)

**Method A: From Gmail Settings**
1. Open Gmail: https://mail.google.com/
2. Click **Settings (⚙️)** → **See all settings**
3. Go to **"Add-ons"** tab
4. Find **"Jira Attachment Saver"** in the list
5. Click **"Manage"** or **"Remove"**
6. Confirm removal
7. Refresh Gmail (Ctrl+Shift+R or Cmd+Shift+R)

**Method B: From Google Account Settings**
1. Go to: https://myaccount.google.com/permissions
2. Look for **"Jira Attachment Saver"** or your app name
3. Click on it
4. Click **"Remove Access"**
5. Confirm removal
6. Refresh Gmail

✅ **Verify**: The add-on icon should no longer appear in Gmail sidebar

---

### Step 2: Remove Test Deployment from Apps Script (3 min)

1. **Open Apps Script Project**
   - Go to: https://script.google.com/
   - Find your "Jira Attachment Saver" project
   - Click to open it

2. **Remove Test Deployments**
   - Click **Deploy** → **Test deployments**
   - If you see "Installed", click **"Uninstall"**

3. **Remove Production Deployments** (Optional)
   - Click **Deploy** → **Manage deployments**
   - For each deployment:
     - Click **⋮** (three dots)
     - Click **"Archive"**
   - This deactivates deployments without deleting code

4. **Unlink from GCP Project** (Optional)
   - Click **Project Settings (⚙️)**
   - Scroll to "Google Cloud Platform (GCP) Project"
   - Click **"Remove project"** (if you want to unlink it)

✅ **Verify**: Deployments should show as "Archived" or removed

---

### Step 3: Clean Up User Properties (1 min)

This removes your stored Jira credentials and settings:

1. **In Apps Script Editor**
   - Select this function from the dropdown:

```javascript
function cleanupUserProperties() {
  var userProps = PropertiesService.getUserProperties();
  userProps.deleteAllProperties();
  console.log('✅ User properties deleted');
}
```

2. **Run the Function**
   - Click **Run ▶️**
   - Check execution log: Should show "✅ User properties deleted"

**Alternatively, manual cleanup**:
   - Apps Script → **Project Settings (⚙️)**
   - Scroll to **"User properties"** (if any)
   - Delete all user properties

✅ **Verify**: No user properties should remain

---

### Step 4: Remove Authorization (Important!)

Remove the OAuth authorization you granted:

1. **Go to Google Account Permissions**
   - Visit: https://myaccount.google.com/permissions
   
2. **Find Your App**
   - Look for "Jira Attachment Saver" or "Gmail Attachment Saver"
   - It might also show as "Untitled project" if you didn't rename it

3. **Remove Access**
   - Click on the app
   - Click **"Remove Access"**
   - Confirm the removal

4. **Clear Third-party Apps with Account Access**
   - Same page: https://myaccount.google.com/permissions
   - Review all listed apps
   - Remove any related to your test deployment

✅ **Verify**: App should no longer appear in your authorized apps list

---

### Step 5: Optional - Delete or Clean GCP Project

You have two options:

**Option A: Keep Project, Remove Configuration** (Recommended if you'll reuse it)
1. Go to: https://console.cloud.google.com/
2. Select your test project
3. Go to: **APIs & Services → Credentials**
4. Delete any OAuth 2.0 Client IDs created for testing
5. Keep the project for future use

**Option B: Delete Entire Project** (Clean slate)
1. Go to: https://console.cloud.google.com/
2. Select your test project from dropdown
3. Go to: **IAM & Admin → Settings**
4. Click **"Shut Down"** button
5. Enter Project ID to confirm
6. Click **"Shut Down"** to delete

⚠️ **Warning**: Deleted projects can be recovered for 30 days, then permanently deleted.

✅ **Verify**: Project should show as "Pending deletion" or removed from your project list

---

## Complete Removal Verification Checklist

After completing the above steps, verify:

- [ ] Add-on icon no longer appears in Gmail sidebar
- [ ] App not listed in https://myaccount.google.com/permissions
- [ ] Test deployment uninstalled in Apps Script
- [ ] Production deployments archived in Apps Script
- [ ] No stored credentials (user properties cleaned)
- [ ] GCP project removed or cleaned up

---

## What About the Code Files?

**Don't delete your Apps Script project!** You'll need it for the new deployment.

### If Using Same Apps Script Project for New Deployment:

**Keep**:
- ✅ Code.js file
- ✅ appsscript.json file
- ✅ All your code

**Only Remove**:
- Old deployments (archive them)
- Old GCP project link
- User properties with old credentials

### If Starting Fresh with New Apps Script Project:

1. **Backup your code first**:
   - Download Code.js: File → Download → Code.js
   - Download appsscript.json: Download manifest file

2. **Optional: Delete old Apps Script project**:
   - Apps Script → Home (https://script.google.com/home)
   - Find project → Click ⋮ → Remove

---

## Preparing for New Deployment

After cleanup, you're ready for a clean deployment:

### What You'll Do Differently:

1. **Use Google Workspace Account** (not personal Gmail)
   - Sign out of personal Gmail
   - Sign in with your company account (e.g., @transporeon.com)

2. **Create New GCP Project** (for company)
   - Follow: [deployment-guide.md](./deployment-guide.md)
   - Use company's GCP organization

3. **Configure for Domain**
   - OAuth Consent Screen: Select "Internal"
   - Authorized domain: Your company domain

4. **Create Fresh Apps Script Project** (or reuse)
   - Link to NEW GCP project (company project)
   - Deploy to company's Google Workspace

---

## Troubleshooting Removal

### Issue: "Can't remove add-on from Gmail"

**Solution**:
```
1. Clear browser cache and cookies for mail.google.com
2. Try in incognito/private window
3. Remove via Google Account settings instead:
   https://myaccount.google.com/permissions
```

### Issue: "Add-on still appears after removal"

**Solution**:
```
1. Hard refresh Gmail: Ctrl+Shift+R (Win) or Cmd+Shift+R (Mac)
2. Clear browser cache completely
3. Close all Gmail tabs and reopen
4. Wait 5-10 minutes for propagation
5. Try different browser
```

### Issue: "Can't delete GCP project"

**Solution**:
```
1. Check if you're the project owner
2. Check for billing account (may need to unlink first)
3. Remove all resources first (APIs, services)
4. Try again after 5 minutes
5. If still fails, use "Shut Down" instead of "Delete"
```

### Issue: "OAuth authorization won't remove"

**Solution**:
```
1. Go to: https://myaccount.google.com/permissions
2. Find ALL instances of the app (might be listed multiple times)
3. Remove each one
4. Also check: https://security.google.com/settings/security/permissions
5. Force sign out and sign back in
```

---

## Quick Command Summary

```
✓ Uninstall from Gmail: Settings → Add-ons → Remove
✓ Remove authorization: https://myaccount.google.com/permissions
✓ Archive deployments: Apps Script → Deploy → Manage → Archive
✓ Clean user props: Run cleanupUserProperties() function
✓ Delete GCP project: GCP Console → Settings → Shut Down
```

---

## After Removal - Fresh Start

Once you've completed the removal:

1. **Sign in with correct account**
   - Your Google Workspace account (@company-domain.com)
   - NOT personal Gmail (@gmail.com)

2. **Follow deployment guide from beginning**
   - Start at: [deployment-guide.md](./deployment-guide.md)
   - This time: Use company GCP project
   - Configure as Internal app for your domain

3. **Key differences for new deployment**:
   - GCP project under company organization
   - OAuth consent screen: Internal (not external)
   - Authorized domain: Your company domain
   - Deploy to entire domain or specific OUs

---

## Need Help?

If you encounter issues during removal:

1. **Check execution logs**: Apps Script → Executions
2. **Try incognito mode**: Rules out browser cache issues
3. **Wait and retry**: Some changes take 5-10 minutes to propagate
4. **Contact support**: Google Workspace admin can force-remove apps

---

**Quick Removal Guide Version**: 1.0  
**Last Updated**: 2025-12-12  
**Next Step**: [deployment-guide.md](./deployment-guide.md) for fresh deployment

---

## Summary: 5-Minute Complete Removal

```
1. Gmail Settings → Add-ons → Remove
   ↓
2. https://myaccount.google.com/permissions → Remove Access
   ↓
3. Apps Script → Deploy → Archive all deployments
   ↓
4. Apps Script → Run: cleanupUserProperties()
   ↓
5. GCP Console → (Optional) Delete test project
   ↓
✅ DONE - Ready for fresh deployment
```



