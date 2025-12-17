# Deployment Workflow Diagram

Visual representation of the deployment process for Gmail Attachment Saver.

---

## Overview: 7 Major Phases

```
┌─────────────────────────────────────────────────────────────────────┐
│                     DEPLOYMENT WORKFLOW                              │
│                                                                      │
│  Phase 1          Phase 2         Phase 3         Phase 4          │
│  ┌──────┐       ┌──────┐        ┌──────┐        ┌──────┐          │
│  │ GCP  │  →    │OAuth │   →    │Apps  │   →    │Test  │          │
│  │Setup │       │Config│        │Script│        │      │          │
│  └──────┘       └──────┘        └──────┘        └──────┘          │
│                                                                      │
│  Phase 5          Phase 6         Phase 7                          │
│  ┌──────┐       ┌──────┐        ┌──────┐                          │
│  │Assets│  →    │Market│   →    │Deploy│                          │
│  │Prep  │       │place │        │Users │                          │
│  └──────┘       └──────┘        └──────┘                          │
│                                                                      │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Detailed Phase Breakdown

### Phase 1: GCP Project Setup (15 min)

```
┌────────────────────────────────────────────────────────┐
│                   GCP CONSOLE                          │
│  https://console.cloud.google.com                      │
└────────────────────────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  1. Create New Project        │
        │     - Name: Gmail Attachment  │
        │     - Org: Your Organization  │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  2. Note Important IDs        │
        │     - Project ID              │
        │     - Project Number ⚠️       │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  3. Enable APIs               │
        │     ✓ Gmail API               │
        │     ✓ Drive API               │
        │     ✓ Apps Script API         │
        │     ✓ Marketplace SDK         │
        └───────────────────────────────┘
                        │
                        ↓
                   [Phase 2]
```

---

### Phase 2: OAuth Consent Screen (10 min)

```
┌────────────────────────────────────────────────────────┐
│        APIs & Services → OAuth Consent Screen          │
└────────────────────────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  1. Select User Type          │
        │     ⦿ Internal                │
        │     ○ External                │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  2. App Information           │
        │     - Name                    │
        │     - Logo (128x128)          │
        │     - Support Email           │
        │     - Authorized Domain       │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  3. Add 8 OAuth Scopes        │
        │     gmail.readonly            │
        │     drive                     │
        │     script.locale             │
        │     gmail.addons.execute      │
        │     script.external_request   │
        │     +3 more...                │
        └───────────────────────────────┘
                        │
                        ↓
                   [Phase 3]
```

---

### Phase 3: Apps Script Deployment (20 min)

```
┌────────────────────────────────────────────────────────┐
│              https://script.google.com                 │
└────────────────────────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  1. Create New Project        │
        │     "Jira Attachment Saver"   │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  2. Link to GCP Project       │
        │     Project Settings →        │
        │     Change Project →          │
        │     Enter Project Number ⚠️   │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  3. Upload Code Files         │
        │     ✓ Code.js                 │
        │     ✓ appsscript.json         │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  4. Configure Properties      │
        │     Run: setupScriptProps()   │
        │     ↓ Authorize               │
        │     ↓ Update Values:          │
        │       - JIRA_URL              │
        │       - CUSTOMERS_FOLDER_ID   │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  5. Create Deployment         │
        │     Deploy → New Deployment   │
        │     Type: Add-on              │
        │     ↓                          │
        │     Note Deployment ID        │
        │     Note Script ID            │
        └───────────────────────────────┘
                        │
                        ↓
                   [Phase 4]
```

---

### Phase 4: Testing (15 min)

```
                ┌───────────────────────────────┐
                │  Install Test Version         │
                │  Deploy → Test Deployments    │
                └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────────────┐
        │           Open Gmail                  │
        │  https://mail.google.com              │
        └───────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────────────┐
        │  Test Checklist:                      │
        │  ✓ Add-on appears in sidebar          │
        │  ✓ Settings page loads                │
        │  ✓ Jira connection works              │
        │  ✓ Tickets fetch successfully         │
        │  ✓ Attachments save to Drive          │
        │  ✓ Folders auto-create                │
        └───────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────────────┐
        │  Review Execution Logs                │
        │  Apps Script → Executions             │
        │  ✓ No errors                          │
        └───────────────────────────────────────┘
                        │
                    [If OK] ↓
                   [Phase 5]
```

---

### Phase 5: Prepare Assets (30 min)

```
        ┌───────────────────────────────────────┐
        │  1. Create Screenshots                │
        │     Size: 1280x800px                  │
        │     Count: 3-5 images                 │
        │                                       │
        │     📷 Add-on in Gmail                │
        │     📷 Ticket selection               │
        │     📷 Settings page                  │
        │     📷 Success message                │
        │     📷 Drive folder structure         │
        └───────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────────────┐
        │  2. Prepare Text Content              │
        │     ✓ Short description (80 chars)    │
        │     ✓ Full description (4000 chars)   │
        │     ✓ Screenshot captions             │
        └───────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────────────┐
        │  3. Prepare Scope Justifications      │
        │     For each of 8 OAuth scopes        │
        └───────────────────────────────────────┘
                        │
                        ↓
                   [Phase 6]
```

---

### Phase 6: Marketplace Publishing (20 min)

```
┌────────────────────────────────────────────────────────┐
│  GCP Console → Google Workspace Marketplace SDK        │
│  → App Configuration → Create New Listing              │
└────────────────────────────────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  App Basics                   │
        │  - Name, Description          │
        │  - Icon, Category             │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  Extensions                   │
        │  + Add Gmail Add-on           │
        │    - Script ID                │
        │    - Deployment ID            │
        │  ✅ Verify Extension          │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  Store Listing                │
        │  - Upload Screenshots         │
        │  - Add Captions               │
        │  - Set Language               │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  Distribution ⚠️ IMPORTANT    │
        │  Visibility:                  │
        │  ⦿ Private → My domain only   │
        │  ○ Public                     │
        │                               │
        │  Installation:                │
        │  ⦿ Available to Install       │
        │  ○ Install for Everyone       │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  OAuth Scopes                 │
        │  ✓ Verify 8 scopes            │
        │  ✓ Add justifications         │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  Support & Contact            │
        │  - Developer info             │
        │  - Support email/URL          │
        │  - Privacy policy (optional)  │
        └───────────────────────────────┘
                        │
                        ↓
        ┌───────────────────────────────┐
        │  📤 PUBLISH / SUBMIT          │
        │  ↓                            │
        │  ⏳ Wait for Approval         │
        │  (Usually instant for         │
        │   internal apps)              │
        └───────────────────────────────┘
                        │
                   [Approved] ↓
                   [Phase 7]
```

---

### Phase 7: User Deployment (10 min + ongoing)

```
                    ┌─────────────┐
                    │  Published  │
                    │   App ✅    │
                    └─────────────┘
                          │
          ┌───────────────┴───────────────┐
          │                               │
          ↓                               ↓
  ┌──────────────────┐          ┌──────────────────┐
  │  Option A:       │          │  Option B:       │
  │  Admin Install   │          │  User Install    │
  └──────────────────┘          └──────────────────┘
          │                               │
          ↓                               ↓
  ┌──────────────────┐          ┌──────────────────┐
  │ Admin Console    │          │ Users install    │
  │ → Marketplace    │          │ from Marketplace │
  │ → Domain Install │          │ themselves       │
  │                  │          │                  │
  │ Install for:     │          │ Provide:         │
  │ ⦿ Everyone       │          │ - Instructions   │
  │ ○ Specific OUs   │          │ - Install link   │
  └──────────────────┘          └──────────────────┘
          │                               │
          └───────────────┬───────────────┘
                          ↓
          ┌───────────────────────────────┐
          │  User Configuration           │
          │  1. Open Gmail                │
          │  2. Click add-on icon         │
          │  3. Go to Settings            │
          │  4. Enter Jira credentials:   │
          │     - URL                     │
          │     - Email                   │
          │     - API Token               │
          │  5. Test Connection ✅        │
          └───────────────────────────────┘
                          │
                          ↓
          ┌───────────────────────────────┐
          │  🎉 DEPLOYMENT COMPLETE 🎉    │
          │                               │
          │  Users can now:               │
          │  ✓ Open emails                │
          │  ✓ Select Jira tickets        │
          │  ✓ Save attachments to Drive  │
          └───────────────────────────────┘
```

---

## Data Flow: Runtime Operation

Once deployed, here's how the add-on works:

```
┌─────────────┐
│   USER      │
│ Opens Email │
│ in Gmail    │
└──────┬──────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│          Gmail Add-on Activates                     │
│          (buildAddOn function)                      │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  1. Load Settings from User Properties              │
│     - Jira URL                                      │
│     - Jira Credentials (encrypted)                  │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  2. Fetch Jira Tickets                              │
│     ┌────────────┐                                  │
│     │  Jira API  │ ← HTTP Request with Auth         │
│     │  (REST)    │ → Returns JSON ticket list       │
│     └────────────┘                                  │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  3. Display UI                                      │
│     - Ticket dropdown                               │
│     - Attachment checkboxes (grouped by type)       │
│     - Subfolder selector                            │
│     - Action buttons                                │
└──────┬──────────────────────────────────────────────┘
       │
       ↓ (User selects & clicks Save)
       │
┌─────────────────────────────────────────────────────┐
│  4. Parse Jira Ticket Summary                       │
│     Input: "620254 - Frosta | TP | Project"        │
│     Extract:                                        │
│     - Customer: "620254 - Frosta"                   │
│     - Ticket: "CXPRODELIVERY-1234"                  │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  5. Find/Create Folder Structure                    │
│     ┌─────────────┐                                 │
│     │  Drive API  │ ← Search for customer folder    │
│     │  (v3)       │ → Create if not exists          │
│     └─────────────┘                                 │
│                                                      │
│     Result:                                         │
│     📁 Customers/                                   │
│     └─📁 620254 - Frosta/                          │
│       └─📁 CXPRODELIVERY-1234/                     │
│         └─📁 [Selected Subfolder]/                 │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  6. Download & Upload Attachments                  │
│     For each selected attachment:                   │
│     ┌─────────────┐         ┌─────────────┐        │
│     │  Gmail API  │ → Get → │  Drive API  │        │
│     │  Attachment │   File  │  Upload     │        │
│     └─────────────┘         └─────────────┘        │
│                                                      │
│     Features:                                       │
│     - Duplicate detection (size check)              │
│     - Timestamp rename if needed                    │
│     - Metadata preservation                         │
└──────┬──────────────────────────────────────────────┘
       │
       ↓
┌─────────────────────────────────────────────────────┐
│  7. Display Success Message                         │
│     ✅ Saved X attachments to:                      │
│     📁 [Folder Link]                                │
└─────────────────────────────────────────────────────┘
```

---

## Permission & Access Flow

```
┌──────────────────────────────────────────────────────┐
│                   USER HIERARCHY                     │
└──────────────────────────────────────────────────────┘

         ┌─────────────────────────┐
         │  Google Workspace       │
         │  Organization           │
         │  (transporeon.com)      │
         └───────────┬─────────────┘
                     │
         ┌───────────┴───────────┐
         │                       │
         ↓                       ↓
┌────────────────┐    ┌────────────────────┐
│  Admin User    │    │  Regular Users     │
│  - Install app │    │  - Use app         │
│  - Configure   │    │  - Configure own   │
│    for domain  │    │    Jira settings   │
└────────────────┘    └────────────────────┘
         │                       │
         └───────────┬───────────┘
                     ↓
         ┌─────────────────────────┐
         │  Gmail Add-on           │
         │  (Apps Script Runtime)  │
         └───────────┬─────────────┘
                     │
         ┌───────────┼───────────┐
         │           │           │
         ↓           ↓           ↓
   ┌─────────┐ ┌─────────┐ ┌─────────┐
   │ Gmail   │ │ Drive   │ │ Jira    │
   │ API     │ │ API     │ │ API     │
   │ (Read)  │ │ (R/W)   │ │ (Read)  │
   └─────────┘ └─────────┘ └─────────┘
```

### OAuth Scopes Required:

| Scope | Purpose | Risk Level |
|-------|---------|------------|
| `gmail.readonly` | Read email & attachments | 🟡 Medium |
| `drive` | Create folders & save files | 🟡 Medium |
| `gmail.addons.*` | Display UI in Gmail | 🟢 Low |
| `script.external_request` | Call Jira API | 🟢 Low |
| `userinfo.email` | Identify user | 🟢 Low |

---

## Common Issue Resolution Paths

### Issue: OAuth Errors

```
User gets "Authorization Error"
         │
         ↓
Check: Scopes match everywhere?
         │
    ┌────┴────┐
    NO        YES
    │         │
    ↓         ↓
Update:   Check: User in
- appsscript.json  correct domain?
- OAuth Screen     │
- Marketplace  ┌───┴───┐
    │          NO      YES
    │          │       │
    └──────────┴───────↓
              Fix domain/
              Re-authorize
```

### Issue: Add-on Not Visible

```
Add-on doesn't appear in Gmail
         │
         ↓
Check: Installed correctly?
         │
    ┌────┴────┐
    NO        YES
    │         │
    ↓         ↓
Install   Check: Right
from      deployment?
Admin/    │
Market    ├── Using HEAD? → Use Production
place     │
         ├── Cached? → Clear cache
         │
         └── Gmail view? → Use Default view
```

---

## Version Update Flow

When you need to update the app:

```
┌─────────────────────────┐
│  Make Code Changes      │
│  in Apps Script Editor  │
└───────────┬─────────────┘
            │
            ↓
┌─────────────────────────┐
│  Test with HEAD         │
│  (Test Deployment)      │
└───────────┬─────────────┘
            │
        [If OK] ↓
            │
┌─────────────────────────┐
│  Create New Deployment  │
│  Deploy → New           │
│  Version auto-increment │
└───────────┬─────────────┘
            │
            ↓
┌─────────────────────────┐
│  Users Auto-Update      │
│  (No reinstall needed)  │
│  Next time they open    │
│  the add-on             │
└─────────────────────────┘
```

**Note**: For major changes requiring new OAuth scopes:
1. Update `appsscript.json`
2. Update OAuth Consent Screen in GCP
3. Update Marketplace listing
4. Users will need to re-authorize

---

## Security & Compliance Checklist

```
┌───────────────────────────────────────────────────┐
│              SECURITY CHECKLIST                   │
├───────────────────────────────────────────────────┤
│                                                   │
│  ✓ Internal app only (domain-restricted)         │
│  ✓ OAuth scopes minimal & justified              │
│  ✓ Credentials stored in Properties Service      │
│  ✓ No hardcoded secrets in code                  │
│  ✓ HTTPS only for external API calls             │
│  ✓ Input validation on user data                 │
│  ✓ Error handling prevents info leakage          │
│  ✓ Logging excludes sensitive data               │
│  ✓ Regular security reviews scheduled            │
│                                                   │
└───────────────────────────────────────────────────┘
```

---

## Maintenance Schedule

```
┌─────────────────────────────────────────────────────┐
│                  MAINTENANCE PLAN                   │
├─────────────────────────────────────────────────────┤
│                                                     │
│  DAILY:                                             │
│  └─ Monitor execution logs for errors              │
│                                                     │
│  WEEKLY:                                            │
│  ├─ Review usage statistics                        │
│  ├─ Check for user issues                          │
│  └─ Update documentation if needed                 │
│                                                     │
│  MONTHLY:                                           │
│  ├─ Review quota usage                             │
│  ├─ Check for API updates                          │
│  └─ Gather user feedback                           │
│                                                     │
│  QUARTERLY:                                         │
│  ├─ Security audit                                 │
│  ├─ Performance review                             │
│  ├─ Feature planning                               │
│  └─ Update dependencies                            │
│                                                     │
│  ANNUALLY:                                          │
│  ├─ Comprehensive review                           │
│  ├─ OAuth scope audit                              │
│  └─ Compliance check                               │
│                                                     │
└─────────────────────────────────────────────────────┘
```

---

## Success Metrics

Track these KPIs after deployment:

```
┌─────────────────────────────────────────┐
│         METRICS DASHBOARD               │
├─────────────────────────────────────────┤
│                                         │
│  📊 Usage Metrics:                      │
│     • Active users (daily/weekly)       │
│     • Attachments saved per day         │
│     • Success rate (%)                  │
│                                         │
│  ⚡ Performance Metrics:                │
│     • Average execution time            │
│     • Error rate (%)                    │
│     • API quota usage (%)               │
│                                         │
│  👥 User Satisfaction:                  │
│     • Support tickets                   │
│     • Feature requests                  │
│     • User feedback scores              │
│                                         │
│  🔒 Security Metrics:                   │
│     • Failed auth attempts              │
│     • Permission denials                │
│     • Suspicious activity alerts        │
│                                         │
└─────────────────────────────────────────┘
```

---

## Quick Reference: Key Differences

### Internal vs External Apps

| Aspect | Internal App | External App |
|--------|--------------|--------------|
| **Review Time** | Instant | Days/Weeks |
| **OAuth Screen** | Internal | External + Verification |
| **User Base** | Domain only | Anyone |
| **Publishing Process** | Simple | Complex |
| **Security Review** | Basic | Extensive |
| **Branding** | Optional | Required |

### Our Choice: ✅ Internal App
- Faster deployment
- Domain-restricted access
- Simpler maintenance
- Better for internal tools

---

**Diagram Version**: 1.0  
**Last Updated**: 2025-12-12  
**Related Docs**: 
- [Full Deployment Guide](./deployment-guide.md)
- [Quick Checklist](./deployment-quick-checklist.md)


