# Deployment Documentation - Complete Guide

**Gmail Attachment Saver - GCP Deployment & Google Workspace Marketplace Publishing**

This folder contains comprehensive documentation for deploying the Gmail Attachment Saver add-on to a new Google Cloud Platform (GCP) project and publishing it as an Internal App to Google Workspace Marketplace.

---

## 📚 Documentation Files

### 1. **[deployment-guide.md](./deployment-guide.md)** - Master Deployment Guide
**Purpose**: Complete step-by-step instructions for the entire deployment process  
**Length**: Comprehensive (~200+ steps)  
**Best For**: First-time deployers, detailed reference

**Contents**:
- Prerequisites checklist
- Phase 1: GCP Project Setup
- Phase 2: OAuth Consent Screen Configuration
- Phase 3: Apps Script Deployment
- Phase 4: Testing Procedures
- Phase 5: Marketplace Asset Preparation
- Phase 6: Marketplace Publishing
- Phase 7: Post-Deployment & User Onboarding
- Common troubleshooting scenarios

**When to Use**: 
- Your first deployment
- Need detailed explanations
- Want to understand each step's purpose
- Training new team members

---

### 2. **[deployment-quick-checklist.md](./deployment-quick-checklist.md)** - Quick Reference
**Purpose**: Condensed checklist for experienced deployers  
**Length**: Concise checklists with key actions  
**Best For**: Quick reference, subsequent deployments

**Contents**:
- Pre-flight checklist
- Phase-by-phase checkboxes
- Time estimates for each phase
- Key URLs and IDs to track
- Success criteria
- Quick troubleshooting table

**When to Use**:
- You've deployed once and need a reminder
- Quick reference during deployment
- Tracking deployment progress
- Subsequent version updates

---

### 3. **[deployment-workflow-diagram.md](./deployment-workflow-diagram.md)** - Visual Guide
**Purpose**: Visual representation of the deployment process  
**Length**: Diagrams and flowcharts  
**Best For**: Understanding the big picture, presentations

**Contents**:
- High-level workflow overview
- Detailed phase diagrams with arrows and decision points
- Data flow diagrams
- Permission hierarchy charts
- Issue resolution flowcharts
- Runtime operation visualization

**When to Use**:
- Need to visualize the process
- Explaining deployment to stakeholders
- Understanding relationships between components
- Planning deployment timeline

---

### 4. **[deployment-troubleshooting.md](./deployment-troubleshooting.md)** - Problem Solving
**Purpose**: Solutions to common deployment issues  
**Length**: Extensive issue database with solutions  
**Best For**: When something goes wrong

**Contents**:
- GCP project issues
- OAuth configuration problems
- Apps Script deployment errors
- Testing failures
- Marketplace publishing issues
- User installation problems
- Runtime errors
- Performance optimization

**When to Use**:
- Encountering errors or unexpected behavior
- Users reporting issues
- Performance problems
- Need debugging strategies

---

## 🚀 Getting Started - Where to Begin

### For First-Time Deployment:

**Recommended Reading Order**:

1. **Start**: Read [deployment-guide.md](./deployment-guide.md) - Introduction & Prerequisites
2. **Visualize**: Skim [deployment-workflow-diagram.md](./deployment-workflow-diagram.md) - Overview section
3. **Prepare**: Use [deployment-quick-checklist.md](./deployment-quick-checklist.md) - Pre-Flight Check
4. **Execute**: Follow [deployment-guide.md](./deployment-guide.md) - Phases 1-7
5. **Track Progress**: Check off items in [deployment-quick-checklist.md](./deployment-quick-checklist.md)
6. **If Issues**: Refer to [deployment-troubleshooting.md](./deployment-troubleshooting.md)

**Estimated Time**: 2-3 hours for complete deployment

---

### For Subsequent Deployments:

**Recommended Approach**:

1. **Quick Check**: [deployment-quick-checklist.md](./deployment-quick-checklist.md) - Verify prerequisites
2. **Execute**: Follow checklist phases
3. **Reference**: [deployment-guide.md](./deployment-guide.md) - Only if you forget details
4. **If Issues**: [deployment-troubleshooting.md](./deployment-troubleshooting.md)

**Estimated Time**: 30-60 minutes for version updates

---

## 📋 Quick Reference Matrix

| Scenario | Primary Doc | Secondary Doc |
|----------|-------------|---------------|
| **First deployment** | deployment-guide.md | deployment-quick-checklist.md |
| **Update/new version** | deployment-quick-checklist.md | deployment-guide.md (reference) |
| **Training someone** | deployment-guide.md | deployment-workflow-diagram.md |
| **Presentation to stakeholders** | deployment-workflow-diagram.md | deployment-guide.md (details) |
| **Something broke** | deployment-troubleshooting.md | deployment-guide.md (context) |
| **User can't install** | deployment-troubleshooting.md → User Installation Issues | deployment-guide.md → Phase 7 |
| **Performance issues** | deployment-troubleshooting.md → Performance Issues | deployment-guide.md → Phase 4 |
| **Publishing failed** | deployment-troubleshooting.md → Marketplace Publishing | deployment-guide.md → Phase 6 |

---

## 🎯 Deployment Phases Overview

### Phase 1: GCP Project Setup (15 min)
**Objective**: Create and configure Google Cloud Platform project  
**Key Actions**:
- Create GCP project
- Enable 4 required APIs
- Note Project ID and Project Number

**Critical Success Factor**: Getting the correct Project Number (numeric)

---

### Phase 2: OAuth Consent Screen (10 min)
**Objective**: Configure user authentication and permissions  
**Key Actions**:
- Set to "Internal" user type
- Add app information
- Configure all 8 OAuth scopes

**Critical Success Factor**: All scopes must match exactly across all configurations

---

### Phase 3: Apps Script Deployment (20 min)
**Objective**: Deploy the add-on code  
**Key Actions**:
- Create Apps Script project
- Link to GCP (using Project Number)
- Upload code files
- Configure Script Properties
- Create production deployment

**Critical Success Factor**: Script Properties must be set correctly (especially CUSTOMERS_FOLDER_ID)

---

### Phase 4: Testing (15 min)
**Objective**: Verify functionality before publishing  
**Key Actions**:
- Install test version
- Test all major features
- Verify Jira connection
- Test attachment saving
- Review execution logs

**Critical Success Factor**: All core features work without errors

---

### Phase 5: Assets Preparation (30 min)
**Objective**: Prepare marketplace listing materials  
**Key Actions**:
- Create 3-5 screenshots (1280x800px)
- Write descriptions
- Prepare scope justifications

**Critical Success Factor**: Professional, clear screenshots showing key features

---

### Phase 6: Marketplace Publishing (20 min)
**Objective**: Publish app to Google Workspace Marketplace  
**Key Actions**:
- Configure Marketplace SDK
- Add Gmail Add-on extension
- Set visibility to "Internal"
- Submit for publishing

**Critical Success Factor**: Visibility must be "Private → My domain only" for internal app

---

### Phase 7: User Deployment (10 min + ongoing)
**Objective**: Make app available to users  
**Key Actions**:
- Admin installation OR user self-installation
- User configuration (Jira credentials)
- Documentation sharing

**Critical Success Factor**: Clear user instructions and support availability

---

## 🔧 Essential Information to Track

### IDs You'll Need:

| ID Type | Where to Find | Example Format | Used For |
|---------|---------------|----------------|----------|
| **GCP Project ID** | GCP Console Dashboard | `gmail-saver-123` | General GCP reference |
| **GCP Project Number** | GCP Console Dashboard | `123456789012` | Linking Apps Script |
| **Apps Script ID** | Apps Script → Project Settings | `1a2b3c...xyz` | Marketplace SDK |
| **Deployment ID** | Apps Script → Deploy → Manage | `AKfycby...` | Marketplace SDK |
| **Customers Folder ID** | Drive folder URL | `1pi6fJzg...` | Script Properties |

### URLs You'll Use:

| Purpose | URL |
|---------|-----|
| **GCP Console** | https://console.cloud.google.com/ |
| **Apps Script Editor** | https://script.google.com/ |
| **Gmail** | https://mail.google.com/ |
| **Admin Console** | https://admin.google.com/ |
| **Jira API Tokens** | https://id.atlassian.com/manage-profile/security/api-tokens |

---

## ✅ Pre-Deployment Checklist

Before you start, ensure you have:

### Access & Permissions:
- [ ] Google Workspace account with admin rights
- [ ] GCP project creation permissions
- [ ] Jira instance access
- [ ] Google Drive access

### Technical Requirements:
- [ ] All project files ready (`Code.js`, `appsscript.json`)
- [ ] Logo image prepared (128x128px recommended)
- [ ] Jira URL known
- [ ] Jira API token generated (or ability to generate)
- [ ] Google Drive "Customers" folder created
- [ ] Customers folder ID obtained

### Knowledge:
- [ ] Read through deployment-guide.md introduction
- [ ] Understand your organization's requirements
- [ ] Know who will use the app
- [ ] Have support plan in place

---

## 🚨 Common Pitfalls to Avoid

1. **❌ Using Project ID instead of Project Number** when linking Apps Script
   - ✅ Use the numeric Project Number (e.g., `123456789012`)

2. **❌ Missing OAuth scopes** or scope mismatches
   - ✅ Ensure all 8 scopes match in appsscript.json, OAuth screen, and Marketplace

3. **❌ Forgetting to run setupScriptProperties()**
   - ✅ Must run this function and verify properties are set

4. **❌ Wrong folder permissions**
   - ✅ All users need access to Customers folder (consider Shared Drive)

5. **❌ Selecting "External" for OAuth consent**
   - ✅ Use "Internal" for faster approval and domain restriction

6. **❌ Not testing before publishing**
   - ✅ Always test with real emails and attachments first

7. **❌ Using HEAD deployment for production**
   - ✅ Create proper versioned deployment for Marketplace

8. **❌ Not updating CUSTOMERS_FOLDER_ID** in Script Properties
   - ✅ Must match your actual Google Drive folder

---

## 📊 Success Criteria

Your deployment is complete and successful when:

### Technical Success:
- ✅ Add-on appears in Gmail sidebar
- ✅ Jira connection test passes
- ✅ Tickets load in dropdown
- ✅ Attachments can be selected
- ✅ Files save to correct Drive folders
- ✅ Folder structure auto-creates properly
- ✅ No errors in execution logs

### User Success:
- ✅ Users can find and install the app
- ✅ Users can configure settings
- ✅ Users successfully save attachments
- ✅ Users report positive experience

### Organizational Success:
- ✅ App published to marketplace
- ✅ Installation process documented
- ✅ Support process established
- ✅ Usage being tracked

---

## 🔄 Maintenance & Updates

### Regular Maintenance:

**Daily** (first week):
- Monitor execution logs
- Respond to user questions
- Track error rates

**Weekly**:
- Review usage statistics
- Check for common issues
- Update documentation as needed

**Monthly**:
- Review quota usage
- Gather user feedback
- Plan feature improvements

**Quarterly**:
- Security audit
- Performance review
- Consider feature additions

### Deploying Updates:

When you need to update the app:

1. Make changes in Apps Script editor
2. Test with HEAD deployment
3. Create new deployment (Deploy → New deployment)
4. Optional: Update Marketplace listing if UI/features changed significantly
5. Users automatically get updates (no reinstall needed)

**Note**: If adding new OAuth scopes, users must re-authorize.

---

## 📞 Support Resources

### Internal Support:
- **Developer**: [Your contact info]
- **Admin**: [Admin contact info]
- **Documentation**: This folder

### External Resources:
- **Apps Script Docs**: https://developers.google.com/apps-script
- **Gmail Add-ons**: https://developers.google.com/apps-script/add-ons/gmail
- **Marketplace Guide**: https://developers.google.com/workspace/marketplace
- **Stack Overflow**: Tag `google-apps-script` and `gmail-addon`
- **Community**: https://support.google.com/code/community

---

## 🎓 Training & Onboarding

### For New Developers:

**Week 1: Understanding**
- Read all 4 documentation files
- Review Code.js to understand functionality
- Study Apps Script basics: https://developers.google.com/apps-script/overview

**Week 2: Practice Deployment**
- Create test GCP project
- Follow deployment guide completely
- Deploy to personal workspace for practice
- Break and fix things to learn

**Week 3: Production**
- Shadow experienced deployer
- Deploy to production with supervision
- Document any new learnings

### For New Users:

**Onboarding Checklist**:
- [ ] Install add-on (via marketplace or admin-pushed)
- [ ] Open any email in Gmail
- [ ] Click add-on icon
- [ ] Go to Settings (⚙️)
- [ ] Configure Jira credentials
- [ ] Test Jira connection
- [ ] Try saving one attachment
- [ ] Verify file in Drive
- [ ] Review user guide (README.md)

---

## 📈 Metrics to Track

After deployment, monitor:

### Usage Metrics:
- Active users (daily/weekly)
- Attachments saved per day
- Most active times/days
- Feature usage (subfolder selection, etc.)

### Performance Metrics:
- Average execution time
- Success rate (%)
- Error rate (%)
- Timeout frequency

### Quality Metrics:
- User satisfaction surveys
- Support ticket volume
- Feature requests
- Bug reports

### Resource Metrics:
- Apps Script quota usage
- Gmail API quota usage
- Drive API quota usage
- Storage usage in Drive

---

## 🎯 Next Steps After Reading This

1. **Assess Your Readiness**:
   - Review pre-deployment checklist above
   - Gather all required access and information
   - Block 2-3 hours in your calendar

2. **Choose Your Path**:
   - First deployment? → Start with [deployment-guide.md](./deployment-guide.md)
   - Quick update? → Use [deployment-quick-checklist.md](./deployment-quick-checklist.md)
   - Need overview? → Review [deployment-workflow-diagram.md](./deployment-workflow-diagram.md)

3. **Prepare Your Environment**:
   - Open all necessary tabs:
     - GCP Console
     - Apps Script Editor
     - Gmail
     - Documentation files
   - Have credentials ready:
     - Google Workspace admin login
     - Jira API token
     - Drive folder ID

4. **Begin Deployment**:
   - Follow your chosen guide
   - Check off items as you complete them
   - Keep notes of any issues for future reference

5. **Post-Deployment**:
   - Test thoroughly
   - Document any deviations from guide
   - Share with team
   - Set up monitoring

---

## 🤝 Contributing to Documentation

If you find issues or improvements:

1. **During Deployment**:
   - Note any steps that weren't clear
   - Document any errors and solutions you found
   - Track any steps that took longer than estimated

2. **After Deployment**:
   - Update documentation with your learnings
   - Add any new troubleshooting scenarios
   - Update time estimates if significantly different
   - Add screenshots if helpful

3. **Continuous Improvement**:
   - Keep docs updated with latest GCP/Apps Script changes
   - Document new features as they're added
   - Maintain troubleshooting database

---

## 📝 Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 1.0 | 2025-12-12 | Initial comprehensive deployment documentation | AI Assistant |
| | | Created 4 main deployment documents | |
| | | Covers GCP setup through marketplace publishing | |

---

## 📄 Document Structure

```
Documentation/
├── README.md (this file)                          ← Start here
├── deployment-guide.md                            ← Complete guide
├── deployment-quick-checklist.md                  ← Quick reference
├── deployment-workflow-diagram.md                 ← Visual guide
├── deployment-troubleshooting.md                  ← Problem solving
├── Code-js-function-reference.md                  ← Code documentation
└── customer-folder-link-implementation.md         ← Feature docs
```

---

## ✨ Quick Tips for Success

1. **Don't Rush**: First deployment takes time. That's normal.
2. **Test Thoroughly**: Test in Gmail before publishing to marketplace.
3. **Keep Notes**: Document your specific setup for future reference.
4. **Use Incognito**: Test authorization flow in incognito window.
5. **Screenshot Everything**: Capture settings for documentation.
6. **Start Small**: Test with 1-2 users before rolling out to everyone.
7. **Monitor Closely**: Watch logs daily for the first week.
8. **Be Patient**: Some changes take time to propagate (up to 24 hours).
9. **Ask for Help**: Use support resources if stuck.
10. **Document Changes**: Keep track of customizations for your environment.

---

## 🎉 Conclusion

You now have everything you need to successfully deploy the Gmail Attachment Saver add-on to your Google Cloud Platform project and publish it to Google Workspace Marketplace as an internal app.

**Remember**:
- Take your time
- Follow the guides
- Test thoroughly
- Ask for help when needed

**Good luck with your deployment!** 🚀

---

**Documentation Version**: 1.0  
**Last Updated**: 2025-12-12  
**Maintained By**: Your Team  
**Questions?**: Contact your admin or developer

---

## License & Usage

This documentation is part of the Gmail Attachment Saver project and is intended for internal use within your organization.



