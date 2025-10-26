// ===== MAIN GMAIL ADD-ON FUNCTIONS =====

function buildAddOn(e) {
    console.log("=== BUILD ADD-ON CALLED ===");
    
    try {
      var card = CardService.newCardBuilder();
      var section = CardService.newCardSection()
        .setHeader("Jira TPM Attachment Saver");
      
      // Check if settings are configured
      var settings = getUserSettings();
      if (!settings.jiraToken || !settings.jiraUrl || !settings.pmoWebhookUrl) {
        console.log("Settings not configured (missing Jira or PMO settings), showing setup");
        return [buildSettingsCard(true)]; // true = first time setup
      }
      
      // Get user's TPM projects and active tickets
      try {
        console.log("Loading TPM projects...");
        var projectResult = getMyJiraProjects();
        var projects = projectResult.projects || {};
        var issues = projectResult.issues || [];
        
        console.log("DEBUG: Projects loaded:", Object.keys(projects).length);
        console.log("DEBUG: Issues loaded:", issues.length);
        
        if (Object.keys(projects).length > 0 && issues.length > 0) {
          console.log("Creating TPM project UI with", issues.length, "tickets...");
          
          // Compact header with just settings button
          var headerSection = CardService.newCardSection();
          
          // Compact settings button (gear only)
          var settingsButtonSet = CardService.newButtonSet()
            .addButton(CardService.newTextButton()
              .setText("⚙️")
              .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
          headerSection.addWidget(settingsButtonSet);
          
          card.addSection(headerSection);
          
          // Create dropdown with active TPM tickets WITH onChange action
          console.log("Creating dropdown selection widget with dynamic info...");
          var ticketSelection = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setFieldName("selectedTicket")
            .setTitle("Select Active Ticket")
            .setOnChangeAction(CardService.newAction()
              .setFunctionName("showTicketDetails")
              .setParameters({threadId: (e && e.gmail) ? e.gmail.threadId : ""}));
          
          // Add "Manual Entry" option first
          console.log("Adding manual entry option...");
          ticketSelection.addItem("Manual Entry - Enter ticket number below", "manual", true);
          
          // Sort and add all active TPM tickets to dropdown
          issues.sort(function(a, b) {
            return a.key.localeCompare(b.key);
          });
          
          console.log("Adding", issues.length, "TPM tickets to dropdown...");
          
          var successfulItems = 0;
          for (var i = 0; i < issues.length; i++) {
            try {
              var issue = issues[i];
              
              // Create compact display text with emoji and shortened info
              var displayText = formatCompactTicketDisplay(issue);
              
              console.log("Adding ticket:", displayText);
              ticketSelection.addItem(displayText, issue.key, false);
              successfulItems++;
              
            } catch (itemError) {
              console.error("Error adding dropdown item for", issue.key, ":", itemError);
              // Continue with other items
              continue;
            }
          }
          
          console.log("Successfully added", successfulItems, "items to dropdown");
          section.addWidget(ticketSelection);
          
          // Add manual ticket number input
          var ticketInput = CardService.newTextInput()
            .setFieldName("manualTicketNumber")
            .setTitle("Manual Ticket Entry")
            .setHint("e.g., 6500 (will become CXPRODELIVERY-6500)");
          section.addWidget(ticketInput);
          
          // Add helpful tip
          var statusInfo = CardService.newTextParagraph()
            .setText("💡 Tip: Select from dropdown to see full ticket details");
          section.addWidget(statusInfo);
          
          console.log("Created main ticket selection interface with dropdown and dynamic info");
          
        } else {
          console.log("No TPM projects/issues found, using fallback UI");
          var debugInfo = CardService.newTextParagraph()
            .setText("⚠️ Could not load TPM tickets. Check your JQL or connection settings.");
          section.addWidget(debugInfo);
          
                // Settings button for troubleshooting
        var settingsButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("⚙️ Check Settings")
            .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
        section.addWidget(settingsButtonSet);
          
          var ticketInput = CardService.newTextInput()
            .setFieldName("jiraTicket")
            .setTitle("Manual Ticket Entry")
            .setHint("e.g., CXPRODELIVERY-6310");
          section.addWidget(ticketInput);
        }
        
      } catch (jiraError) {
        console.error('Failed to load TPM projects:', jiraError);
        var errorInfo = CardService.newTextParagraph()
          .setText("⚠️ Jira API error: " + jiraError.message + ". Check your settings.");
        section.addWidget(errorInfo);
        
        // Settings button for troubleshooting
        var settingsButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("⚙️ Fix Settings")
            .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
        section.addWidget(settingsButtonSet);
        
        var ticketInput = CardService.newTextInput()
          .setFieldName("jiraTicket")
          .setTitle("Full Jira Ticket")
          .setHint("e.g., CXPRODELIVERY-6310");
        section.addWidget(ticketInput);
      }
      
      // Handle Gmail context for attachments
      if (!e || !e.gmail || !e.gmail.threadId) {
        console.log("No Gmail context - using basic interface");
        var infoSection = CardService.newCardSection()
          .setHeader("Open an email to see attachments");
        
        var infoText = CardService.newTextParagraph()
          .setText("Please open an email with attachments to use this add-on.");
        infoSection.addWidget(infoText);
        card.addSection(infoSection);
        
        // Add save button
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Save to PMO Folder")
            .setOnClickAction(CardService.newAction().setFunctionName("saveSelectedAttachmentsToGDrive")));
        section.addWidget(saveButtonSet);
        
        card.addSection(section);
        
        return [card.build()];
      }
      
      console.log("Gmail context found, loading attachments...");
      
      // Get attachments from the email thread
      try {
        var threadId = e.gmail.threadId;
        var thread = GmailApp.getThreadById(threadId);
        var messages = thread.getMessages();
        var allAttachments = [];
        
        for (var i = 0; i < messages.length; i++) {
          var message = messages[i];
          var attachments = message.getAttachments();
          
          for (var j = 0; j < attachments.length; j++) {
            var attachment = attachments[j];
            allAttachments.push({
              name: attachment.getName(),
              size: attachment.getSize(),
              attachment: attachment,
              messageIndex: i
            });
          }
        }
        
                if (allAttachments.length > 0) {
            // Group attachments by file extension
            var attachmentsByExtension = {};
            var extensionCounts = {};
            
            for (var i = 0; i < allAttachments.length; i++) {
              var attachment = allAttachments[i];
              var fileName = attachment.name;
              var extension = fileName.lastIndexOf('.') > -1 ? 
                fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              
              if (!attachmentsByExtension[extension]) {
                attachmentsByExtension[extension] = [];
                extensionCounts[extension] = 0;
              }
              
              attachmentsByExtension[extension].push({
                attachment: attachment,
                index: i
              });
              extensionCounts[extension]++;
            }
            
            // Get stored attachment selections for this thread
            var storedSelections = getStoredAttachmentSelections(threadId);
            
            // Sort extensions alphabetically for consistent display
            var sortedExtensions = Object.keys(attachmentsByExtension).sort();
            
            // Create a section for each file extension
            for (var extIndex = 0; extIndex < sortedExtensions.length; extIndex++) {
              var extension = sortedExtensions[extIndex];
              var attachmentsInGroup = attachmentsByExtension[extension];
              var count = extensionCounts[extension];
              
              // Create section header with extension and count (show total in first section)
              var extDisplayName = extension === 'no-extension' ? 'Files without extension' : 
                extension.toUpperCase() + ' files';
              var headerText = extDisplayName + " (" + count + ")";
              if (extIndex === 0) {
                headerText = "📎 Select Attachments: " + headerText + " (Total: " + allAttachments.length + ")";
              }
              
              var extensionSection = CardService.newCardSection()
                .setHeader(headerText)
                .setCollapsible(true)
                .setNumUncollapsibleWidgets(count > 5 ? 2 : count); // Show first 2 if more than 5 files
              
              // Add checkboxes for each attachment in this extension group
              for (var j = 0; j < attachmentsInGroup.length; j++) {
                var attachmentData = attachmentsInGroup[j];
                var attachment = attachmentData.attachment;
                var originalIndex = attachmentData.index;
                
                // Check if this attachment was previously selected using unique key
                var isSelected = false; // Default to false - let user explicitly choose
                var uniqueKey = attachment.name + "_" + originalIndex;
                if (storedSelections && storedSelections.hasOwnProperty(uniqueKey)) {
                  isSelected = storedSelections[uniqueKey];
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(attachment.name + " (" + formatFileSize(attachment.size) + ")", originalIndex.toString(), isSelected);
                
                extensionSection.addWidget(checkbox);
              }
              
              card.addSection(extensionSection);
            }
          } else {
          var noAttachSection = CardService.newCardSection()
            .setHeader("No Attachments Found");
          
          var noAttachText = CardService.newTextParagraph()
            .setText("This email thread doesn't contain any attachments.");
          noAttachSection.addWidget(noAttachText);
          card.addSection(noAttachSection);
        }
      } catch (attachmentError) {
        console.error("Error loading attachments:", attachmentError);
      }
      
      // Add main section to card first
      card.addSection(section);
      
      // Save button - add as separate section
      var saveSection = CardService.newCardSection();
      var saveAction = CardService.newAction()
        .setFunctionName("saveSelectedAttachmentsToGDrive")
        .setParameters({threadId: (e && e.gmail) ? e.gmail.threadId : ""});
      
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("Save to PMO Folder")
          .setOnClickAction(saveAction));
      
      saveSection.addWidget(saveButtonSet);
      card.addSection(saveSection);
      
      console.log("Card built successfully");
      return [card.build()];
      
    } catch (error) {
      console.error("Critical error in buildAddOn:", error);
      
      // Fallback card with settings option
      var fallbackCard = CardService.newCardBuilder();
      var fallbackSection = CardService.newCardSection()
        .setHeader("Error - Check Settings");
      
      var errorText = CardService.newTextParagraph()
        .setText("Error: " + error.message);
      fallbackSection.addWidget(errorText);
      
      var settingsButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("⚙️ Configure Settings")
          .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
      fallbackSection.addWidget(settingsButtonSet);
      
      var ticketInput = CardService.newTextInput()
        .setFieldName("jiraTicket")
        .setTitle("Jira Ticket")
        .setHint("e.g., CXPRODELIVERY-6310");
      fallbackSection.addWidget(ticketInput);
      
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("Save to PMO Folder")
          .setOnClickAction(CardService.newAction().setFunctionName("saveSelectedAttachmentsToGDrive")));
      fallbackSection.addWidget(saveButtonSet);
      
      fallbackCard.addSection(fallbackSection);
      return [fallbackCard.build()];
    }
  }
  
  // ===== SETTINGS FUNCTIONS =====
  
  function buildSettingsCard(isFirstTime) {
    console.log("=== BUILD SETTINGS CARD ===");
    
    try {
      var card = CardService.newCardBuilder();
      console.log("Card builder created successfully");
      
      // Get current settings first and validate
      var settings = getUserSettings();
      console.log("Retrieved settings:", settings ? "Success" : "Failed");
    
    // Header section
    var headerSection = CardService.newCardSection()
      .setHeader(isFirstTime ? "⚙️ Initial Setup Required" : "⚙️ Settings");
    
    if (isFirstTime) {
      var setupInfo = CardService.newTextParagraph()
        .setText("Welcome! Please configure your Jira connection to get started.");
      headerSection.addWidget(setupInfo);
    }
    
    card.addSection(headerSection);
    console.log("Header section added successfully");
    
    // Settings section
    console.log("Creating settings section...");
    var settingsSection = CardService.newCardSection()
      .setHeader("Jira Configuration");
    console.log("Settings section created successfully");
    
    // Jira URL input
    console.log("Creating Jira URL input...");
    var jiraUrlInput = CardService.newTextInput()
      .setFieldName("jiraUrl")
      .setTitle("Jira URL")
      .setHint("e.g., https://your-company.atlassian.net")
      .setValue(settings.jiraUrl || "https://support.transporeon.com");
    settingsSection.addWidget(jiraUrlInput);
    console.log("Jira URL input added successfully");
    
    // Jira Token input
    var jiraTokenInput = CardService.newTextInput()
      .setFieldName("jiraToken")
      .setTitle("Jira API Token")
      .setHint("Your personal Jira API token")
      .setValue(settings.jiraToken || "");
    settingsSection.addWidget(jiraTokenInput);
    
    // Instructions for getting token
    var tokenHelp = CardService.newTextParagraph()
      .setText("📖 How to get API token:\n1. Go to Jira → Profile → Security\n2. Create API token\n3. Copy and paste above");
    settingsSection.addWidget(tokenHelp);
    
    // JQL input - using regular text input instead of multiline
    var jqlInput = CardService.newTextInput()
      .setFieldName("customJql")
      .setTitle("Custom JQL Query")
      .setHint("Optional: Custom JQL to filter your tickets")
      .setValue(settings.customJql || getDefaultJQL());
    settingsSection.addWidget(jqlInput);
    
    // JQL help
    var jqlHelp = CardService.newTextParagraph()
      .setText("💡 JQL Examples:\n• assignee = currentUser()\n• project = MYPROJECT AND status = 'In Progress'\n• 'Technical Project Manager' in (currentUser())");
    settingsSection.addWidget(jqlHelp);
    
    card.addSection(settingsSection);
    
    // PMO Integration Settings Section
    var pmoSection = CardService.newCardSection()
      .setHeader("PMO Project Folders (Required)");
    
    // PMO Webhook URL
    var pmoUrlInput = CardService.newTextInput()
      .setFieldName("pmoWebhookUrl")
      .setTitle("PMO Webhook URL")
      .setHint("URL for PMO project folder lookup (required)")
      .setValue(settings.pmoWebhookUrl || getDefaultPMOWebhookUrl());
    pmoSection.addWidget(pmoUrlInput);
    
    // PMO Timeout Setting
    var pmoTimeoutInput = CardService.newTextInput()
      .setFieldName("pmoTimeout")
      .setTitle("PMO Timeout (ms)")
      .setHint("Timeout for PMO webhook calls (default: 10000)")
      .setValue(String(settings.pmoTimeout || 10000));
    pmoSection.addWidget(pmoTimeoutInput);
    
    // PMO Retry Attempts
    var pmoRetryInput = CardService.newTextInput()
      .setFieldName("pmoRetryAttempts")
      .setTitle("PMO Retry Attempts")
      .setHint("Retry attempts for undefined responses (default: 2)")
      .setValue(String(settings.pmoRetryAttempts || 2));
    pmoSection.addWidget(pmoRetryInput);
    
    // PMO Help Text
    var pmoHelpText = CardService.newTextParagraph()
      .setText("📋 PMO Integration (Mandatory):\n• All attachments saved to PMO-managed project folders\n• PMO webhook automatically creates folders if they don't exist\n• Retry logic handles folder creation timing issues");
    pmoSection.addWidget(pmoHelpText);
    
    card.addSection(pmoSection);
    console.log("Settings section added to card successfully");
    
    // Action buttons section
    console.log("Creating button section...");
    var buttonSection = CardService.newCardSection();
    console.log("Button section created successfully");
    
    // Save button
    var saveButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("💾 Save Settings")
        .setOnClickAction(CardService.newAction().setFunctionName("saveSettings")));
    buttonSection.addWidget(saveButtonSet);
    
    // Test buttons
    var testButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🧪 Test Jira Connection")
        .setOnClickAction(CardService.newAction().setFunctionName("testJiraConnection")));
    buttonSection.addWidget(testButtonSet);
    
    var testPMOButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🧪 Test PMO Connection")
        .setOnClickAction(CardService.newAction().setFunctionName("testPMOConnection")));
    buttonSection.addWidget(testPMOButtonSet);
    
    // Back button (only if not first time)
    if (!isFirstTime) {
      var backButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("← Back to Main")
          .setOnClickAction(CardService.newAction().setFunctionName("backToMain")));
      buttonSection.addWidget(backButtonSet);
    }
    
    card.addSection(buttonSection);
    
    // Current settings display (if configured)
    if (settings.jiraUrl && settings.jiraToken) {
      var currentSection = CardService.newCardSection()
        .setHeader("Current Configuration")
        .setCollapsible(true)
        .setNumUncollapsibleWidgets(0);
      
      var currentInfo = CardService.newTextParagraph()
        .setText("🔗 URL: " + settings.jiraUrl + "\n🔑 Token: " + maskToken(settings.jiraToken) + "\n📋 JQL: " + (settings.customJql ? "Custom" : "Default"));
      currentSection.addWidget(currentInfo);
      
      card.addSection(currentSection);
    }
    
      console.log("About to build settings card");
      var builtCard = card.build();
      console.log("Settings card built successfully");
      return builtCard;
      
    } catch (error) {
      console.error("Error in buildSettingsCard:", error);
      
      // Return a basic fallback card
      var fallbackCard = CardService.newCardBuilder();
      var fallbackSection = CardService.newCardSection()
        .setHeader("Settings Error");
      
      var errorText = CardService.newTextParagraph()
        .setText("Error loading settings: " + error.message);
      fallbackSection.addWidget(errorText);
      
      fallbackCard.addSection(fallbackSection);
      return fallbackCard.build();
    }
  }
  
  function showSettings(e) {
    console.log("=== SHOW SETTINGS ===");
    
    try {
      // Build a simplified settings card with single section
      var card = CardService.newCardBuilder();
      
      // Get current settings
      var settings = getUserSettings();
      
      // Single section with everything
      var mainSection = CardService.newCardSection()
        .setHeader("⚙️ Jira Settings");
      
      // Jira URL input with current value
      var jiraUrlInput = CardService.newTextInput()
        .setFieldName("jiraUrl")
        .setTitle("Jira URL")
        .setHint("e.g., https://your-company.atlassian.net")
        .setValue(settings.jiraUrl || "https://support.transporeon.com");
      mainSection.addWidget(jiraUrlInput);
      
      // Jira Token input with masked value for security
      var maskedToken = "";
      if (settings.jiraToken) {
        // Show only first 4 and last 4 characters with asterisks in between
        if (settings.jiraToken.length > 8) {
          maskedToken = settings.jiraToken.substring(0, 4) + "****" + settings.jiraToken.substring(settings.jiraToken.length - 4);
        } else {
          maskedToken = "****" + settings.jiraToken.substring(settings.jiraToken.length - 2);
        }
      }
      
      var jiraTokenInput = CardService.newTextInput()
        .setFieldName("jiraToken")
        .setTitle("Jira API Token")
        .setHint("Leave blank to keep current token, or enter new token")
        .setValue(maskedToken);
      mainSection.addWidget(jiraTokenInput);
      
      // Instructions
      var tokenHelp = CardService.newTextParagraph()
        .setText("📖 How to get API token:\n1. Go to Jira → Profile → Manage Account → Security\n2. Create API token\n3. Copy the generated token (NOT your password)\n\n🔒 Security: Current token is masked for security. Leave unchanged to keep current token.");
      mainSection.addWidget(tokenHelp);
      
      // JQL input as regular text input
      var jqlInput = CardService.newTextInput()
        .setFieldName("customJql")
        .setTitle("Custom JQL Query")
        .setHint("Optional: Custom JQL to filter your tickets")
        .setValue(settings.customJql || getDefaultJQL());
      mainSection.addWidget(jqlInput);
      
      // JQL help
      var jqlHelp = CardService.newTextParagraph()
        .setText("💡 JQL Examples:\n• assignee = currentUser()\n• project = MYPROJECT AND status = 'In Progress'\n• 'Technical Project Manager' in (currentUser())");
      mainSection.addWidget(jqlHelp);
      
      // PMO Integration Settings Section
      var pmoSeparator = CardService.newTextParagraph()
        .setText("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
      mainSection.addWidget(pmoSeparator);
      
      var pmoHeaderText = CardService.newTextParagraph()
        .setText("📁 **PMO Project Folders (Required)**");
      mainSection.addWidget(pmoHeaderText);
      
      // PMO Webhook URL
      var pmoUrlInput = CardService.newTextInput()
        .setFieldName("pmoWebhookUrl")
        .setTitle("PMO Webhook URL")
        .setHint("URL for PMO project folder lookup (required)")
        .setValue(settings.pmoWebhookUrl || getDefaultPMOWebhookUrl());
      mainSection.addWidget(pmoUrlInput);
      
      // PMO Timeout Setting
      var pmoTimeoutInput = CardService.newTextInput()
        .setFieldName("pmoTimeout")
        .setTitle("PMO Timeout (ms)")
        .setHint("Timeout for PMO webhook calls (default: 10000)")
        .setValue(String(settings.pmoTimeout || 10000));
      mainSection.addWidget(pmoTimeoutInput);
      
      // PMO Retry Attempts
      var pmoRetryInput = CardService.newTextInput()
        .setFieldName("pmoRetryAttempts")
        .setTitle("PMO Retry Attempts")
        .setHint("Retry attempts for undefined responses (default: 2)")
        .setValue(String(settings.pmoRetryAttempts || 2));
      mainSection.addWidget(pmoRetryInput);
      
      // PMO Help Text
      var pmoHelp = CardService.newTextParagraph()
        .setText("📋 PMO Integration (Mandatory):\n• All attachments saved to PMO-managed project folders\n• PMO webhook automatically creates folders if they don't exist\n• Retry logic handles folder creation timing issues\n• Attachment save fails only if PMO webhook is unreachable");
      mainSection.addWidget(pmoHelp);
      
      // Save button
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("💾 Save Settings")
          .setOnClickAction(CardService.newAction().setFunctionName("saveSettings")));
      mainSection.addWidget(saveButtonSet);
      
      // Test button
      var testButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("🧪 Test Jira Connection")
          .setOnClickAction(CardService.newAction().setFunctionName("testJiraConnection")));
      mainSection.addWidget(testButtonSet);
      
      // PMO Test button
      var testPMOButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("🧪 Test PMO Connection")
          .setOnClickAction(CardService.newAction().setFunctionName("testPMOConnection")));
      mainSection.addWidget(testPMOButtonSet);
      
      // Back button
      var backButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("← Back to Main")
          .setOnClickAction(CardService.newAction().setFunctionName("backToMain")));
      mainSection.addWidget(backButtonSet);
      
      card.addSection(mainSection);
      
      var builtCard = card.build();
      console.log("Settings card built successfully");
      
      // Try with updateCard instead of pushCard
      var navigation = CardService.newNavigation().updateCard(builtCard);
      console.log("Navigation created successfully");
      
      var response = CardService.newActionResponseBuilder()
        .setNavigation(navigation)
        .build();
      
      console.log("Action response built successfully");
      return response;
      
    } catch (error) {
      console.error("Error in showSettings:", error);
      
      // Return a simple notification instead of navigation
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("Error opening settings: " + error.message))
        .build();
    }
  }
  
  function saveSettings(e) {
    console.log("=== SAVE SETTINGS ===");
    console.log("Form input:", JSON.stringify(e.formInput));
    
    try {
      var jiraUrl = e.formInput.jiraUrl;
      var jiraToken = e.formInput.jiraToken;
      var customJql = e.formInput.customJql;
      
      // Get existing settings to check for current token
      var existingSettings = getUserSettings();
      
      // Handle masked token - if user didn't change it, keep the existing token
      var finalToken = jiraToken;
      if (existingSettings.jiraToken && jiraToken && jiraToken.includes("****")) {
        // User left the masked token unchanged, keep the existing token
        finalToken = existingSettings.jiraToken;
      }
      
      if (!jiraUrl || (!finalToken && !existingSettings.jiraToken)) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please provide both Jira URL and API token"))
          .build();
      }
      
      // Clean up URL (remove trailing slash)
      if (jiraUrl.endsWith('/')) {
        jiraUrl = jiraUrl.slice(0, -1);
      }
      
      // Validate URL format
      if (!jiraUrl.startsWith('http')) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please provide a valid URL starting with http"))
          .build();
      }
      
      // Get PMO settings from form input
      var pmoWebhookUrl = e.formInput.pmoWebhookUrl || getDefaultPMOWebhookUrl();
      var pmoTimeout = parseInt(e.formInput.pmoTimeout) || 10000;
      var pmoRetryAttempts = parseInt(e.formInput.pmoRetryAttempts) || 2;
      
      // Validate PMO webhook URL
      if (!pmoWebhookUrl || !pmoWebhookUrl.startsWith('http')) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please provide a valid PMO webhook URL starting with http"))
          .build();
      }
      
      // Save settings including PMO configuration
      var settings = {
        jiraUrl: jiraUrl,
        jiraToken: finalToken,
        customJql: customJql || getDefaultJQL(),
        pmoWebhookUrl: pmoWebhookUrl,
        pmoTimeout: pmoTimeout,
        pmoRetryAttempts: pmoRetryAttempts,
        savedAt: new Date().toISOString()
      };
      
      setUserSettings(settings);
      
      console.log("Settings saved successfully");
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("✅ Settings saved successfully! Use 'Test Connection' to verify."))
        .build();
        
    } catch (error) {
      console.error("Error saving settings:", error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Error saving settings: " + error.message))
        .build();
    }
  }
  
  function testJiraConnection(e) {
    console.log("=== TEST JIRA CONNECTION ===");
    
    try {
      var settings = getUserSettings();
      
      if (!settings.jiraUrl || !settings.jiraToken) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Please save your settings first"))
          .build();
      }
      
      // Test the connection
      var testResult = testJiraAPI(settings);
      
      if (testResult.success) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("✅ Connection successful! Logged in as: " + testResult.userInfo))
          .build();
      } else {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Connection failed: " + testResult.error))
          .build();
      }
      
    } catch (error) {
      console.error("Error testing connection:", error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Test failed: " + error.message))
        .build();
    }
  }
  
  function testPMOConnection(e) {
    console.log("=== TEST PMO CONNECTION ===");
    
    try {
      var settings = getUserSettings();
      var webhookUrl = settings.pmoWebhookUrl || getDefaultPMOWebhookUrl();
      
      if (!webhookUrl) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ PMO webhook URL not configured. Please set URL first."))
          .build();
      }
      
      // Test with a known ticket pattern (using CXPRODELIVERY format)
      var testTicket = "CXPRODELIVERY-TEST";
      console.log("Testing PMO connection with test ticket:", testTicket);
      
      var result = getPMOProjectFolder(testTicket);
      
      if (result.success) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("✅ PMO webhook connection successful! Found folder ID: " + result.folderId + (result.created ? " (folder created)" : " (existing folder)")))
          .build();
      } else {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ PMO webhook test failed: " + result.error))
          .build();
      }
      
    } catch (error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ PMO connection test failed: " + error.message))
        .build();
    }
  }
  
  function backToMain(e) {
    console.log("=== BACK TO MAIN ===");
    
    try {
      var cards = buildAddOn(e);
      
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(cards[0]))
        .build();
        
    } catch (error) {
      console.error("Error in backToMain:", error);
      
      // Fallback: just pop the current card
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popCard())
        .build();
    }
  }
  
  // ===== DYNAMIC TICKET DETAILS FUNCTION =====
  
  function showTicketDetails(e) {
    console.log("=== SHOW TICKET DETAILS CALLED ===");
    console.log("Form input:", JSON.stringify(e.formInput));
    console.log("Parameters:", JSON.stringify(e.parameters));
    
    try {
      // Get threadId first (needed for dropdown onChange actions)
      var threadId = "";
      if (e.parameters && e.parameters.threadId) {
        threadId = e.parameters.threadId;
        console.log("Got threadId from parameters:", threadId);
      } else if (e.gmail && e.gmail.threadId) {
        threadId = e.gmail.threadId;
        console.log("Got threadId from gmail:", threadId);
      } else {
        console.log("No threadId found in parameters or gmail context");
      }
      
  
      
      var selectedTicketKey = null;
      
      // Try to get selected ticket from form input
      if (e.formInput && e.formInput.selectedTicket) {
        selectedTicketKey = e.formInput.selectedTicket;
        console.log("Got ticket from formInput:", selectedTicketKey);
      } else if (e.parameters && e.parameters.selectedTicket) {
        selectedTicketKey = e.parameters.selectedTicket;
        console.log("Got ticket from parameters:", selectedTicketKey);
      } else {
        console.log("No selected ticket found in formInput or parameters");
      }
      
      console.log("Final selected ticket key:", selectedTicketKey);
      
      if (!selectedTicketKey || selectedTicketKey === "manual") {
        console.log("Manual entry selected or no ticket selected, rebuilding main card");
        var cards = buildAddOn(e);
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().updateCard(cards[0]))
          .build();
      }
      
      // Get the selected ticket details
      console.log("Getting Jira projects...");
      var projectResult = getMyJiraProjects();
      console.log("Got", projectResult.issues ? projectResult.issues.length : 0, "issues");
      var selectedTicket = null;
      
      for (var i = 0; i < projectResult.issues.length; i++) {
        if (projectResult.issues[i].key === selectedTicketKey) {
          selectedTicket = projectResult.issues[i];
          break;
        }
      }
      
      if (!selectedTicket) {
        console.log("Selected ticket not found:", selectedTicketKey);
        console.log("Available tickets:", projectResult.issues.map(function(issue) { return issue.key; }));
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().popCard())
          .build();
      }
      
      console.log("Found selected ticket:", selectedTicket.key);
      
      // Build card with ticket details
      console.log("Building card...");
      var card = CardService.newCardBuilder();
      
      // Compact header with just settings button
      console.log("Creating header section...");
      var headerSection = CardService.newCardSection();
      
      // Compact settings button (gear only)
      var settingsButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("⚙️")
          .setOnClickAction(CardService.newAction().setFunctionName("showSettings")));
      headerSection.addWidget(settingsButtonSet);
      
      card.addSection(headerSection);
      console.log("Header section added");
      
      // Main section with dropdown
      console.log("Creating main section...");
      var mainSection = CardService.newCardSection();
      
      // Recreate dropdown with current selection
      console.log("Creating ticket dropdown...");
      var ticketDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName("selectedTicket")
        .setTitle("Select Active Ticket")
        .setOnChangeAction(CardService.newAction()
          .setFunctionName("showTicketDetails")
          .setParameters({threadId: threadId}));
      
      // Add manual entry option
      console.log("Adding manual entry option...");
      ticketDropdown.addItem("Manual Entry - Enter ticket number below", "manual", false);
      
      // Add all tickets with current one selected
      console.log("Sorting and adding", projectResult.issues.length, "tickets to dropdown...");
      projectResult.issues.sort(function(a, b) {
        return a.key.localeCompare(b.key);
      });
      
      for (var i = 0; i < projectResult.issues.length; i++) {
        try {
          var issue = projectResult.issues[i];
          console.log("Processing ticket", i + 1, "of", projectResult.issues.length, ":", issue.key);
          var displayText = formatCompactTicketDisplay(issue);
          var isSelected = (issue.key === selectedTicketKey);
          ticketDropdown.addItem(displayText, issue.key, isSelected);
          console.log("Successfully added ticket:", issue.key);
        } catch (itemError) {
          console.error("Error processing ticket", issue ? issue.key : "unknown", ":", itemError);
          // Continue with other tickets
        }
      }
      console.log("Finished adding tickets to dropdown");
      
      console.log("Adding dropdown to main section...");
      mainSection.addWidget(ticketDropdown);
      console.log("Dropdown added to main section");
      
  
      
      // Add manual ticket input (pre-filled with selected ticket number)
      console.log("Adding manual ticket input...");
      var ticketNumber = selectedTicketKey.replace('CXPRODELIVERY-', '');
      var ticketInput = CardService.newTextInput()
        .setFieldName("manualTicketNumber")
        .setTitle("Manual Ticket Entry")
        .setHint("e.g., 6500 (will become CXPRODELIVERY-6500)")
        .setValue(ticketNumber);
      mainSection.addWidget(ticketInput);
      console.log("Manual ticket input added");
      
      // Create ticket details section
      console.log("Creating ticket details section...");
      var detailsSection = CardService.newCardSection()
        .setHeader("📋 Selected Ticket Details");
      
      var statusEmoji = getStatusEmoji(selectedTicket.status);
      var detailsText = statusEmoji + " **" + selectedTicket.key + "**\n\n";
      detailsText += "📝 **Summary:**\n" + selectedTicket.summary + "\n\n";
      detailsText += "📊 **Status:** " + selectedTicket.status + "\n";
      detailsText += "🏷️ **Type:** " + selectedTicket.issueType;
      
      var detailsParagraph = CardService.newTextParagraph().setText(detailsText);
      detailsSection.addWidget(detailsParagraph);
      console.log("Ticket details section created");
      
      // Handle Gmail context for attachments (threadId already determined above)
      
      if (threadId) {
        console.log("Processing attachments for threadId:", threadId);
        try {
          var thread = GmailApp.getThreadById(threadId);
          var messages = thread.getMessages();
          var allAttachments = [];
          console.log("Found", messages.length, "messages in thread");
          
          for (var i = 0; i < messages.length; i++) {
            var message = messages[i];
            var attachments = message.getAttachments();
            console.log("Message", i + 1, "has", attachments.length, "attachments");
            
            for (var j = 0; j < attachments.length; j++) {
              var attachment = attachments[j];
              allAttachments.push({
                name: attachment.getName(),
                size: attachment.getSize(),
                attachment: attachment,
                messageIndex: i
              });
            }
          }
          
          console.log("Total attachments found:", allAttachments.length);
          
          if (allAttachments.length > 0) {
            console.log("Creating attachment section...");
            
            // First, check if we have current form selections to preserve
            var currentSelections = {};
            var hasCurrentSelections = false;
            
            if (e.formInput) {
              console.log("=== CHECKING CURRENT FORM SELECTIONS ===");
              for (var i = 0; i < allAttachments.length; i++) {
                var formFieldName = "attachment_" + i;
                // If field exists in form input → selected = true, if missing → selected = false
                var isCurrentlySelected = e.formInput.hasOwnProperty(formFieldName) && e.formInput[formFieldName] && e.formInput[formFieldName].length > 0;
                
                // Use unique key to handle duplicate filenames (filename + index)
                var uniqueKey = allAttachments[i].name + "_" + i;
                currentSelections[uniqueKey] = isCurrentlySelected;
                hasCurrentSelections = true;
                console.log("Current form state for", allAttachments[i].name, "- field:", formFieldName, "- exists:", e.formInput.hasOwnProperty(formFieldName), "- selected:", isCurrentlySelected, "- stored as:", uniqueKey);
              }
            }
            
            // If we have current selections, store them for future use
            if (hasCurrentSelections) {
              console.log("Storing current form selections for future use");
              storeAttachmentSelections(threadId, currentSelections);
            }
            
            // Get stored attachment selections for this thread (might be what we just stored)
            console.log("Getting stored attachment selections...");
            var storedSelections = getStoredAttachmentSelections(threadId);
            console.log("Stored selections:", storedSelections);
            
            // Group attachments by file extension
            var attachmentsByExtension = {};
            var extensionCounts = {};
            
            for (var i = 0; i < allAttachments.length; i++) {
              var attachment = allAttachments[i];
              var fileName = attachment.name;
              var extension = fileName.lastIndexOf('.') > -1 ? 
                fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              
              if (!attachmentsByExtension[extension]) {
                attachmentsByExtension[extension] = [];
                extensionCounts[extension] = 0;
              }
              
              attachmentsByExtension[extension].push({
                attachment: attachment,
                index: i
              });
              extensionCounts[extension]++;
            }
            
            // Sort extensions alphabetically for consistent display
            var sortedExtensions = Object.keys(attachmentsByExtension).sort();
            console.log("Attachment extensions found:", sortedExtensions);
            
            // Create a section for each file extension
            for (var extIndex = 0; extIndex < sortedExtensions.length; extIndex++) {
              var extension = sortedExtensions[extIndex];
              var attachmentsInGroup = attachmentsByExtension[extension];
              var count = extensionCounts[extension];
              
              console.log("Processing extension group:", extension, "with", count, "files");
              
              // Create section header with extension and count (show total in first section)
              var extDisplayName = extension === 'no-extension' ? 'Files without extension' : 
                extension.toUpperCase() + ' files';
              var headerText = extDisplayName + " (" + count + ")";
              if (extIndex === 0) {
                headerText = "📎 Select Attachments: " + headerText + " (Total: " + allAttachments.length + ")";
              }
              
              var extensionSection = CardService.newCardSection()
                .setHeader(headerText)
                .setCollapsible(true)
                .setNumUncollapsibleWidgets(count > 5 ? 2 : count); // Show first 2 if more than 5 files
              
              // Add checkboxes for each attachment in this extension group
              for (var j = 0; j < attachmentsInGroup.length; j++) {
                var attachmentData = attachmentsInGroup[j];
                var attachment = attachmentData.attachment;
                var originalIndex = attachmentData.index;
                
                console.log("Processing attachment", originalIndex, ":", attachment.name);
                
                // Check if this attachment was previously selected
                var isSelected = false; // Default to false - let user explicitly choose
                
                // Use unique key to handle duplicate filenames (filename + index)
                var uniqueKey = attachment.name + "_" + originalIndex;
                
                if (storedSelections && storedSelections.hasOwnProperty(uniqueKey)) {
                  isSelected = storedSelections[uniqueKey];
                  console.log("Found stored selection for", attachment.name, "(", uniqueKey, "):", isSelected);
                } else {
                  console.log("No stored selection found for", attachment.name, "(", uniqueKey, ") - using default (unselected):", isSelected);
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(attachment.name + " (" + formatFileSize(attachment.size) + ")", originalIndex.toString(), isSelected);
                
                extensionSection.addWidget(checkbox);
              }
              
              console.log("Adding", extension, "section to card");
              card.addSection(extensionSection);
            }
            
            console.log("All attachment sections added to card");
          } else {
            console.log("No attachments found, skipping attachment section");
          }
        } catch (attachmentError) {
          console.error("Error loading attachments in showTicketDetails:", attachmentError);
        }
      } else {
        console.log("No threadId, skipping attachment processing");
      }
      
      // Add save button to main section
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("Save to PMO Folder")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("saveSelectedAttachmentsToGDrive")
            .setParameters({threadId: threadId})));
      mainSection.addWidget(saveButtonSet);
      
      // Add sections to card in the correct order
      console.log("Adding main section to card...");
      card.addSection(mainSection);
      console.log("Adding details section to card...");
      card.addSection(detailsSection);
      
      console.log("All sections added. Built card with ticket details for:", selectedTicket.key);
      
      var builtCard = card.build();
      console.log("Card built successfully");
      
      var navigation = CardService.newNavigation().updateCard(builtCard);
      console.log("Navigation created");
      
      var response = CardService.newActionResponseBuilder()
        .setNavigation(navigation)
        .build();
      console.log("Response built, returning...");
      
      return response;
        
    } catch (error) {
      console.error("Error in showTicketDetails:", error);
      console.error("Error stack:", error.stack);
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popCard())
        .build();
    }
  }
  
  // ===== SAVE FUNCTION =====
  
  function saveSelectedAttachmentsToGDrive(e) {
    try {
      console.log("=== SAVE FUNCTION CALLED ===");
      console.log("Form input:", JSON.stringify(e.formInput));
      console.log("Parameters:", JSON.stringify(e.parameters));
      
      var selectedTicket = e.formInput.selectedTicket;
      var manualTicketNumber = e.formInput.manualTicketNumber;
      var fullTicket = e.formInput.jiraTicket;
      var threadId = e.parameters ? e.parameters.threadId : "";
      
      // Determine final ticket
      var finalTicket;
      if (selectedTicket && selectedTicket !== "manual") {
        finalTicket = selectedTicket;
        console.log("Using selected ticket:", finalTicket);
      } else if (manualTicketNumber) {
        finalTicket = "CXPRODELIVERY-" + manualTicketNumber;
        console.log("Using manual ticket:", finalTicket);
      } else if (fullTicket) {
        finalTicket = fullTicket;
        console.log("Using full ticket:", finalTicket);
      } else {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please select a ticket or enter a manual ticket number"))
          .build();
      }
      
      if (!threadId) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please open an email to save attachments"))
          .build();
      }
      
      console.log("Loading attachments from thread:", threadId);
      
      // Get all attachments from the thread
      var thread = GmailApp.getThreadById(threadId);
      var messages = thread.getMessages();
      var allAttachments = [];
      
      for (var i = 0; i < messages.length; i++) {
        var message = messages[i];
        var attachments = message.getAttachments();
        
        for (var j = 0; j < attachments.length; j++) {
          var attachment = attachments[j];
          allAttachments.push({
            name: attachment.getName(),
            size: attachment.getSize(),
            attachment: attachment,
            messageIndex: i
          });
        }
      }
      
      console.log("Found", allAttachments.length, "total attachments");
      
      if (allAttachments.length === 0) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("No attachments found in this thread"))
          .build();
      }
      
      // Find selected attachments from form input and store selection state
      var selectedAttachments = [];
      var selectionState = {};
      
      console.log("=== PROCESSING ATTACHMENT SELECTIONS FOR SAVE ===");
      console.log("ThreadId for storage:", threadId);
      console.log("Total attachments to process:", allAttachments.length);
      console.log("Form input keys:", Object.keys(e.formInput));
      
      for (var i = 0; i < allAttachments.length; i++) {
        var checkboxValue = e.formInput["attachment_" + i];
        var isSelected = checkboxValue && checkboxValue.length > 0;
        var attachmentName = allAttachments[i].name;
        
        console.log("Attachment", i, "(", attachmentName, ") - checkbox value:", checkboxValue, "- selected:", isSelected);
        
        // Store selection state for this attachment using unique key
        var uniqueKey = attachmentName + "_" + i;
        selectionState[uniqueKey] = isSelected;
        
        if (isSelected) {
          selectedAttachments.push(allAttachments[i]);
          console.log("Added to selected list:", attachmentName);
        }
      }
      
      console.log("Final selection state to store:", JSON.stringify(selectionState));
      
      // Store the current selection state for this thread
      storeAttachmentSelections(threadId, selectionState);
      
      console.log("Selected", selectedAttachments.length, "attachments for saving");
      
      if (selectedAttachments.length === 0) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("Please select at least one attachment"))
          .build();
      }
      
      // PMO Integration Logic (ONLY METHOD) - Lookup/Create folder
      console.log("PMO lookup/create for:", finalTicket);
      
      var pmoResult = getPMOProjectFolder(finalTicket);
      
      if (!pmoResult.success) {
        // PMO webhook failed - this is an error condition
        console.error("PMO lookup failed:", pmoResult.error);
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ PMO folder lookup/creation failed for " + finalTicket + ": " + pmoResult.error))
          .build();
      }
      
      console.log("PMO returned folder ID:", pmoResult.folderId, "- Created:", pmoResult.created || false);
      
      var folderResult = getFolderByIdSafely(pmoResult.folderId);
      
      if (!folderResult.success) {
        // Folder access failed - this is an error condition  
        console.error("PMO folder access failed:", folderResult.error);
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText("❌ Cannot access PMO project folder (ID: " + pmoResult.folderId + "): " + folderResult.error))
          .build();
      }
      
      var ticketFolder = folderResult.folder;
      console.log("Using PMO folder:", folderResult.name);
      
      // Save selected attachments with duplicate handling
      var savedCount = 0;
      var skippedCount = 0;
      var savedFiles = [];
      var skippedFiles = [];
      
      for (var i = 0; i < selectedAttachments.length; i++) {
        try {
          var attachmentData = selectedAttachments[i];
          var fileName = attachmentData.name;
          console.log("Processing attachment:", fileName);
          
          // Check if file already exists
          var existingFiles = ticketFolder.getFilesByName(fileName);
          
          if (existingFiles.hasNext()) {
            // File already exists
            var existingFile = existingFiles.next();
            var existingSize = existingFile.getSize();
            var newSize = attachmentData.size;
            
            console.log("File exists:", fileName, "- existing size:", existingSize, "new size:", newSize);
            
            if (existingSize === newSize) {
              // Same size, likely duplicate - skip
              console.log("Skipping duplicate file:", fileName);
              skippedCount++;
              skippedFiles.push(fileName);
              continue;
            } else {
              // Different size, create with timestamp
              var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
              var fileExtension = fileName.lastIndexOf('.') > -1 ? fileName.substring(fileName.lastIndexOf('.')) : '';
              var baseName = fileName.lastIndexOf('.') > -1 ? fileName.substring(0, fileName.lastIndexOf('.')) : fileName;
              var newFileName = baseName + "_" + timestamp + fileExtension;
              
              console.log("Creating file with new name:", newFileName);
              var file = ticketFolder.createFile(attachmentData.attachment);
              file.setName(newFileName);
              savedCount++;
              savedFiles.push(newFileName);
            }
          } else {
            // File doesn't exist, save normally
            console.log("Saving new file:", fileName);
            var file = ticketFolder.createFile(attachmentData.attachment);
            savedCount++;
            savedFiles.push(file.getName());
          }
          
          console.log("Processed:", fileName);
        } catch (saveError) {
          console.error("Error saving attachment", attachmentData.name, ":", saveError);
        }
      }
      
      console.log("=== SAVE COMPLETE ===");
      console.log("Saved", savedCount, "files to", finalTicket);
      console.log("Skipped", skippedCount, "duplicate files");
      console.log("Files saved:", savedFiles);
      console.log("Files skipped:", skippedFiles);
      
      // Create notification message with PMO info
      var notificationText;
      if (pmoResult.created) {
        notificationText = "✅ Created PMO folder and saved " + savedCount + " attachments: " + folderResult.name;
      } else {
        notificationText = "✅ Saved " + savedCount + " attachments to PMO project folder: " + folderResult.name;
      }
      
      if (skippedCount > 0) {
        notificationText += " (⚠️ " + skippedCount + " duplicates skipped)";
      }
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(notificationText))
        .build();
        
    } catch (error) {
      console.error("=== SAVE ERROR ===");
      console.error("Error:", error);
      console.error("Stack:", error.stack);
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Error: " + error.message))
        .build();
    }
  }
  
  // ===== SETTINGS STORAGE FUNCTIONS =====
  
  function getUserSettings() {
    try {
      var userProperties = PropertiesService.getUserProperties();
      var settingsJson = userProperties.getProperty('JIRA_SETTINGS');
      
      if (settingsJson) {
        var settings = JSON.parse(settingsJson);
        
        // Add PMO defaults for existing users who don't have PMO settings
        if (!settings.pmoWebhookUrl) {
          settings.pmoWebhookUrl = getDefaultPMOWebhookUrl();
        }
        if (!settings.pmoTimeout) {
          settings.pmoTimeout = 10000;
        }
        if (!settings.pmoRetryAttempts) {
          settings.pmoRetryAttempts = 2;
        }
        
        return settings;
      } else {
        // Return defaults for new users
        return {
          pmoWebhookUrl: getDefaultPMOWebhookUrl(),
          pmoTimeout: 10000,
          pmoRetryAttempts: 2
        };
      }
    } catch (error) {
      console.error("Error getting user settings:", error);
      return {
        pmoWebhookUrl: getDefaultPMOWebhookUrl(),
        pmoTimeout: 10000,
        pmoRetryAttempts: 2
      };
    }
  }
  
  function setUserSettings(settings) {
    try {
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('JIRA_SETTINGS', JSON.stringify(settings));
      console.log("Settings saved to user properties");
    } catch (error) {
      console.error("Error setting user settings:", error);
      throw error;
    }
  }
  
  function getDefaultJQL() {
    return 'project = CXPRODELIVERY AND issuetype in (Project, "Project (Standard Solution)") AND status in (HYPERCARE, "Order received", "Test system available", "Project go-live/productive start", "System Design Assigned", "Implementation Assigned", "Implementation Order Assigned", "Test System Available (Implementation Order)", "Handover Check Needed", "HYPERCARE (WITH CHECK)", "LIVE SYSTEM AVAILABLE", "System Design Order Received", "System Design Started", "Requirements Clarified", "Implementation Started") AND "Technical Project Manager" in (currentUser())';
  }
  
  function maskToken(token) {
    if (!token || token.length < 8) return "****";
    return token.substring(0, 4) + "****" + token.substring(token.length - 4);
  }
  
  function testJiraAPI(settings) {
    try {
      var testUrl = settings.jiraUrl + '/rest/api/2/myself';
      
      var response = UrlFetchApp.fetch(testUrl, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + settings.jiraToken,
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      });
      
      var responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        var data = JSON.parse(response.getContentText());
        return {
          success: true,
          userInfo: data.displayName || data.name || 'User',
          email: data.emailAddress || 'Unknown'
        };
      } else {
        var errorData = response.getContentText();
        return {
          success: false,
          error: "HTTP " + responseCode + ": " + errorData
        };
      }
      
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  // ===== PMO INTEGRATION FUNCTIONS =====
  
  function getPMOProjectFolder(ticketKey) {
    try {
      var settings = getUserSettings();
      var webhookUrl = settings.pmoWebhookUrl || getDefaultPMOWebhookUrl();
      var maxRetries = settings.pmoRetryAttempts || 2;
      var timeout = settings.pmoTimeout || 10000;
      
      console.log("PMO webhook lookup/create for ticket:", ticketKey);
      console.log("Using webhook URL:", webhookUrl);
      
      for (var attempt = 1; attempt <= maxRetries; attempt++) {
        console.log("PMO request attempt", attempt, "of", maxRetries);
        
        try {
          var response = UrlFetchApp.fetch(webhookUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            payload: JSON.stringify({"text": ticketKey}),
            muteHttpExceptions: true
          });
          
          console.log("PMO webhook response code:", response.getResponseCode());
          
          if (response.getResponseCode() === 200) {
            var responseText = response.getContentText();
            console.log("PMO webhook response:", responseText);
            
            var data = JSON.parse(responseText);
            
            if (data && data.length > 0) {
              var folderid = data[0].folderid;
              console.log("PMO returned folderid:", folderid, "- type:", typeof folderid);
              
              if (folderid && folderid !== 'undefined' && folderid.toString().trim() !== '') {
                console.log("PMO webhook returned valid folder ID:", folderid);
                return {
                  success: true,
                  folderId: folderid,
                  created: attempt > 1 // True if folder was created on retry
                };
              } else {
                console.log("PMO webhook returned undefined/empty folder ID on attempt", attempt);
                
                if (attempt === maxRetries) {
                  return {
                    success: false,
                    error: 'PMO webhook returned undefined folder ID after ' + maxRetries + ' attempts for ticket: ' + ticketKey
                  };
                }
                
                // Wait briefly before retry (folder creation might need time)
                console.log("Waiting 1 second before retry...");
                Utilities.sleep(1000);
                continue;
              }
            } else {
              return {
                success: false,
                error: 'PMO webhook returned invalid response format for ticket: ' + ticketKey
              };
            }
          } else {
            return {
              success: false,
              error: 'PMO webhook HTTP error: ' + response.getResponseCode() + ' - ' + response.getContentText()
            };
          }
        } catch (requestError) {
          if (attempt === maxRetries) {
            return {
              success: false,
              error: 'PMO webhook network error after ' + maxRetries + ' attempts: ' + requestError.message
            };
          }
          
          console.log("Request error on attempt", attempt, ":", requestError.message);
          Utilities.sleep(1000);
          continue;
        }
      }
      
    } catch (error) {
      return {
        success: false,
        error: 'PMO webhook error: ' + error.message
      };
    }
  }
  
  function getFolderByIdSafely(folderId) {
    try {
      console.log("Attempting to access folder with ID:", folderId);
      
      var folder = DriveApp.getFolderById(folderId);
      
      // Test folder access by trying to get name
      var folderName = folder.getName();
      console.log("Successfully accessed folder:", folderName);
      
      return {
        success: true,
        folder: folder,
        name: folderName
      };
    } catch (error) {
      console.error("Cannot access folder with ID", folderId, ":", error.message);
      return {
        success: false,
        error: 'Cannot access folder (ID: ' + folderId + '): ' + error.message
      };
    }
  }
  
  function getDefaultPMOWebhookUrl() {
    return 'https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299';
  }
  
  // ===== HELPER FUNCTIONS =====
  
  function formatCompactTicketDisplay(issue) {
    var summary = issue.summary || 'No summary';
    var statusEmoji = getStatusEmoji(issue.status);
    
    // Extract client info from summary
    var clientMatch = summary.match(/^(\d+)\s*-\s*([^|]+)/);
    
    if (clientMatch && clientMatch[2]) {
      var clientName = clientMatch[2].trim();
      
      // Shorten client name to fit dropdown (keep it reasonable for dropdown width)
      if (clientName.length > 20) {
        clientName = clientName.substring(0, 17) + "...";
      }
      
      // Format: [Emoji] [Full Ticket Key] - [Client Name]
      return statusEmoji + " " + issue.key + " - " + clientName;
      
    } else {
      // Fallback: just show ticket key with emoji
      return statusEmoji + " " + issue.key;
    }
  }
  
  function getStatusEmoji(status) {
    var statusEmojis = {
      'HYPERCARE': '🔧',
      'HYPERCARE (WITH CHECK)': '🔍',
      'Order received': '📋',
      'Test system available': '🧪',
      'Test System Available (Implementation Order)': '🧪',
      'Project go-live/productive start': '🚀',
      'System Design Assigned': '👤',          // User assigned to design
      'System Design Order Received': '📨',
      'System Design Started': '🎨',
      'Implementation Assigned': '👨‍💻',        // User assigned to implement
      'Implementation Order Assigned': '🧑‍💼',  // User assigned to manage order
      'Implementation Started': '🔨',
      'Requirements Clarified': '✅',
      'Handover Check Needed': '🔍',
      'LIVE SYSTEM AVAILABLE': '🟢'
    };
    
    return statusEmojis[status] || '📝';
  }
  
  function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    var k = 1024;
    var sizes = ['B', 'KB', 'MB', 'GB'];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }
  
  function getOrCreateFolder(folderName, parentFolder) {
    var parent = parentFolder || DriveApp.getRootFolder();
    var folders = parent.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return parent.createFolder(folderName);
    }
  }
  
  // ===== JIRA INTEGRATION (Updated to use user settings) =====
  
  function getMyJiraProjects() {
    try {
      var settings = getUserSettings();
      
      if (!settings.jiraUrl || !settings.jiraToken) {
        console.error('Jira settings not configured');
        return { projects: {}, issues: [], totalIssues: 0 };
      }
      
      var searchUrl = settings.jiraUrl + '/rest/api/2/search';
      var jqlQuery = settings.customJql || getDefaultJQL();
      
      var payload = {
        jql: jqlQuery,
        fields: ['project', 'key', 'summary', 'status', 'issuetype'],
        maxResults: 1000,
        startAt: 0
      };
      
      console.log('Using JQL filter:', jqlQuery);
      
      var response = UrlFetchApp.fetch(searchUrl, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + settings.jiraToken,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.error('Jira API failed:', response.getResponseCode(), response.getContentText());
        return { projects: {}, issues: [], totalIssues: 0 };
      }
      
      var data = JSON.parse(response.getContentText());
      var projects = {};
      var issues = [];
      
      console.log('JQL Filter returned', data.total || 0, 'issues');
      
      if (data.issues && data.issues.length > 0) {
        data.issues.forEach(function(issue) {
          var project = issue.fields.project;
          if (project) {
            projects[project.key] = {
              key: project.key,
              name: project.name,
              id: project.id
            };
            
            issues.push({
              key: issue.key,
              summary: issue.fields.summary || 'No summary',
              status: issue.fields.status ? issue.fields.status.name : 'Unknown',
              issueType: issue.fields.issuetype ? issue.fields.issuetype.name : 'Unknown',
              projectKey: project.key,
              projectName: project.name
            });
          }
        });
      }
      
      console.log('Projects found:', Object.keys(projects));
      console.log('Issues found:', issues.length);
      
      return { 
        projects: projects, 
        issues: issues,
        totalIssues: data.total || 0 
      };
      
    } catch (error) {
      console.error('Error getting Jira projects:', error);
      return { projects: {}, issues: [], totalIssues: 0 };
    }
  }
  
  // ===== ATTACHMENT SELECTION MEMORY FUNCTIONS =====
  
  function storeAttachmentSelections(threadId, selectionState) {
    try {
      if (!threadId) {
        console.log("No threadId provided, skipping attachment selection storage");
        return;
      }
      
      var userProperties = PropertiesService.getUserProperties();
      var key = 'ATTACHMENT_SELECTIONS_' + threadId;
      
      console.log("=== STORING ATTACHMENT SELECTIONS ===");
      console.log("ThreadId:", threadId);
      console.log("Storage key:", key);
      console.log("Selection state to store:", JSON.stringify(selectionState));
      
      userProperties.setProperty(key, JSON.stringify(selectionState));
      
      // Verify storage worked
      var stored = userProperties.getProperty(key);
      console.log("Verification - stored value:", stored);
      console.log("Storage successful:", stored !== null);
      
    } catch (error) {
      console.error("Error storing attachment selections:", error);
    }
  }
  
  function getStoredAttachmentSelections(threadId) {
    try {
      if (!threadId) {
        console.log("No threadId provided, returning null");
        return null;
      }
      
      var userProperties = PropertiesService.getUserProperties();
      var key = 'ATTACHMENT_SELECTIONS_' + threadId;
      
      console.log("=== RETRIEVING ATTACHMENT SELECTIONS ===");
      console.log("ThreadId:", threadId);
      console.log("Storage key:", key);
      
      var storedSelectionsJson = userProperties.getProperty(key);
      console.log("Raw stored value:", storedSelectionsJson);
      
      if (storedSelectionsJson) {
        var selections = JSON.parse(storedSelectionsJson);
        console.log("Parsed selections:", selections);
        console.log("Selection keys:", Object.keys(selections));
        return selections;
      } else {
        console.log("No stored attachment selections found for thread:", threadId);
        
        // Debug: List all stored keys to see what's actually stored
        var allProperties = userProperties.getProperties();
        var attachmentKeys = [];
        for (var prop in allProperties) {
          if (prop.startsWith('ATTACHMENT_SELECTIONS_')) {
            attachmentKeys.push(prop);
          }
        }
        console.log("All attachment selection keys found:", attachmentKeys);
        
        return null;
      }
      
    } catch (error) {
      console.error("Error retrieving attachment selections:", error);
      return null;
    }
  }