// ===== SCRIPT PROPERTIES SETUP =====

/**
 * Setup Script Properties for secure credential storage
 * Run this function once to configure authentication credentials
 */
function setupScriptProperties() {
  console.log("=== SETTING UP SCRIPT PROPERTIES ===");
  
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    
    // Set the required environment variables for 2-stage OAuth authentication
    scriptProperties.setProperties({
       
      // Core service URLs
      'TRIMBLE_AUTH_SERVER_TOKEN': 'your_auth_server_token_here',  // Add your Trimble auth server token
      'PMO_WEBHOOK_URL': 'https://flows.stage.trimble-ai.com/agentic/workflows/v1/webhook-test/55633332-0344-4811-8f3a-46e6be725a42',
      'TRIMBLE_OAUTH_URL': 'https://stage.id.trimblecloud.com/oauth/token',
      'DEFAULT_JIRA_URL': 'https://support.transporeon.com',
      
      // OAuth configuration
      'OAUTH_GRANT_TYPE': 'client_credentials',
      'OAUTH_SCOPE': 'Agentic-N8N-Webhook',
      
      // Storage configuration
      'SETTINGS_STORAGE_KEY': 'JIRA_SETTINGS'
    });
    
    console.log("✅ Script properties configured successfully:");
    console.log("  - WEBHOOK_USERNAME: larmax (legacy)");
    console.log("  - WEBHOOK_PASSWORD: ****** (hidden, legacy)");
    console.log("  - TRIMBLE_AUTH_SERVER_TOKEN: ****** (hidden)");
    console.log("  - PMO_WEBHOOK_URL: " + scriptProperties.getProperty('PMO_WEBHOOK_URL'));
    console.log("  - TRIMBLE_OAUTH_URL: " + scriptProperties.getProperty('TRIMBLE_OAUTH_URL'));
    console.log("  - DEFAULT_JIRA_URL: " + scriptProperties.getProperty('DEFAULT_JIRA_URL'));
    console.log("  - OAUTH_GRANT_TYPE: " + scriptProperties.getProperty('OAUTH_GRANT_TYPE'));
    console.log("  - OAUTH_SCOPE: " + scriptProperties.getProperty('OAUTH_SCOPE'));
    console.log("  - SETTINGS_STORAGE_KEY: " + scriptProperties.getProperty('SETTINGS_STORAGE_KEY'));
    
    return true;
    
  } catch (error) {
    console.error("❌ Failed to setup script properties:", error.message);
    return false;
  }
}

/**
 * Get webhook credentials from Script Properties
 */
function getWebhookCredentials() {
  var properties = PropertiesService.getScriptProperties();
  var username = properties.getProperty('WEBHOOK_USERNAME');
  var password = properties.getProperty('WEBHOOK_PASSWORD');
  
  if (!username || !password) {
    throw new Error('Webhook credentials not configured. Please run setupScriptProperties() first.');
  }
  
  return {
    username: username,
    password: password,
    encodedCredentials: Utilities.base64Encode(username + ':' + password)
  };
}

/**
 * Get PMO Webhook URL from Script Properties
 */
function getPMOWebhookURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('PMO_WEBHOOK_URL');
  
  if (!url) {
    throw new Error('PMO_WEBHOOK_URL not configured. Please run setupScriptProperties() first.');
  }
  
  return url;
}

/**
 * Environment Variable Getter Functions
 */

/**
 * Get Trimble Auth Server Token from Script Properties
 */
function getTrimbleAuthServerToken() {
  var properties = PropertiesService.getScriptProperties();
  var token = properties.getProperty('TRIMBLE_AUTH_SERVER_TOKEN');
  
  if (!token || token === 'your_auth_server_token_here') {
    throw new Error('TRIMBLE_AUTH_SERVER_TOKEN not configured. Please set your auth server token in setupScriptProperties().');
  }
  
  return token;
}

/**
 * Get Trimble OAuth URL from Script Properties
 */
function getTrimbleOAuthURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('TRIMBLE_OAUTH_URL');
  
  if (!url) {
    // Fallback to hardcoded value for backward compatibility
    return 'https://stage.id.trimblecloud.com/oauth/token';
  }
  
  return url;
}

/**
 * Get Default Jira URL from Script Properties
 */
function getDefaultJiraURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('DEFAULT_JIRA_URL');
  
  if (!url) {
    // Fallback to hardcoded value for backward compatibility
    return 'https://support.transporeon.com';
  }
  
  return url;
}

/**
 * Get OAuth Grant Type from Script Properties
 */
function getOAuthGrantType() {
  var properties = PropertiesService.getScriptProperties();
  var grantType = properties.getProperty('OAUTH_GRANT_TYPE');
  
  if (!grantType) {
    // Fallback to default value
    return 'client_credentials';
  }
  
  return grantType;
}

/**
 * Get OAuth Scope from Script Properties
 */
function getOAuthScope() {
  var properties = PropertiesService.getScriptProperties();
  var scope = properties.getProperty('OAUTH_SCOPE');
  
  if (!scope) {
    // Fallback to default value
    return 'Agentic-N8N-Webhook';
  }
  
  return scope;
}

/**
 * Get Settings Storage Key from Script Properties
 */
function getSettingsStorageKey() {
  var properties = PropertiesService.getScriptProperties();
  var key = properties.getProperty('SETTINGS_STORAGE_KEY');
  
  if (!key) {
    // Fallback to default value
    return 'JIRA_SETTINGS';
  }
  
  return key;
}

/**
 * Token Cache Management
 */

/**
 * Get cached OAuth token if still valid
 */
function getCachedOAuthToken() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var tokenData = properties.getProperty('CACHED_OAUTH_TOKEN');
    
    if (!tokenData) {
      console.log("No cached OAuth token found");
      return null;
    }
    
    var cached = JSON.parse(tokenData);
    var now = new Date().getTime();
    var expiresAt = cached.expiresAt || 0;
    
    // Add 60 second buffer to avoid using tokens about to expire
    var bufferTime = 60 * 1000; // 60 seconds
    
    if (now < (expiresAt - bufferTime)) {
      var timeLeft = Math.floor((expiresAt - now) / 1000);
      console.log("✅ Using cached OAuth token (expires in " + timeLeft + " seconds)");
      return cached.accessToken;
    } else {
      console.log("❌ Cached OAuth token expired, will request new one");
      // Clear expired token
      properties.deleteProperty('CACHED_OAUTH_TOKEN');
      return null;
    }
    
  } catch (error) {
    console.error("Error checking cached token:", error.message);
    return null;
  }
}

/**
 * Cache OAuth token with expiration time
 */
function cacheOAuthToken(accessToken, expiresIn) {
  try {
    var properties = PropertiesService.getScriptProperties();
    var now = new Date().getTime();
    var expiresAt = now + (expiresIn * 1000); // Convert seconds to milliseconds
    
    var tokenData = {
      accessToken: accessToken,
      expiresIn: expiresIn,
      expiresAt: expiresAt,
      cachedAt: now
    };
    
    properties.setProperty('CACHED_OAUTH_TOKEN', JSON.stringify(tokenData));
    
    var expiresInMinutes = Math.floor(expiresIn / 60);
    console.log("✅ OAuth token cached successfully (expires in " + expiresInMinutes + " minutes)");
    
  } catch (error) {
    console.error("Error caching OAuth token:", error.message);
    // Don't throw error - caching is optional optimization
  }
}

/**
 * Clear cached OAuth token (force renewal)
 */
function clearCachedOAuthToken() {
  try {
    var properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('CACHED_OAUTH_TOKEN');
    console.log("🗑️ Cached OAuth token cleared");
  } catch (error) {
    console.error("Error clearing cached token:", error.message);
  }
}

/**
 * Get OAuth token status information
 */
function getOAuthTokenStatus() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var tokenData = properties.getProperty('CACHED_OAUTH_TOKEN');
    
    if (!tokenData) {
      return {
        cached: false,
        message: "No cached token"
      };
    }
    
    var cached = JSON.parse(tokenData);
    var now = new Date().getTime();
    var expiresAt = cached.expiresAt || 0;
    var timeLeft = Math.floor((expiresAt - now) / 1000);
    
    if (timeLeft > 60) {
      return {
        cached: true,
        valid: true,
        timeLeft: timeLeft,
        message: "Valid token (expires in " + Math.floor(timeLeft / 60) + " minutes)"
      };
    } else if (timeLeft > 0) {
      return {
        cached: true,
        valid: false,
        timeLeft: timeLeft,
        message: "Token expires soon (" + timeLeft + " seconds)"
      };
    } else {
      return {
        cached: true,
        valid: false,
        timeLeft: 0,
        message: "Token expired " + Math.abs(timeLeft) + " seconds ago"
      };
    }
    
  } catch (error) {
    return {
      cached: false,
      error: error.message,
      message: "Error checking token status"
    };
  }
}

/**
 * Handle expired token errors and retry with new token
 */
function handleExpiredTokenError(originalError, retryFunction, retryParams) {
  console.log("🔄 Handling potential expired token error...");
  
  // Check if error indicates expired token
  if (originalError.message.includes('401') || originalError.message.includes('403') || 
      originalError.message.includes('Unauthorized') || originalError.message.includes('Forbidden')) {
    
    console.log("⚠️ Detected authentication error, clearing cached token and retrying...");
    clearCachedOAuthToken();
    
    try {
      // Retry the operation with a fresh token
      console.log("🔄 Retrying operation with fresh OAuth token...");
      return retryFunction.apply(null, retryParams);
    } catch (retryError) {
      console.error("❌ Retry failed:", retryError.message);
      throw new Error('Authentication failed even after token renewal: ' + retryError.message);
    }
  } else {
    // Not an authentication error, re-throw original error
    throw originalError;
  }
}

/**
 * Stage 1: Get OAuth token from Trimble auth server with caching
 * Returns access token for Stage 2 API calls
 * IMPORTANT: Token expires after 3600 seconds (1 hour)
 */
function getTrimbleOAuthToken() {
  console.log("=== STAGE 1: TRIMBLE OAUTH TOKEN RETRIEVAL ===");
  
  try {
    // Check if we have a valid cached token first
    var cachedToken = getCachedOAuthToken();
    if (cachedToken) {
      return cachedToken;
    }
    
    // No valid cached token, request new one
    console.log("Requesting new OAuth token from server...");
    
    var authServerToken = getTrimbleAuthServerToken();
    var authUrl = getTrimbleOAuthURL();
    var grantType = getOAuthGrantType();
    var scope = getOAuthScope();
    
    console.log("OAuth request to:", authUrl);
    console.log("Grant type:", grantType);
    console.log("Scope:", scope);
    
    var payload = 'grant_type=' + encodeURIComponent(grantType) + '&scope=' + encodeURIComponent(scope);
    
    var response = UrlFetchApp.fetch(authUrl, {
      method: 'POST',
      headers: {
        'Authorization': authServerToken,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: payload,
      muteHttpExceptions: true,
      timeout: 10000
    });
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    console.log("OAuth response code:", responseCode);
    console.log("OAuth response:", responseText);
    
    if (responseCode === 200) {
      var data = JSON.parse(responseText);
      if (data.access_token) {
        var expiresIn = data.expires_in || 3600; // Default to 1 hour if not specified
        console.log("✅ OAuth token retrieved successfully (expires in " + expiresIn + " seconds)");
        
        // Cache the token for future use
        cacheOAuthToken(data.access_token, expiresIn);
        
        return data.access_token;
      } else {
        throw new Error('No access_token in response: ' + responseText);
      }
    } else {
      throw new Error('OAuth request failed with code ' + responseCode + ': ' + responseText);
    }
    
  } catch (error) {
    console.error("❌ Stage 1 OAuth failed:", error.message);
    throw new Error('Failed to get Trimble OAuth token: ' + error.message);
  }
}

// ===== MAIN GMAIL ADD-ON FUNCTIONS =====

function buildAddOn(e) {
    console.log("=== BUILD ADD-ON CALLED ===");
    
    try {
      var card = CardService.newCardBuilder();
      var section = CardService.newCardSection()
        .setHeader("Jira TPM Attachment Saver");
      
      // Check if settings are configured
      console.log("Checking user settings configuration...");
      var settings = getUserSettings();
      
      console.log("Settings validation results:");
      console.log("- Jira Token present:", !!settings.jiraToken);
      console.log("- Jira URL present:", !!settings.jiraUrl);
      
      // Check environment variables for PMO integration
      var pmoConfigured = false;
      try {
        getPMOWebhookURL();
        getTrimbleAuthServerToken();
        pmoConfigured = true;
        console.log("- PMO Environment Variables: Configured");
      } catch (error) {
        console.log("- PMO Environment Variables: NOT CONFIGURED -", error.message);
      }
      
      if (!settings.jiraToken || !settings.jiraUrl || !pmoConfigured) {
        console.log("SETTINGS INCOMPLETE - showing setup screen");
        console.log("Missing:");
        if (!settings.jiraToken) console.log("- Jira Token");
        if (!settings.jiraUrl) console.log("- Jira URL");
        if (!pmoConfigured) console.log("- PMO Environment Variables (run setupScriptProperties())");
        
        return [buildSettingsCard(true)]; // true = first time setup
      }
      
      console.log("Settings validation PASSED - proceeding with main interface");
      
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
      
      // Get attachments and Google Docs links from the email thread
      try {
        var threadId = e.gmail.threadId;
        var thread = GmailApp.getThreadById(threadId);
        var messages = thread.getMessages();
        var allAttachments = [];
        var allGoogleDocsLinks = [];
        
        for (var i = 0; i < messages.length; i++) {
          var message = messages[i];
          var attachments = message.getAttachments();
          
          // Process regular attachments
          for (var j = 0; j < attachments.length; j++) {
            var attachment = attachments[j];
            allAttachments.push({
              name: attachment.getName(),
              size: attachment.getSize(),
              attachment: attachment,
              messageIndex: i,
              type: 'regular'
            });
          }
          
          // Extract Google Docs links from message content
          var googleDocsLinks = extractGoogleDocsLinks(message);
          for (var k = 0; k < googleDocsLinks.length; k++) {
            var docLink = googleDocsLinks[k];
            allGoogleDocsLinks.push({
              name: docLink.title,
              size: 0, // Google Docs links don't have a file size
              url: docLink.url,
              docType: docLink.type,
              messageIndex: i,
              type: 'google_docs'
            });
          }
        }
        
        // Combine regular attachments and Google Docs links
        var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
        
                if (combinedAttachments.length > 0) {
            // Group attachments by file extension (including Google Docs links)
            var attachmentsByExtension = {};
            var extensionCounts = {};
            
            for (var i = 0; i < combinedAttachments.length; i++) {
              var attachment = combinedAttachments[i];
              var fileName = attachment.name;
              
              // Determine extension/category
              var extension;
              if (attachment.type === 'google_docs') {
                // Group Google Docs by their type (docs, sheets, slides, etc.)
                extension = 'google-' + attachment.docType;
              } else {
                // Regular file extension
                extension = fileName.lastIndexOf('.') > -1 ? 
                  fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              }
              
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
                headerText = "📎 Select Attachments & Docs: " + headerText + " (Total: " + combinedAttachments.length + ")";
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
                
                // Create display text based on attachment type
                var displayText;
                if (attachment.type === 'google_docs') {
                  var docTypeIcon = getGoogleDocsIcon(attachment.docType);
                  displayText = docTypeIcon + " " + attachment.name + " (Google " + attachment.docType.charAt(0).toUpperCase() + attachment.docType.slice(1) + ")";
                } else {
                  displayText = attachment.name + " (" + formatFileSize(attachment.size) + ")";
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(displayText, originalIndex.toString(), isSelected);
                
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
      
      // ===== SUBFOLDER SELECTION & SAVE BUTTON SECTION =====
      var saveSection = CardService.newCardSection()
        .setHeader("📁 Save Attachments");
      
      // Check if subfolder feature is enabled
      var useEnhancedUI = shouldUseEnhancedUI();
      
      if (useEnhancedUI) {
        // Add subfolder selection dropdown (no title to avoid text overlap)
        var subfolderDropdown = CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("selectedSubfolder")
          .addItem("Project Root (Default)", "", true)
          .addItem("01_System_Design", "01_System_Design", false)
          .addItem("02_Meet_Recordings", "02_Meet_Recordings", false)
          .addItem("03_Correspondence", "03_Correspondence", false)
          .addItem("04_Project_Documentation", "04_Project_Documentation", false)
          .addItem("04_Project_Documentation > Project_Management", "04_Project_Documentation/Project_Management", false)
          .addItem("04_Project_Documentation > Carrier_Onboarding", "04_Project_Documentation/Carrier_Onboarding", false);
        
        saveSection.addWidget(subfolderDropdown);
        
        // Add helpful text with better spacing
        var helperText = CardService.newTextParagraph()
          .setText("<font color=\"#888888\"><i>Folders created automatically if needed</i></font>");
        saveSection.addWidget(helperText);
      }
      
      var saveAction = CardService.newAction()
        .setFunctionName("saveSelectedAttachmentsToGDrive")
        .setParameters({threadId: (e && e.gmail) ? e.gmail.threadId : ""});
      
      var saveButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("💾 Save to PMO Folder")
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
          .setText("Save to Folder")
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
      .setValue(settings.jiraUrl || getDefaultJiraURL());
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
    
    // OAuth Test Button
    var oauthTestButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🔐 Test 2-Stage OAuth")
        .setOnClickAction(CardService.newAction().setFunctionName("testTrimbleOAuthFlow")));
    buttonSection.addWidget(oauthTestButtonSet);
    
    // Environment Variables Test Button
    var envTestButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("⚙️ Test Environment Variables")
        .setOnClickAction(CardService.newAction().setFunctionName("testEnvironmentVariables")));
    buttonSection.addWidget(envTestButtonSet);
    
    // OAuth Token Caching Test Button
    var tokenTestButtonSet = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("🕒 Test Token Caching")
        .setOnClickAction(CardService.newAction().setFunctionName("testOAuthTokenCaching")));
    buttonSection.addWidget(tokenTestButtonSet);
    
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
        .setValue(settings.jiraUrl || getDefaultJiraURL());
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
      
      // OAuth Test button
      var oauthTestButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("🔐 Test 2-Stage OAuth")
          .setOnClickAction(CardService.newAction().setFunctionName("testTrimbleOAuthFlow")));
      mainSection.addWidget(oauthTestButtonSet);
      
      // Environment Variables Test button
      var envTestButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("⚙️ Test Environment Variables")
          .setOnClickAction(CardService.newAction().setFunctionName("testEnvironmentVariables")));
      mainSection.addWidget(envTestButtonSet);
      
      // OAuth Token Caching Test button
      var tokenTestButtonSet = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("🕒 Test Token Caching")
          .setOnClickAction(CardService.newAction().setFunctionName("testOAuthTokenCaching")));
      mainSection.addWidget(tokenTestButtonSet);
      
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
      
      
      // Save settings (PMO variables are configured in environment)
      var settings = {
        jiraUrl: jiraUrl,
        jiraToken: finalToken,
        customJql: customJql || getDefaultJQL(),
        savedAt: new Date().toISOString()
      };
      
      setUserSettings(settings);
      
      console.log("Settings saved successfully");
      
      // Success notification for Jira settings only (PMO configured in environment)
      var successMsg = "✅ Jira Settings Saved Successfully!\n";
      successMsg += "• Jira URL: " + jiraUrl + "\n";
      successMsg += "• API Token: " + (finalToken ? "Configured" : "Not Set") + "\n";
      successMsg += "• JQL Query: " + (customJql ? "Custom" : "Default") + "\n\n";
      successMsg += "📋 PMO Integration: Configured via environment variables\n";
      successMsg += "💡 Use 'Test Connections' to verify functionality";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(successMsg))
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
  
  /**
   * Test all environment variables configuration
   */
  function testEnvironmentVariables(e) {
    console.log("=== TEST ENVIRONMENT VARIABLES ===");
    
    try {
      var results = [];
      var allPassed = true;
      
      // Test each environment variable
      try {
        var authToken = getTrimbleAuthServerToken();
        results.push("✅ TRIMBLE_AUTH_SERVER_TOKEN: Configured (" + authToken.length + " chars)");
      } catch (error) {
        results.push("❌ TRIMBLE_AUTH_SERVER_TOKEN: " + error.message);
        allPassed = false;
      }
      
      try {
        var pmoUrl = getPMOWebhookURL();
        results.push("✅ PMO_WEBHOOK_URL: " + pmoUrl);
      } catch (error) {
        results.push("❌ PMO_WEBHOOK_URL: " + error.message);
        allPassed = false;
      }
      
      try {
        var oauthUrl = getTrimbleOAuthURL();
        results.push("✅ TRIMBLE_OAUTH_URL: " + oauthUrl);
      } catch (error) {
        results.push("❌ TRIMBLE_OAUTH_URL: " + error.message);
        allPassed = false;
      }
      
      try {
        var jiraUrl = getDefaultJiraURL();
        results.push("✅ DEFAULT_JIRA_URL: " + jiraUrl);
      } catch (error) {
        results.push("❌ DEFAULT_JIRA_URL: " + error.message);
        allPassed = false;
      }
      
      try {
        var grantType = getOAuthGrantType();
        results.push("✅ OAUTH_GRANT_TYPE: " + grantType);
      } catch (error) {
        results.push("❌ OAUTH_GRANT_TYPE: " + error.message);
        allPassed = false;
      }
      
      try {
        var scope = getOAuthScope();
        results.push("✅ OAUTH_SCOPE: " + scope);
      } catch (error) {
        results.push("❌ OAUTH_SCOPE: " + error.message);
        allPassed = false;
      }
      
      try {
        var storageKey = getSettingsStorageKey();
        results.push("✅ SETTINGS_STORAGE_KEY: " + storageKey);
      } catch (error) {
        results.push("❌ SETTINGS_STORAGE_KEY: " + error.message);
        allPassed = false;
      }
      
      // Test OAuth token status
      try {
        var tokenStatus = getOAuthTokenStatus();
        results.push("🔐 OAUTH_TOKEN_STATUS: " + tokenStatus.message);
      } catch (error) {
        results.push("❌ OAUTH_TOKEN_STATUS: " + error.message);
      }
      
      var statusMsg = allPassed ? "✅ Environment Variables Test PASSED!" : "⚠️ Environment Variables Test - Issues Found";
      var resultText = statusMsg + "\n\n" + results.join("\n");
      
      if (!allPassed) {
        resultText += "\n\n💡 Run setupScriptProperties() to configure missing variables.";
      }
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(resultText))
        .build();
        
    } catch (error) {
      console.error("❌ Environment variables test failed:", error.message);
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("❌ Environment Variables Test FAILED\n• Error: " + error.message + "\n• Run setupScriptProperties() first"))
        .build();
    }
  }

  /**
   * Test OAuth token caching and expiration handling
   */
  function testOAuthTokenCaching(e) {
    console.log("=== TEST OAUTH TOKEN CACHING ===");
    
    try {
      var results = [];
      
      // Step 1: Check initial token status
      var initialStatus = getOAuthTokenStatus();
      results.push("📋 Initial Status: " + initialStatus.message);
      
      // Step 2: Get a token (this will cache it)
      console.log("Getting OAuth token (will be cached)...");
      var token1 = getTrimbleOAuthToken();
      results.push("✅ Token Retrieved: " + token1.substring(0, 20) + "... (cached)");
      
      // Step 3: Get token again (should use cached)
      console.log("Getting OAuth token again (should use cached)...");
      var token2 = getTrimbleOAuthToken();
      var isSameToken = token1 === token2;
      results.push((isSameToken ? "✅" : "❌") + " Cache Test: " + (isSameToken ? "Used cached token" : "New token requested"));
      
      // Step 4: Check current status
      var currentStatus = getOAuthTokenStatus();
      results.push("📋 Current Status: " + currentStatus.message);
      
      // Step 5: Test manual cache clearing
      console.log("Testing manual cache clearing...");
      clearCachedOAuthToken();
      var clearedStatus = getOAuthTokenStatus();
      results.push("🗑️ After Clear: " + clearedStatus.message);
      
      // Step 6: Get fresh token after clearing
      console.log("Getting fresh token after clearing cache...");
      var token3 = getTrimbleOAuthToken();
      var isDifferentToken = token1 !== token3;
      results.push((isDifferentToken ? "✅" : "❌") + " Fresh Token: " + (isDifferentToken ? "New token obtained" : "Same token (unexpected)"));
      
      // Step 7: Final status
      var finalStatus = getOAuthTokenStatus();
      results.push("📋 Final Status: " + finalStatus.message);
      
      var successMsg = "✅ OAuth Token Caching Test COMPLETED!\n\n";
      successMsg += results.join("\n");
      successMsg += "\n\n💡 Token expires in 1 hour (3600 seconds)";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(successMsg))
        .build();
        
    } catch (error) {
      console.error("❌ OAuth token caching test failed:", error.message);
      
      var errorMsg = "❌ OAuth Token Caching Test FAILED\n";
      errorMsg += "• Error: " + error.message + "\n";
      errorMsg += "• Check environment variable configuration\n";
      errorMsg += "• Ensure TRIMBLE_AUTH_SERVER_TOKEN is set correctly";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(errorMsg))
        .build();
    }
  }

  /**
   * Test the complete 2-stage authentication flow
   */
  function testTrimbleOAuthFlow(e) {
    console.log("=== TEST TRIMBLE 2-STAGE OAUTH FLOW ===");
    
    try {
      // Stage 1: Test OAuth token retrieval
      console.log("Testing Stage 1: OAuth token retrieval...");
      var oauthToken = getTrimbleOAuthToken();
      
      if (oauthToken) {
        console.log("✅ Stage 1 SUCCESS: OAuth token retrieved");
        
        // Stage 2: Test PMO webhook call with the token
        console.log("Testing Stage 2: PMO webhook call with Bearer token...");
        var testTicket = "CXPRODELIVERY-OAUTH-TEST-" + new Date().getTime();
        var webhookUrl = getPMOWebhookURL();
        
        var response = UrlFetchApp.fetch(webhookUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + oauthToken
          },
          payload: JSON.stringify({"text": testTicket}),
          muteHttpExceptions: true,
          timeout: 10000
        });
        
        var responseCode = response.getResponseCode();
        console.log("Stage 2 response code:", responseCode);
        console.log("Stage 2 response:", response.getContentText());
        
        if (responseCode === 200) {
          var successMsg = "✅ 2-Stage OAuth Authentication Test SUCCESSFUL!\n";
          successMsg += "• Stage 1: OAuth token retrieved ✓\n";
          successMsg += "• Stage 2: PMO webhook call with Bearer token ✓\n";
          successMsg += "• Test Ticket: " + testTicket + "\n";
          successMsg += "• Response Code: " + responseCode + "\n";
          successMsg += "🔐 Authentication flow is working correctly!";
          
          return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
              .setText(successMsg))
            .build();
        } else {
          throw new Error("Stage 2 failed with code " + responseCode + ": " + response.getContentText());
        }
      }
    } catch (error) {
      console.error("❌ 2-Stage OAuth test failed:", error.message);
      
      var errorMsg = "❌ 2-Stage OAuth Authentication Test FAILED\n";
      errorMsg += "• Error: " + error.message + "\n";
      errorMsg += "• Check auth server token configuration\n";
      errorMsg += "• Verify network connectivity\n";
      errorMsg += "• Review logs for detailed error information";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(errorMsg))
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
      
      // Enhanced testing with multiple scenarios
      console.log("Testing PMO webhook connectivity and folder operations...");
      
      // Test 1: Basic connectivity with test ticket
      var testTicket = "CXPRODELIVERY-TEST-" + new Date().getTime();
      console.log("Testing PMO connection with unique test ticket:", testTicket);
      
      var result = getPMOProjectFolder(testTicket);
      
      if (result.success) {
        // Test 2: Verify folder access
        var folderTest = getFolderByIdSafely(result.folderId);
        
        if (folderTest.success) {
          var successMsg = "✅ PMO Connection Test Successful!\n";
          successMsg += "• Webhook: Responsive\n";  
          successMsg += "• Folder ID: " + result.folderId + "\n";
          successMsg += "• Folder Name: " + folderTest.name + "\n";
          successMsg += "• Status: " + (result.created ? "New folder created" : "Existing folder found") + "\n";
          successMsg += "• Access: Read/Write permissions confirmed";
          
          return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
              .setText(successMsg))
            .build();
        } else {
          return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
              .setText("⚠️ PMO webhook responded but folder access failed: " + folderTest.error))
            .build();
        }
      } else {
        // Enhanced error reporting
        var errorMsg = "❌ PMO Connection Test Failed\n";
        errorMsg += "• Webhook URL: " + webhookUrl + "\n";
        errorMsg += "• Test Ticket: " + testTicket + "\n";
        errorMsg += "• Error: " + result.error + "\n";
        errorMsg += "• Suggestion: Check webhook URL and network connectivity";
        
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText(errorMsg))
          .build();
      }
      
    } catch (error) {
      var errorMsg = "❌ PMO Connection Test Error\n";
      errorMsg += "• Error: " + error.message + "\n";
      errorMsg += "• Check: Network connectivity and webhook configuration";
      
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(errorMsg))
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
          var allGoogleDocsLinks = [];
          console.log("Found", messages.length, "messages in thread");
          
          for (var i = 0; i < messages.length; i++) {
            var message = messages[i];
            var attachments = message.getAttachments();
            console.log("Message", i + 1, "has", attachments.length, "attachments");
            
            // Process regular attachments
            for (var j = 0; j < attachments.length; j++) {
              var attachment = attachments[j];
              allAttachments.push({
                name: attachment.getName(),
                size: attachment.getSize(),
                attachment: attachment,
                messageIndex: i,
                type: 'regular'
              });
            }
            
            // Extract Google Docs links from message content
            var googleDocsLinks = extractGoogleDocsLinks(message);
            console.log("Message", i + 1, "has", googleDocsLinks.length, "Google Docs links");
            for (var k = 0; k < googleDocsLinks.length; k++) {
              var docLink = googleDocsLinks[k];
              allGoogleDocsLinks.push({
                name: docLink.title,
                size: 0, // Google Docs links don't have a file size
                url: docLink.url,
                docType: docLink.type,
                messageIndex: i,
                type: 'google_docs'
              });
            }
          }
          
          // Combine regular attachments and Google Docs links
          var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
          
          console.log("Total attachments and Google Docs links found:", combinedAttachments.length);
          
          if (combinedAttachments.length > 0) {
            console.log("Creating attachment section...");
            
            // First, check if we have current form selections to preserve
            var currentSelections = {};
            var hasCurrentSelections = false;
            
            if (e.formInput) {
              console.log("=== CHECKING CURRENT FORM SELECTIONS ===");
              for (var i = 0; i < combinedAttachments.length; i++) {
                var formFieldName = "attachment_" + i;
                // If field exists in form input → selected = true, if missing → selected = false
                var isCurrentlySelected = e.formInput.hasOwnProperty(formFieldName) && e.formInput[formFieldName] && e.formInput[formFieldName].length > 0;
                
                // Use unique key to handle duplicate filenames (filename + index)
                var uniqueKey = combinedAttachments[i].name + "_" + i;
                currentSelections[uniqueKey] = isCurrentlySelected;
                hasCurrentSelections = true;
                console.log("Current form state for", combinedAttachments[i].name, "- field:", formFieldName, "- exists:", e.formInput.hasOwnProperty(formFieldName), "- selected:", isCurrentlySelected, "- stored as:", uniqueKey);
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
            
            for (var i = 0; i < combinedAttachments.length; i++) {
              var attachment = combinedAttachments[i];
              var fileName = attachment.name;
              
              // Determine extension/category
              var extension;
              if (attachment.type === 'google_docs') {
                // Group Google Docs by their type (docs, sheets, slides, etc.)
                extension = 'google-' + attachment.docType;
              } else {
                // Regular file extension
                extension = fileName.lastIndexOf('.') > -1 ? 
                fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase() : 'no-extension';
              }
              
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
                headerText = "📎 Select Attachments & Docs: " + headerText + " (Total: " + combinedAttachments.length + ")";
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
                
                // Create display text based on attachment type
                var displayText;
                if (attachment.type === 'google_docs') {
                  var docTypeIcon = getGoogleDocsIcon(attachment.docType);
                  displayText = docTypeIcon + " " + attachment.name + " (Google " + attachment.docType.charAt(0).toUpperCase() + attachment.docType.slice(1) + ")";
                } else {
                  displayText = attachment.name + " (" + formatFileSize(attachment.size) + ")";
                }
                
                var checkbox = CardService.newSelectionInput()
                  .setType(CardService.SelectionInputType.CHECK_BOX)
                  .setFieldName("attachment_" + originalIndex)
                  .addItem(displayText, originalIndex.toString(), isSelected);
                
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
      
      // ===== SUBFOLDER SELECTION & SAVE BUTTON =====
      // Check if subfolder feature is enabled
      var useEnhancedUI = shouldUseEnhancedUI();
      
      if (useEnhancedUI) {
        // Create separate section for subfolder selection
        var subfolderSection = CardService.newCardSection()
          .setHeader("📁 Save Attachments");
        
        // Dropdown without title to avoid text overlap
        var subfolderDropdown = CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.DROPDOWN)
          .setFieldName("selectedSubfolder")
          .addItem("Project Root (Default)", "", true)
          .addItem("01_System_Design", "01_System_Design", false)
          .addItem("02_Meet_Recordings", "02_Meet_Recordings", false)
          .addItem("03_Correspondence", "03_Correspondence", false)
          .addItem("04_Project_Documentation", "04_Project_Documentation", false)
          .addItem("04_Project_Documentation > Project_Management", "04_Project_Documentation/Project_Management", false)
          .addItem("04_Project_Documentation > Carrier_Onboarding", "04_Project_Documentation/Carrier_Onboarding", false);
        
        subfolderSection.addWidget(subfolderDropdown);
        
        // Add helpful text with better spacing
        var helperText = CardService.newTextParagraph()
          .setText("<font color=\"#888888\"><i>Folders created automatically if needed</i></font>");
        subfolderSection.addWidget(helperText);
        
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("💾 Save to PMO Folder")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("saveSelectedAttachmentsToGDrive")
              .setParameters({threadId: threadId})));
        
        subfolderSection.addWidget(saveButtonSet);
        card.addSection(subfolderSection);
      } else {
        // Legacy UI - simple save button
        var saveButtonSet = CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Save to PMO Folder")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("saveSelectedAttachmentsToGDrive")
              .setParameters({threadId: threadId})));
        mainSection.addWidget(saveButtonSet);
      }
      
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
      console.log("Save operation timestamp:", new Date().toISOString());
      console.log("Form input:", JSON.stringify(e.formInput));
      console.log("Parameters:", JSON.stringify(e.parameters));
      
      // Add comprehensive diagnostic information for troubleshooting
      console.log("\n=== DIAGNOSTIC INFO FOR TROUBLESHOOTING ===");
      logDiagnosticInfo();
      
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
      
      // Get all attachments and Google Docs links from the thread
      var thread = GmailApp.getThreadById(threadId);
      var messages = thread.getMessages();
      var allAttachments = [];
      var allGoogleDocsLinks = [];
      
      for (var i = 0; i < messages.length; i++) {
        var message = messages[i];
        var attachments = message.getAttachments();
        
        // Process regular attachments
        for (var j = 0; j < attachments.length; j++) {
          var attachment = attachments[j];
          allAttachments.push({
            name: attachment.getName(),
            size: attachment.getSize(),
            attachment: attachment,
            messageIndex: i,
            type: 'regular'
          });
        }
        
        // Extract Google Docs links from message content
        var googleDocsLinks = extractGoogleDocsLinks(message);
        var emailSubject = message.getSubject();
        for (var k = 0; k < googleDocsLinks.length; k++) {
          var docLink = googleDocsLinks[k];
          allGoogleDocsLinks.push({
            name: docLink.title,
            size: 0, // Google Docs links don't have a file size
            url: docLink.url,
            docType: docLink.type,
            messageIndex: i,
            type: 'google_docs',
            emailSubject: emailSubject
          });
        }
      }
      
      // Combine regular attachments and Google Docs links
      var combinedAttachments = allAttachments.concat(allGoogleDocsLinks);
      
      console.log("Found", combinedAttachments.length, "total attachments and Google Docs links");
      
      if (combinedAttachments.length === 0) {
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
      console.log("Total attachments to process:", combinedAttachments.length);
      console.log("Form input keys:", Object.keys(e.formInput));
      
      for (var i = 0; i < combinedAttachments.length; i++) {
        var checkboxValue = e.formInput["attachment_" + i];
        var isSelected = checkboxValue && checkboxValue.length > 0;
        var attachmentName = combinedAttachments[i].name;
        var attachmentType = combinedAttachments[i].type || 'regular';
        
        console.log("Attachment", i, "(", attachmentName, ") - type:", attachmentType, "- checkbox value:", checkboxValue, "- selected:", isSelected);
        
        // Store selection state for this attachment using unique key
        var uniqueKey = attachmentName + "_" + i;
        selectionState[uniqueKey] = isSelected;
        
        if (isSelected) {
          selectedAttachments.push(combinedAttachments[i]);
          console.log("Added to selected list:", attachmentName, "(" + attachmentType + ")");
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
      
  // N8N webhook Integration Logic (ONLY METHOD) - Lookup/Create folder
      console.log("=== STARTING N8N webhook INTEGRATION ===");
      console.log("Final ticket for Jira lookup:", finalTicket);
      console.log("Selected attachments count:", selectedAttachments.length);
      console.log("Integration timestamp:", new Date().toISOString());
      
      // Enhanced user feedback during N8N webhook operation
      console.log("🔍 Contacting N8N Webhook system for project folder...");
      
      var pmoResult = getPMOProjectFolder(finalTicket);
      
      console.log("Lookup completed. Result:", JSON.stringify(pmoResult));
      
      if (!pmoResult.success) {
        // Enhanced N8N webhook error handling with user-friendly messages
        console.error("Lookup failed:", pmoResult.error);
        
        // Get selected subfolder for enhanced error message
        var selectedSubfolder = e.formInput.selectedSubfolder || "";
        var userFriendlyError = getPMOSubfolderErrorMessage(pmoResult.error, finalTicket, selectedSubfolder);
        
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification()
            .setText(userFriendlyError))
          .build();
      }
      
      console.log("N8N webhook returned folder ID:", pmoResult.folderId, "- Created:", pmoResult.created || false);
      
      var folderResult = getFolderByIdSafely(pmoResult.folderId);
      
      if (!folderResult.success) {
        // Enhanced folder access error handling - Show request access card
        console.error("Gdrive folder access failed:", folderResult.error);
        
        // Build and show request access card
        var requestAccessCard = buildRequestAccessCard(finalTicket, pmoResult.folderId, folderResult.error);
        
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation().pushCard(requestAccessCard))
          .build();
      }
      
      var projectRootFolder = folderResult.folder;
      console.log("Using PMO folder:", folderResult.name);
      
      // ===== NEW: SUBFOLDER SELECTION SUPPORT =====
      var selectedSubfolder = e.formInput.selectedSubfolder || "";
      console.log("=== SUBFOLDER PROCESSING ===");
      console.log("Selected subfolder:", selectedSubfolder || "(Project Root)");
      
      // Validate subfolder selection
      var validation = validateSubfolderSelection(selectedSubfolder);
      if (!validation.valid) {
        console.warn("Subfolder validation warning:", validation.warning);
        selectedSubfolder = validation.sanitized;
      }
      
      // Get or create target subfolder
      var targetFolderResult = getOrCreateProjectSubfolder(projectRootFolder, selectedSubfolder);
      
      if (!targetFolderResult.success) {
        // Attempt fallback strategy
        console.error("Subfolder creation failed, attempting fallback");
        var fallbackResult = handleSubfolderCreationFailure(projectRootFolder, selectedSubfolder, targetFolderResult);
        
        if (!fallbackResult.success) {
          // Complete failure - show error
          var subfolderError = getSubfolderCreationErrorMessage(
            targetFolderResult.error,
            projectRootFolder,
            selectedSubfolder
          );
          
          return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
              .setText(subfolderError))
            .build();
        }
        
        // Fallback succeeded - use fallback folder
        targetFolderResult = fallbackResult;
        console.log("Using fallback strategy:", fallbackResult.fallbackUsed);
      }
      
      var ticketFolder = targetFolderResult.folder;
      var actualFolderPath = targetFolderResult.path;
      console.log("Target folder ready:", actualFolderPath);
      console.log("Target folder ID:", ticketFolder.getId());
      // ===== END SUBFOLDER SUPPORT =====
      
      // Save selected attachments with duplicate handling
      var savedCount = 0;
      var skippedCount = 0;
      var savedFiles = [];
      var skippedFiles = [];
      
      for (var i = 0; i < selectedAttachments.length; i++) {
        try {
          var attachmentData = selectedAttachments[i];
          var fileName = attachmentData.name;
          var attachmentType = attachmentData.type || 'regular';
          console.log("Processing", attachmentType, "attachment:", fileName);
          
          var file;
          var finalFileName;
          
          if (attachmentType === 'google_docs') {
            // Handle Google Docs links
            console.log("Copying Google Doc to folder:", fileName);
            file = copyGoogleDocToFolder({
              title: attachmentData.name,
              url: attachmentData.url,
              type: attachmentData.docType,
              emailSubject: attachmentData.emailSubject
            }, ticketFolder);
            finalFileName = file.getName();
            savedCount++;
            savedFiles.push(finalFileName);
            console.log("Successfully copied Google Doc:", finalFileName);
            
          } else {
            // Handle regular attachments
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
                file = ticketFolder.createFile(attachmentData.attachment);
                file.setName(newFileName);
                finalFileName = newFileName;
                savedCount++;
                savedFiles.push(finalFileName);
              }
            } else {
              // File doesn't exist, save normally
              console.log("Saving new file:", fileName);
              file = ticketFolder.createFile(attachmentData.attachment);
              finalFileName = file.getName();
              savedCount++;
              savedFiles.push(finalFileName);
            }
          }
          
          console.log("Processed:", fileName, "->", finalFileName);
        } catch (saveError) {
          console.error("Error saving", attachmentData.type || 'regular', "attachment", attachmentData.name, ":", saveError);
        }
      }
      
      console.log("=== SAVE COMPLETE ===");
      console.log("Saved", savedCount, "files to", finalTicket);
      console.log("Skipped", skippedCount, "duplicate files");
      console.log("Files saved:", savedFiles);
      console.log("Files skipped:", skippedFiles);
      
      // Enhanced notification message with detailed PMO info and subfolder path
      var notificationText = "✅ Attachment Save Complete!\n\n";
      
      // PMO folder information
      if (pmoResult.created) {
        notificationText += "📁 New PMO folder created: " + folderResult.name + "\n";
      } else {
        notificationText += "📁 PMO project folder: " + folderResult.name + "\n";
      }
      
      // Subfolder information
      if (actualFolderPath && actualFolderPath !== "Project Root") {
        notificationText += "📂 Subfolder: " + actualFolderPath + "\n";
        if (targetFolderResult.created) {
          notificationText += "🆕 Created new subfolder(s): " + targetFolderResult.createdFolders.join(", ") + "\n";
        }
        if (targetFolderResult.fallbackUsed) {
          notificationText += "⚠️ Fallback: " + targetFolderResult.warning + "\n";
        }
      } else {
        notificationText += "📂 Location: Project Root\n";
      }
      
      // File statistics
      notificationText += "📎 Files processed: " + selectedAttachments.length + "\n";
      notificationText += "💾 Files saved: " + savedCount + "\n";
      
      if (skippedCount > 0) {
        notificationText += "⚠️ Duplicates skipped: " + skippedCount + "\n";
      }
      
      // Project information
      notificationText += "🎫 Project: " + finalTicket + "\n";
      
      // Success summary
      var successRate = Math.round((savedCount / selectedAttachments.length) * 100);
      notificationText += "✨ Success rate: " + successRate + "%";
      
      if (savedFiles.length > 0 && savedFiles.length <= 3) {
        notificationText += "\n\n📋 Saved files:\n• " + savedFiles.join("\n• ");
      } else if (savedFiles.length > 3) {
        notificationText += "\n\n📋 Recent files: " + savedFiles.slice(0, 2).join(", ") + " and " + (savedFiles.length - 2) + " more";
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
    console.log("=== LOADING USER SETTINGS ===");
    console.log("Timestamp:", new Date().toISOString());
    
    try {
      var userProperties = PropertiesService.getUserProperties();
      var settingsJson = userProperties.getProperty(getSettingsStorageKey());
      
      console.log("Settings JSON from storage:", settingsJson ? "found" : "not found");
      
      if (settingsJson) {
        console.log("Raw settings JSON length:", settingsJson.length);
        var settings = JSON.parse(settingsJson);
        console.log("Parsed settings keys:", Object.keys(settings));
        
        // Add PMO defaults for existing users who don't have PMO settings
        if (!settings.pmoWebhookUrl) {
          console.log("Adding default PMO webhook URL for existing user");
          settings.pmoWebhookUrl = getDefaultPMOWebhookUrl();
        }
        if (!settings.pmoTimeout) {
          console.log("Adding default PMO timeout for existing user");
          settings.pmoTimeout = 10000;
        }
        if (!settings.pmoRetryAttempts) {
          console.log("Adding default PMO retry attempts for existing user");
          settings.pmoRetryAttempts = 2;
        }
        
        console.log("Final settings configuration:");
        console.log("- Jira URL:", settings.jiraUrl ? "configured" : "NOT SET");
        console.log("- Jira Token:", settings.jiraToken ? "configured (length: " + settings.jiraToken.length + ")" : "NOT SET");
        console.log("- JQL Query:", settings.jqlQuery ? "configured" : "NOT SET");
        console.log("- PMO Webhook URL:", settings.pmoWebhookUrl ? "configured" : "NOT SET");
        console.log("- Trimble Auth Token:", settings.trimbleAuthToken ? "configured (length: " + settings.trimbleAuthToken.length + ")" : "NOT SET");
        console.log("- PMO Timeout:", settings.pmoTimeout);
        console.log("- PMO Retry Attempts:", settings.pmoRetryAttempts);
        console.log("- Using default PMO URL:", settings.pmoWebhookUrl === getDefaultPMOWebhookUrl());
        
        return settings;
      } else {
        console.log("No existing settings found - returning defaults for new user");
        var defaultSettings = {
          pmoWebhookUrl: getDefaultPMOWebhookUrl(),
          pmoTimeout: 10000,
          pmoRetryAttempts: 2
        };
        console.log("Default settings:", JSON.stringify(defaultSettings));
        return defaultSettings;
      }
    } catch (error) {
      console.error("CRITICAL: Error getting user settings:", error.name, "-", error.message);
      console.error("Settings error stack:", error.stack);
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
      userProperties.setProperty(getSettingsStorageKey(), JSON.stringify(settings));
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
  
  // ===== FOLDER ACCESS REQUEST CARD =====
  
  /**
   * Build a card for requesting folder access
   * @param {string} ticketKey - The Jira ticket key
   * @param {string} folderId - The Google Drive folder ID
   * @param {string} errorMessage - The detailed error message
   * @returns {Card} A card with request access button
   */
  function buildRequestAccessCard(ticketKey, folderId, errorMessage) {
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle("🔒 Folder Access Required")
        .setSubtitle(ticketKey));
    
    var section = CardService.newCardSection();
    
    // Error explanation
    var errorText = CardService.newTextParagraph()
      .setText("<b>⚠️ Cannot Access Project Folder</b><br><br>" +
               "The PMO system found the project folder, but you don't have permission to access it yet.");
    section.addWidget(errorText);
    
    // Folder info
    var folderInfo = CardService.newKeyValue()
      .setTopLabel("Project Ticket")
      .setContent(ticketKey)
      .setIcon(CardService.Icon.TICKET);
    section.addWidget(folderInfo);
    
    var folderIdInfo = CardService.newKeyValue()
      .setTopLabel("Folder ID")
      .setContent(folderId)
      .setIcon(CardService.Icon.DESCRIPTION);
    section.addWidget(folderIdInfo);
    
    // Technical error (collapsible)
    var technicalError = CardService.newTextParagraph()
      .setText("<font color=\"#888888\"><i>Technical details: " + errorMessage + "</i></font>");
    section.addWidget(technicalError);
    
    card.addSection(section);
    
    // Action section
    var actionSection = CardService.newCardSection()
      .setHeader("🔧 What to Do");
    
    var instructions = CardService.newTextParagraph()
      .setText("Click the button below to open the folder in Google Drive and request access:");
    actionSection.addWidget(instructions);
    
    // Request access button with direct folder URL
    var folderUrl = "https://drive.google.com/drive/folders/" + folderId;
    var requestAccessButton = CardService.newTextButton()
      .setText("🔗 Open Folder & Request Access")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(folderUrl)
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING));
    
    var buttonSet = CardService.newButtonSet()
      .addButton(requestAccessButton);
    actionSection.addWidget(buttonSet);
    
    // Additional instructions
    var additionalHelp = CardService.newTextParagraph()
      .setText("<br><b>After requesting access:</b><br>" +
               "• The folder owner will receive your request<br>" +
               "• You'll get an email when access is granted<br>" +
               "• Come back and try saving again<br><br>" +
               "<b>Alternative:</b> Contact your project manager for immediate access");
    actionSection.addWidget(additionalHelp);
    
    card.addSection(actionSection);
    
    return card.build();
  }
  
  // ===== SUBFOLDER ERROR HANDLING =====
  
  /**
   * Get enhanced error message for PMO errors with subfolder context
   * @param {string} error - The error message
   * @param {string} ticketKey - The Jira ticket key
   * @param {string} subfolderPath - The subfolder path (optional)
   * @returns {string} User-friendly error message
   */
  function getPMOSubfolderErrorMessage(error, ticketKey, subfolderPath) {
    var baseError = getPMOErrorMessage(error, ticketKey); // Use existing function
    
    if (subfolderPath && subfolderPath.trim() !== "") {
      var subfolderError = baseError + "\n\n📁 Additional Context:\n";
      subfolderError += "• Target subfolder: " + subfolderPath + "\n";
      subfolderError += "• Subfolder creation was blocked by PMO access issue\n";
      subfolderError += "• Files will be saved to project root once PMO access is restored";
      return subfolderError;
    }
    
    return baseError;
  }
  
  /**
   * Get error message for subfolder creation failures
   * @param {string} error - The error message
   * @param {Folder} projectFolder - The project root folder
   * @param {string} subfolderPath - The subfolder path
   * @returns {string} User-friendly error message
   */
  function getSubfolderCreationErrorMessage(error, projectFolder, subfolderPath) {
    var errorMsg = "❌ Cannot create subfolder in PMO project folder\n\n";
    errorMsg += "📁 Project Folder: " + projectFolder.getName() + "\n";
    errorMsg += "🎯 Subfolder Path: " + subfolderPath + "\n";
    errorMsg += "📋 Technical Error: " + error + "\n\n";
    
    errorMsg += "🔧 Possible Solutions:\n";
    errorMsg += "• Check if you have write permissions to the project folder\n";
    errorMsg += "• Try saving to 'Project Root' as a fallback\n";
    errorMsg += "• Contact project admin if folder permissions are restricted\n\n";
    
    errorMsg += "💡 Tip: You can still save attachments by selecting 'Project Root'";
    
    return errorMsg;
  }
  
  // ===== PMO INTEGRATION FUNCTIONS =====
  
  /**
   * Internal function to make PMO webhook request (for retry logic)
   */
  function makePMOWebhookRequest(ticketKey, webhookUrl, timeout) {
    var payload = JSON.stringify({"text": ticketKey});
    console.log("Sending PMO webhook request:");
    console.log("- Method: POST");
    console.log("- URL:", webhookUrl);
    console.log("- Payload:", payload);
    console.log("- Timeout:", timeout + "ms");
    
    var requestStartTime = new Date().getTime();
    
    // Stage 1: Get OAuth token from Trimble auth server
    console.log("=== STAGE 2: PMO WEBHOOK CALL WITH BEARER TOKEN ===");
    var oauthToken = getTrimbleOAuthToken();
    
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + oauthToken
      },
      payload: payload,
      muteHttpExceptions: true,
      timeout: timeout
    });
    
    var requestDuration = new Date().getTime() - requestStartTime;
    console.log("PMO webhook request completed in:", requestDuration + "ms");
    console.log("PMO webhook response code:", response.getResponseCode());
    console.log("PMO webhook response headers:", JSON.stringify(response.getAllHeaders()));
    
    return response;
  }
  
  function getPMOProjectFolder(ticketKey) {
    try {
      var settings = getUserSettings();
      var webhookUrl = getPMOWebhookURL();
      var maxRetries = settings.pmoRetryAttempts || 2;
      var timeout = settings.pmoTimeout || 10000;
      
      console.log("=== PMO WEBHOOK LOOKUP/CREATE ===");
      console.log("Timestamp:", new Date().toISOString());
      console.log("Ticket key:", ticketKey);
      console.log("PMO Configuration:");
      console.log("- Webhook URL:", webhookUrl);
      console.log("- Max retries:", maxRetries);  
      console.log("- Timeout (ms):", timeout);
      console.log("- Using default URL:", webhookUrl === getDefaultPMOWebhookUrl());
      
      // Check OAuth token status
      var tokenStatus = getOAuthTokenStatus();
      console.log("OAuth token status:", tokenStatus.message);
      
      if (!webhookUrl || webhookUrl.trim() === '') {
        console.error("CRITICAL: PMO webhook URL is empty or undefined!");
        return {
          success: false,
          error: 'PMO webhook URL not configured'
        };
      }
      
      for (var attempt = 1; attempt <= maxRetries; attempt++) {
        console.log("PMO request attempt", attempt, "of", maxRetries);
        
        try {
          var response = makePMOWebhookRequest(ticketKey, webhookUrl, timeout);
          var responseCode = response.getResponseCode();
          
          // Check for authentication errors (token expiration)
          if (responseCode === 401 || responseCode === 403) {
            console.log("🔄 Detected authentication error (code " + responseCode + "), may be expired token");
            
            // Try with fresh token (only once per attempt to avoid infinite loops)
            try {
              console.log("🔄 Clearing cached token and retrying with fresh token...");
              clearCachedOAuthToken();
              var retryResponse = makePMOWebhookRequest(ticketKey, webhookUrl, timeout);
              
              if (retryResponse.getResponseCode() === 200) {
                console.log("✅ Retry with fresh token succeeded");
                response = retryResponse; // Use the successful retry response
                responseCode = 200;
              } else {
                console.log("❌ Retry with fresh token also failed:", retryResponse.getResponseCode());
                // Fall through to regular error handling
              }
            } catch (retryError) {
              console.error("❌ Token refresh retry failed:", retryError.message);
              // Fall through to regular error handling
            }
          }
          
          if (responseCode === 200) {
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
            // Enhanced HTTP error logging
            var errorBody = response.getContentText();
            console.error("PMO webhook HTTP error details:");
            console.error("- Response code:", response.getResponseCode());
            console.error("- Response body:", errorBody);
            console.error("- Response headers:", JSON.stringify(response.getAllHeaders()));
            console.error("- Request duration:", (new Date().getTime() - requestStartTime) + "ms");
            
            return {
              success: false,
              error: 'PMO webhook HTTP error: ' + response.getResponseCode() + ' - ' + errorBody
            };
          }
        } catch (requestError) {
          // Enhanced network error logging
          console.error("PMO webhook network error on attempt", attempt + ":");
          console.error("- Error type:", requestError.name || 'Unknown');
          console.error("- Error message:", requestError.message || 'No message');
          console.error("- Error stack:", requestError.stack || 'No stack');
          console.error("- Request duration before error:", (new Date().getTime() - requestStartTime) + "ms");
          console.error("- Timeout setting:", timeout + "ms");
          
          // Check specific error types
          if (requestError.message && requestError.message.includes('timeout')) {
            console.error("TIMEOUT ERROR: Request exceeded timeout of", timeout + "ms");
          } else if (requestError.message && requestError.message.includes('network')) {
            console.error("NETWORK ERROR: Connection failed to", webhookUrl);
          }
          
          if (attempt === maxRetries) {
            console.error("PMO webhook failed after all", maxRetries, "retry attempts");
            return {
              success: false,
              error: 'PMO webhook network error after ' + maxRetries + ' attempts: ' + requestError.message
            };
          }
          
          console.log("Retrying PMO request in 1 second... (attempt", attempt, "of", maxRetries + ")");
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
  
  function logDiagnosticInfo() {
    console.log("=== DIAGNOSTIC INFORMATION DUMP ===");
    console.log("Timestamp:", new Date().toISOString());
    console.log("Apps Script execution info:");
    
    try {
      // Log execution environment (with permission error handling)
      try {
        console.log("- Time zone:", Session.getScriptTimeZone());
      } catch (tzError) {
        console.log("- Time zone: Error getting timezone -", tzError.message);
      }
      
      try {
        console.log("- Active user email:", Session.getActiveUser().getEmail());
      } catch (userError) {
        console.log("- Active user email: Permission error -", userError.message);
        console.log("- Note: userinfo.email OAuth scope may need authorization");
      }
      
      try {
        console.log("- Script permissions:", Session.getScriptTimeZone() ? "timezone authorized" : "timezone not authorized");
      } catch (permError) {
        console.log("- Script permissions: Error checking permissions -", permError.message);
      }
      
      // Log settings
      console.log("\nUser settings diagnostic:");
      var settings = getUserSettings();
      console.log("- Settings object type:", typeof settings);
      console.log("- Settings keys:", Object.keys(settings || {}));
      if (settings) {
        console.log("- Jira URL configured:", !!settings.jiraUrl);
        console.log("- Jira Token configured:", !!settings.jiraToken);
        console.log("- PMO URL configured:", !!settings.pmoWebhookUrl);
        console.log("- PMO timeout:", settings.pmoTimeout);
        console.log("- PMO retries:", settings.pmoRetryAttempts);
      }
      
      // Log PMO connectivity 
      console.log("\nPMO connectivity diagnostic:");
      console.log("- Default PMO URL:", getDefaultPMOWebhookUrl());
      console.log("- URL fetch available:", typeof UrlFetchApp !== 'undefined');
      
      // Log Gmail context (with API access error handling)
      console.log("\nAPI availability diagnostic:");
      try {
        if (typeof GmailApp !== 'undefined') {
          console.log("- Gmail API: Available");
          // Test basic Gmail access
          try {
            var gmailQuota = GmailApp.getRemainingDailyQuota();
            console.log("- Gmail quota remaining:", gmailQuota);
          } catch (gmailError) {
            console.log("- Gmail access: Permission error -", gmailError.message);
          }
        } else {
          console.log("- Gmail API: Not available (manual testing mode)");
        }
      } catch (gmailCheckError) {
        console.log("- Gmail API: Error checking availability -", gmailCheckError.message);
      }
      
      try {
        if (typeof DriveApp !== 'undefined') {
          console.log("- Drive API: Available");
          // Test basic Drive access
          try {
            var driveQuota = DriveApp.getStorageUsed();
            console.log("- Drive storage used:", driveQuota, "bytes");
          } catch (driveError) {
            console.log("- Drive access: Permission error -", driveError.message);
          }
        } else {
          console.log("- Drive API: Not available");
        }
      } catch (driveCheckError) {
        console.log("- Drive API: Error checking availability -", driveCheckError.message);
      }
      
      try {
        console.log("- UrlFetch API: Available -", typeof UrlFetchApp !== 'undefined');
        console.log("- CardService API: Available -", typeof CardService !== 'undefined');
        console.log("- PropertiesService API: Available -", typeof PropertiesService !== 'undefined');
      } catch (apiCheckError) {
        console.log("- API availability check error:", apiCheckError.message);
      }
      
      console.log("=== END DIAGNOSTIC INFO ===");
    } catch (diagError) {
      console.error("Error during diagnostic logging:", diagError.message);
    }
  }
  
  // Permission testing function - run this first if you get OAuth errors
  function debugTestPermissions() {
    console.log("=== OAUTH PERMISSIONS DEBUG TEST ===");
    console.log("Testing all required OAuth scopes...\n");
    
    var permissionResults = {
      gmail: false,
      drive: false,
      userinfo: false,
      urlfetch: false,
      total: 0,
      passed: 0
    };
    
    // Test Gmail permissions
    try {
      var quota = GmailApp.getRemainingDailyQuota();
      console.log("✅ Gmail API: PASSED (quota:", quota + ")");
      permissionResults.gmail = true;
      permissionResults.passed++;
    } catch (error) {
      console.log("❌ Gmail API: FAILED -", error.message);
    }
    permissionResults.total++;
    
    // Test Drive permissions
    try {
      var storage = DriveApp.getStorageUsed();
      console.log("✅ Drive API: PASSED (storage used:", storage, "bytes)");
      permissionResults.drive = true;
      permissionResults.passed++;
    } catch (error) {
      console.log("❌ Drive API: FAILED -", error.message);
    }
    permissionResults.total++;
    
    // Test user info permissions  
    try {
      var email = Session.getActiveUser().getEmail();
      console.log("✅ User Info API: PASSED (email:", email + ")");
      permissionResults.userinfo = true;
      permissionResults.passed++;
    } catch (error) {
      console.log("❌ User Info API: FAILED -", error.message);
      console.log("   Fix: Ensure 'https://www.googleapis.com/auth/userinfo.email' is in OAuth scopes");
    }
    permissionResults.total++;
    
    // Test external requests (for PMO webhook)
    try {
      // Simple test request to check if external requests are allowed
      if (typeof UrlFetchApp !== 'undefined') {
        console.log("✅ External Requests: PASSED (UrlFetchApp available)");
        permissionResults.urlfetch = true;
        permissionResults.passed++;
      } else {
        console.log("❌ External Requests: FAILED (UrlFetchApp not available)");
      }
    } catch (error) {
      console.log("❌ External Requests: FAILED -", error.message);
    }
    permissionResults.total++;
    
    console.log("\n=== PERMISSION TEST SUMMARY ===");
    console.log("Passed:", permissionResults.passed, "of", permissionResults.total);
    if (permissionResults.passed === permissionResults.total) {
      console.log("🎉 ALL PERMISSIONS OK - Ready for PMO testing");
    } else {
      console.log("⚠️  SOME PERMISSIONS MISSING - Check OAuth scopes in appsscript.json");
      console.log("Required scopes:");
      console.log("- https://www.googleapis.com/auth/gmail.readonly");
      console.log("- https://www.googleapis.com/auth/drive");  
      console.log("- https://www.googleapis.com/auth/userinfo.email");
      console.log("- https://www.googleapis.com/auth/script.external_request");
    }
    
    return permissionResults;
  }
  
  // Test alternative PMO URLs to diagnose DNS issues
  function debugTestPMOUrls() {
    console.log("=== TESTING ALTERNATIVE PMO URLS ===");
    
    var testUrls = [
      "https://n8n-pmo.office.transporeon.com/webhook/ad028ac7-647f-48a8-ba0c-f259d8671299",
      "https://httpbin.org/post", // Public test endpoint
      "https://www.google.com",   // Basic connectivity test
      "https://transporeon.com"   // Company main domain test
    ];
    
    for (var i = 0; i < testUrls.length; i++) {
      var url = testUrls[i];
      console.log("\nTesting URL:", url);
      
      try {
        var startTime = new Date().getTime();
        var response = UrlFetchApp.fetch(url, {
          method: 'GET',
          muteHttpExceptions: true,
          timeout: 5000
        });
        var duration = new Date().getTime() - startTime;
        
        console.log("✅ SUCCESS - Response code:", response.getResponseCode());
        console.log("- Duration:", duration + "ms");
        console.log("- Headers:", Object.keys(response.getAllHeaders()).length, "headers received");
        
      } catch (error) {
        var duration = new Date().getTime() - startTime;
        console.log("❌ FAILED - Error:", error.message);
        console.log("- Duration:", duration + "ms");
        console.log("- Error type:", error.name);
        
        if (error.message.includes("DNS")) {
          console.log("- Issue: DNS resolution failed - hostname not reachable from Google's network");
        } else if (error.message.includes("timeout")) {
          console.log("- Issue: Network timeout - server may be slow or blocking requests");
        }
      }
    }
    
    console.log("\n=== URL TEST ANALYSIS ===");
    console.log("If Google/httpbin work but PMO fails → Corporate firewall blocking external access");  
    console.log("If all URLs fail → Google Apps Script network issue");
    console.log("If transporeon.com works but n8n-pmo subdomain fails → Internal-only subdomain");
    
    return "URL connectivity test completed - check logs above";
  }
  
  // Simple test to check if Google Apps Script can reach external sites
  function debugTestBasicConnectivity() {
    console.log("=== BASIC CONNECTIVITY TEST ===");
    
    try {
      console.log("Testing basic internet connectivity from Google Apps Script...");
      
      var response = UrlFetchApp.fetch("https://httpbin.org/get", {
        method: 'GET',
        timeout: 5000,
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        console.log("✅ INTERNET CONNECTIVITY: WORKING");
        console.log("Google Apps Script can reach external websites");
        console.log("The PMO DNS error is specific to n8n-pmo.office.transporeon.com");
        console.log("\n🔍 This confirms the issue is:");
        console.log("• Corporate firewall blocking external access to PMO server");
        console.log("• Internal-only hostname not resolvable from external networks");
        console.log("• PMO server configured for internal access only");
        return true;
      } else {
        console.log("⚠️ INTERNET CONNECTIVITY: LIMITED - HTTP " + response.getResponseCode());
        return false;
      }
      
    } catch (error) {
      console.log("❌ INTERNET CONNECTIVITY: FAILED");
      console.log("Error:", error.message);
      console.log("Google Apps Script may have network restrictions");
      return false;
    }
  }
  
  // Manual testing function for PMO connectivity (can be run directly from Apps Script editor)
  function debugTestPMOConnectivity() {
    console.log("=== MANUAL PMO CONNECTIVITY DEBUG TEST ===");
    
    try {
      // Test permissions first
      console.log("Step 1: Testing OAuth permissions...");
      var permissionTest = debugTestPermissions();
      
      if (permissionTest.passed !== permissionTest.total) {
        console.log("⚠️  Cannot proceed with PMO test - permission issues detected");
        console.log("Fix permissions first, then run debugTestPMOConnectivity() again");
        return { success: false, error: "OAuth permissions not configured" };
      }
      
      console.log("\nStep 2: Testing URL connectivity...");  
      debugTestPMOUrls();
      
      console.log("\nStep 3: Running system diagnostics...");
      logDiagnosticInfo();
      
      console.log("\nStep 4: Testing PMO webhook connectivity...");
      console.log("=== TESTING PMO WEBHOOK CONNECTIVITY ===");
      
      // Test with a debug ticket
      var testTicket = "CXPRODELIVERY-DEBUG-" + new Date().getTime();
      console.log("Testing with ticket:", testTicket);
      
      var result = getPMOProjectFolder(testTicket);
      
      console.log("\n=== PMO TEST RESULTS ===");
      console.log("Success:", result.success);
      if (result.success) {
        console.log("Folder ID:", result.folderId);
        console.log("Created:", result.created);
        
        // Test folder access
        console.log("\n=== TESTING FOLDER ACCESS ===");
        var folderResult = getFolderByIdSafely(result.folderId);
        console.log("Folder access success:", folderResult.success);
        if (folderResult.success) {
          console.log("Folder name:", folderResult.name);
        } else {
          console.log("Folder access error:", folderResult.error);
        }
      } else {
        console.log("PMO Error:", result.error);
      }
      
      console.log("=== END MANUAL PMO TEST ===");
      return result;
      
    } catch (error) {
      console.error("Manual PMO test failed:", error.message);
      console.error("Error stack:", error.stack);
      return { success: false, error: error.message };
    }
  }
  
  function getPMOErrorMessage(error, ticketKey) {
    var baseMessage = "❌ Cannot save attachments for " + ticketKey;
    
    if (error.includes("DNS error")) {
      return baseMessage + "\n\n🌐 DNS Resolution Error:\n" + 
             "• PMO server hostname cannot be resolved from Google's network\n" +
             "• This suggests the PMO server (n8n-pmo.office.transporeon.com) is behind corporate firewall\n" +
             "• External access may be blocked for security reasons\n\n" +
             "🔧 Next Steps:\n" +
             "• Contact Transporeon IT/PMO team about external access\n" +
             "• Verify the PMO webhook URL is correct\n" +
             "• Check if an alternative public endpoint exists\n\n" +
             "💡 This is an infrastructure issue, not a configuration problem";
    }
    
    if (error.includes("timeout") || error.includes("network")) {
      return baseMessage + "\n\n🌐 Network Issue:\n" + 
             "• PMO system is temporarily unreachable\n" +
             "• Please check your internet connection\n" +
             "• Try again in a few moments\n\n" +
             "💡 If problem persists, contact IT support";
    } 
    
    if (error.includes("HTTP 404") || error.includes("HTTP 500")) {
      return baseMessage + "\n\n🔧 PMO System Issue:\n" + 
             "• PMO webhook is temporarily unavailable\n" +
             "• This is likely a temporary service issue\n" +
             "• Please try again in a few minutes\n\n" +
             "💡 If problem persists, contact PMO team";
    }
    
    if (error.includes("HTTP 403") || error.includes("unauthorized")) {
      return baseMessage + "\n\n🔐 Access Issue:\n" + 
             "• You may not have permission to create project folders\n" +
             "• Check with your project manager\n" +
             "• Verify the ticket number is correct\n\n" +
             "💡 Contact PMO team if you should have access";
    }
    
    if (error.includes("undefined folder ID")) {
      return baseMessage + "\n\n⏱️ PMO Folder Creation Issue:\n" + 
             "• PMO system is processing your project folder\n" +
             "• Folder creation took longer than expected\n" +
             "• Please try again in a moment\n\n" +
             "💡 The folder may be created now - retry to check";
    }
    
    if (error.includes("invalid response format")) {
      return baseMessage + "\n\n📋 PMO Response Issue:\n" + 
             "• PMO system returned an unexpected response\n" +
             "• This may be a temporary issue\n" +
             "• Please try again\n\n" +
             "💡 If error continues, contact IT support";
    }
    
    // Generic error message
    return baseMessage + "\n\n⚠️ PMO Integration Error:\n" +
           "• " + error + "\n" +
           "• Please try again in a few moments\n\n" +
           "💡 If problem persists, contact IT or PMO support";
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
  
  /**
   * Get appropriate icon for Google Docs type
   * @param {string} docType - The type of Google Doc (docs, sheets, slides, etc.)
   * @returns {string} Emoji icon for the doc type
   */
  function getGoogleDocsIcon(docType) {
    var icons = {
      'docs': '📄',      // Document icon
      'sheets': '📊',    // Spreadsheet icon  
      'slides': '📽️',    // Presentation icon
      'forms': '📝',     // Form icon
      'drive': '📁',     // Generic drive file icon
      'drawings': '🎨'   // Drawing icon
    };
    
    return icons[docType] || '🔗'; // Default link icon for unknown types
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
  
  // ===== SUBFOLDER MANAGEMENT FOR PMO PROJECTS =====
  
  /**
   * Get standardized project subfolder configuration
   * @returns {Array} Array of subfolder definitions
   */
  function getProjectSubfolderConfig() {
    return [
      {
        name: "01_System_Design",
        icon: "📁",
        description: "System architecture, diagrams, technical specifications"
      },
      {
        name: "02_Meet_Recordings", 
        icon: "📁",
        description: "Meeting recordings, session notes, call summaries"
      },
      {
        name: "03_Correspondence",
        icon: "📁", 
        description: "Email threads, communications, external correspondence"
      },
      {
        name: "04_Project_Documentation",
        icon: "📁",
        description: "General project documents, reports, deliverables"
      }
    ];
  }
  
  /**
   * Create or access project subfolder with support for nested paths
   * @param {Folder} parentFolder - The PMO project root folder
   * @param {string} subfolderPath - Path like "01_System_Design" or "04_Project_Documentation/Project_Management"
   * @returns {Object} Result object with folder info
   */
  function getOrCreateProjectSubfolder(parentFolder, subfolderPath) {
    try {
      console.log("=== SUBFOLDER CREATION/ACCESS ===");
      console.log("Parent folder:", parentFolder.getName(), "ID:", parentFolder.getId());
      console.log("Requested subfolder path:", subfolderPath);
      
      if (!subfolderPath || subfolderPath.trim() === "") {
        // Return parent folder for root level saves
        console.log("No subfolder specified, using project root");
        return {
          success: true,
          folder: parentFolder,
          path: "Project Root",
          created: false
        };
      }
      
      // Split path for nested folders (e.g., "04_Project_Documentation/Project_Management")
      var pathParts = subfolderPath.split('/');
      var currentFolder = parentFolder;
      var createdFolders = [];
      var fullPath = "";
      
      for (var i = 0; i < pathParts.length; i++) {
        var folderName = pathParts[i].trim();
        if (!folderName) continue;
        
        fullPath += (fullPath ? "/" : "") + folderName;
        console.log("Processing folder level", (i + 1) + ":", folderName);
        
        // Check if folder already exists
        var existingFolders = currentFolder.getFoldersByName(folderName);
        
        if (existingFolders.hasNext()) {
          currentFolder = existingFolders.next();
          console.log("✓ Found existing folder:", folderName);
          
          // Check for duplicates (multiple folders with same name)
          if (existingFolders.hasNext()) {
            console.warn("⚠ WARNING: Multiple folders found with name:", folderName);
          }
        } else {
          // Create new folder
          currentFolder = currentFolder.createFolder(folderName);
          createdFolders.push(folderName);
          console.log("✓ Created new folder:", folderName);
        }
      }
      
      console.log("=== SUBFOLDER RESULT ===");
      console.log("Final path:", fullPath);
      console.log("Created folders:", createdFolders.length > 0 ? createdFolders.join(", ") : "None");
      console.log("Final folder ID:", currentFolder.getId());
      
      return {
        success: true,
        folder: currentFolder,
        path: fullPath,
        created: createdFolders.length > 0,
        createdFolders: createdFolders
      };
      
    } catch (error) {
      console.error("=== SUBFOLDER CREATION ERROR ===");
      console.error("Error:", error.message);
      console.error("Stack:", error.stack);
      
      return {
        success: false,
        error: "Failed to create/access subfolder: " + error.message,
        path: subfolderPath
      };
    }
  }
  
  /**
   * Validate subfolder selection against allowed paths
   * @param {string} subfolderPath - The subfolder path to validate
   * @returns {Object} Validation result
   */
  function validateSubfolderSelection(subfolderPath) {
    // Define allowed subfolder patterns
    var allowedPaths = [
      "", // Project root
      "01_System_Design",
      "02_Meet_Recordings", 
      "03_Correspondence",
      "04_Project_Documentation",
      "04_Project_Documentation/Project_Management",
      "04_Project_Documentation/Carrier_Onboarding"
    ];
    
    if (!subfolderPath || subfolderPath.trim() === "") {
      return { valid: true, path: "", sanitized: "" };
    }
    
    // Check against allowed paths
    if (allowedPaths.indexOf(subfolderPath) !== -1) {
      return { 
        valid: true, 
        path: subfolderPath,
        sanitized: subfolderPath
      };
    }
    
    // Sanitize path - remove invalid characters
    var sanitized = subfolderPath
      .replace(/[<>:"|?*]/g, '_') // Replace invalid chars
      .replace(/\.\./g, '_') // Prevent directory traversal
      .replace(/^\/+|\/+$/g, '') // Remove leading/trailing slashes
      .substring(0, 100); // Limit length
    
    console.log("⚠ Subfolder path sanitized:", subfolderPath, "→", sanitized);
    
    return {
      valid: false,
      path: subfolderPath,
      sanitized: sanitized,
      warning: "Path was not in predefined list and was sanitized"
    };
  }
  
  /**
   * Handle subfolder creation failure with graceful degradation
   * @param {Folder} projectFolder - The project root folder
   * @param {string} subfolderPath - The failed subfolder path
   * @param {Object} errorDetails - Error information
   * @returns {Object} Fallback result
   */
  function handleSubfolderCreationFailure(projectFolder, subfolderPath, errorDetails) {
    console.log("=== SUBFOLDER FALLBACK STRATEGY ===");
    console.log("Original path:", subfolderPath);
    console.log("Error:", errorDetails.error);
    
    // Strategy 1: Try simplified folder name (remove nesting, use only first level)
    if (subfolderPath.indexOf("/") !== -1) {
      var simplifiedPath = subfolderPath.split("/")[0]; // Use only first level
      console.log("Trying simplified path:", simplifiedPath);
      
      var fallbackResult = getOrCreateProjectSubfolder(projectFolder, simplifiedPath);
      if (fallbackResult.success) {
        console.log("✓ Fallback successful with simplified path");
        return {
          success: true,
          folder: fallbackResult.folder,
          fallbackUsed: "simplified_path",
          originalPath: subfolderPath,
          actualPath: simplifiedPath,
          warning: "Used simplified folder path due to nesting issue"
        };
      }
    }
    
    // Strategy 2: Fall back to project root
    console.log("All subfolder creation failed, using project root");
    return {
      success: true,
      folder: projectFolder,
      fallbackUsed: "project_root",
      originalPath: subfolderPath,
      actualPath: "Project Root",
      warning: "Saved to project root due to subfolder creation error"
    };
  }
  
  /**
   * Get user preferences for subfolder feature
   * @returns {Object} User preferences
   */
  function getUserSubfolderPreferences() {
    var userProperties = PropertiesService.getUserProperties();
    var preferences = {
      enableSubfolderSelection: userProperties.getProperty('ENABLE_SUBFOLDER_SELECTION') !== 'false', // Default true
      defaultSubfolder: userProperties.getProperty('DEFAULT_SUBFOLDER') || '',
      showSubfolderTooltips: userProperties.getProperty('SHOW_SUBFOLDER_TOOLTIPS') !== 'false'
    };
    
    return preferences;
  }
  
  /**
   * Check if enhanced UI should be used
   * @returns {boolean} True if enhanced UI should be shown
   */
  function shouldUseEnhancedUI() {
    try {
      var preferences = getUserSubfolderPreferences();
      return preferences.enableSubfolderSelection;
    } catch (error) {
      console.log("Error loading preferences, defaulting to enhanced UI:", error);
      return true; // Default to enhanced UI
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
  
  // ===== GOOGLE DOCS LINK EXTRACTION FUNCTIONS =====
  
  /**
   * Extract Google Docs links from email message content
   * @param {GmailMessage} message - The Gmail message to analyze
   * @returns {Array} Array of Google Docs link objects
   */
  function extractGoogleDocsLinks(message) {
    try {
      console.log("Extracting Google Docs links from message...");
      
      var links = [];
      var htmlBody = '';
      var plainBody = '';
      
      // Get message content
      try {
        htmlBody = message.getBody() || '';
        plainBody = message.getPlainBody() || '';
      } catch (e) {
        console.log("Could not get message body:", e.message);
        return links;
      }
      
      // Define Google service URL patterns
      var patterns = [
        // Google Docs
        {
          regex: /https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'docs',
          defaultTitle: 'Google Doc'
        },
        // Google Sheets  
        {
          regex: /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'sheets',
          defaultTitle: 'Google Sheet'
        },
        // Google Slides
        {
          regex: /https:\/\/docs\.google\.com\/presentation\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'slides', 
          defaultTitle: 'Google Slides'
        },
        // Google Forms
        {
          regex: /https:\/\/docs\.google\.com\/forms\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'forms',
          defaultTitle: 'Google Form'
        },
        // Google Drive files (generic)
        {
          regex: /https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9-_]+)/g,
          type: 'drive',
          defaultTitle: 'Google Drive File'
        }
      ];
      
      // Search in both HTML and plain text
      var contentToSearch = [
        { content: htmlBody, isHtml: true },
        { content: plainBody, isHtml: false }
      ];
      
      var foundUrls = new Set(); // Prevent duplicates
      
      contentToSearch.forEach(function(contentObj) {
        patterns.forEach(function(pattern) {
          var matches;
          pattern.regex.lastIndex = 0; // Reset regex state
          
          while ((matches = pattern.regex.exec(contentObj.content)) !== null) {
            var fullUrl = matches[0];
            var fileId = matches[1];
            
            // Skip if we already found this URL
            if (foundUrls.has(fullUrl)) {
              continue;
            }
            foundUrls.add(fullUrl);
            
            // Try to extract title from HTML context
            var title = extractTitleFromContext(contentObj.content, fullUrl, contentObj.isHtml);
            if (!title) {
              title = pattern.defaultTitle;
            }
            
            // Ensure unique title by adding file ID if needed
            var finalTitle = title;
            var titleExists = links.some(function(link) { return link.title === finalTitle; });
            if (titleExists) {
              finalTitle = title + " (" + fileId.substring(0, 8) + "...)";
            }
            
            links.push({
              url: fullUrl,
              fileId: fileId,
              title: finalTitle,
              type: pattern.type
            });
            
            console.log("Found Google " + pattern.type + " link:", finalTitle, "->", fullUrl);
          }
        });
      });
      
      console.log("Extracted", links.length, "Google Docs links from message");
      return links;
      
    } catch (error) {
      console.error("Error extracting Google Docs links:", error.message);
      return [];
    }
  }
  
  /**
   * Try to extract title from the context around a Google Docs URL
   * @param {string} content - The content to search in
   * @param {string} url - The URL to find context for
   * @param {boolean} isHtml - Whether the content is HTML
   * @returns {string|null} Extracted title or null
   */
  function extractTitleFromContext(content, url, isHtml) {
    try {
      var urlIndex = content.indexOf(url);
      if (urlIndex === -1) return null;
      
      if (isHtml) {
        // For HTML content, look for link text or nearby text
        var beforeUrl = content.substring(Math.max(0, urlIndex - 200), urlIndex);
        var afterUrl = content.substring(urlIndex + url.length, Math.min(content.length, urlIndex + url.length + 200));
        
        // Look for <a> tag with text
        var linkRegex = /<a[^>]*href[^>]*>([^<]+)<\/a>/i;
        var contextRegex = />([^<]{2,50})<.*?href.*?google\.com/i;
        var reverseContextRegex = /href.*?google\.com.*?>([^<]{2,50})</i;
        
        var match = beforeUrl.match(contextRegex) || afterUrl.match(reverseContextRegex) || 
                   (beforeUrl + url + afterUrl).match(linkRegex);
        
        if (match && match[1]) {
          var title = match[1].trim();
          // Clean up common HTML artifacts
          title = title.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
          if (title.length > 3 && title.length < 100) {
            return title;
          }
        }
      } else {
        // For plain text, look for text near the URL
        var beforeUrl = content.substring(Math.max(0, urlIndex - 100), urlIndex);
        var afterUrl = content.substring(urlIndex + url.length, Math.min(content.length, urlIndex + url.length + 100));
        
        // Look for text patterns that might be titles
        var patterns = [
          /([^\n\r]{5,80})\s*$/,  // Text before URL on same line
          /^\s*([^\n\r]{5,80})/   // Text after URL on same line
        ];
        
        for (var i = 0; i < patterns.length; i++) {
          var match = (i === 0 ? beforeUrl : afterUrl).match(patterns[i]);
          if (match && match[1]) {
            var title = match[1].trim();
            // Skip if it looks like part of a URL or email
            if (!title.includes('http') && !title.includes('@') && title.length < 80) {
              return title;
            }
          }
        }
      }
      
      return null;
      
    } catch (error) {
      console.error("Error extracting title from context:", error.message);
      return null;
    }
  }
  
  /**
   * Copy a Google Doc to the target folder while preserving its format
   * @param {Object} docLink - The Google Docs link object
   * @param {Folder} targetFolder - The folder to copy the Google Doc to
   * @returns {File} The copied Google Doc file
   */
  function copyGoogleDocToFolder(docLink, targetFolder) {
    try {
      console.log("Copying Google Doc to target folder:", docLink.title);
      
      // Extract document ID from Google Docs URL
      var docId = extractGoogleDocId(docLink.url);
      if (!docId) {
        console.error("Could not extract document ID from URL:", docLink.url);
        // Fallback to shortcut creation if ID extraction fails
        return createFallbackShortcut(docLink, targetFolder);
      }
      
      console.log("Extracted document ID:", docId);
      
      // Get the Google Doc file from Drive
      var driveFile;
      try {
        driveFile = DriveApp.getFileById(docId);
      } catch (driveError) {
        console.error("Could not access Google Drive file:", driveError.message);
        // Fallback to shortcut creation if file access fails
        return createFallbackShortcut(docLink, targetFolder);
      }
      
      // Use email subject as filename if available, otherwise fall back to original name
      var fileName;
      if (docLink.emailSubject && docLink.emailSubject.trim() !== '') {
        fileName = sanitizeFileName(docLink.emailSubject);
        console.log("Using email subject as filename:", fileName);
      } else {
        fileName = docLink.title || driveFile.getName();
        console.log("Using original Google Doc name:", fileName);
      }
      
      // Check if a file with the same name already exists in the target folder
      var existingFiles = targetFolder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        var existingFile = existingFiles.next();
        
        // Check if it's the same file (same ID)
        if (existingFile.getId() === docId) {
          console.log("Google Doc already exists in target folder:", fileName);
          return existingFile;
        }
        
        // Different file with same name - add timestamp to avoid conflict
        var timestamp = new Date().getTime();
        var nameParts = fileName.split('.');
        if (nameParts.length > 1) {
          fileName = nameParts.slice(0, -1).join('.') + '_' + timestamp + '.' + nameParts[nameParts.length - 1];
        } else {
          fileName = fileName + '_' + timestamp;
        }
        console.log("Renamed to avoid conflict:", fileName);
      }
      
      // Create a copy of the Google Doc in the target folder
      var copiedFile;
      try {
        copiedFile = driveFile.makeCopy(fileName, targetFolder);
        console.log("Successfully copied Google Doc to target folder as:", fileName);
        return copiedFile;
        
      } catch (copyError) {
        console.error("Could not copy Google Doc:", copyError.message);
        // Fallback to shortcut creation if copy fails
        return createFallbackShortcut(docLink, targetFolder);
      }
      
    } catch (error) {
      console.error("Error copying Google Doc:", error.message);
      // Fallback to shortcut creation
      return createFallbackShortcut(docLink, targetFolder);
    }
  }

  // Helper function to sanitize email subject for use as filename
  function sanitizeFileName(subject) {
    try {
      if (!subject || typeof subject !== 'string') {
        return 'Google Doc';
      }
      
      // Remove or replace characters that are not allowed in filenames
      var sanitized = subject
        .replace(/[<>:"/\\|?*]/g, '_')  // Replace invalid filename characters
        .replace(/\s+/g, ' ')          // Replace multiple spaces with single space
        .trim();                       // Remove leading/trailing whitespace
      
      // Limit length to reasonable filename size
      if (sanitized.length > 100) {
        sanitized = sanitized.substring(0, 100).trim();
      }
      
      // Ensure we have a non-empty filename
      if (sanitized === '') {
        return 'Google Doc';
      }
      
      console.log("Sanitized filename:", sanitized);
      return sanitized;
      
    } catch (error) {
      console.error("Error sanitizing filename:", error.message);
      return 'Google Doc';
    }
  }

  // Helper function to extract Google Doc ID from URL
  function extractGoogleDocId(url) {
    try {
      // Handle different Google Docs URL formats
      var patterns = [
        /\/document\/d\/([a-zA-Z0-9-_]+)/,  // Documents
        /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/, // Sheets  
        /\/presentation\/d\/([a-zA-Z0-9-_]+)/, // Slides
        /\/file\/d\/([a-zA-Z0-9-_]+)/,        // Generic Drive files
        /id=([a-zA-Z0-9-_]+)/                 // ID parameter format
      ];
      
      for (var i = 0; i < patterns.length; i++) {
        var match = url.match(patterns[i]);
        if (match && match[1]) {
          return match[1];
        }
      }
      
      console.log("No matching pattern found for URL:", url);
      return null;
    } catch (error) {
      console.error("Error extracting Google Doc ID:", error.message);
      return null;
    }
  }


  // Fallback function to create shortcut if download fails
  function createFallbackShortcut(docLink, targetFolder) {
    try {
      console.log("Creating fallback shortcut for:", docLink.title);
      
      // Create a shortcut file content pointing to the Google Doc
      var shortcutContent = '[InternetShortcut]\nURL=' + docLink.url + '\n';
      
      // Use email subject as filename if available, otherwise use original title
      var baseName;
      if (docLink.emailSubject && docLink.emailSubject.trim() !== '') {
        baseName = sanitizeFileName(docLink.emailSubject);
        console.log("Using email subject for shortcut filename:", baseName);
      } else {
        baseName = docLink.title || 'Google Doc';
        console.log("Using original title for shortcut filename:", baseName);
      }
      
      var fileName = baseName + '.url';
      
      // Handle duplicate names
      var existingFiles = targetFolder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        var timestamp = new Date().getTime();
        var nameParts = fileName.split('.');
        if (nameParts.length > 1) {
          fileName = nameParts.slice(0, -1).join('.') + '_' + timestamp + '.' + nameParts[nameParts.length - 1];
        } else {
          fileName = fileName + '_' + timestamp;
        }
      }
      
      // Create the shortcut file
      var blob = Utilities.newBlob(shortcutContent, 'text/plain', fileName);
      var file = targetFolder.createFile(blob);
      
      console.log("Created fallback shortcut file:", fileName);
      return file;
      
    } catch (error) {
      console.error("Error creating fallback shortcut:", error.message);
      throw error;
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