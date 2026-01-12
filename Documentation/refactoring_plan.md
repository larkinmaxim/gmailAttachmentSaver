# Code.js Refactoring Plan

## Executive Summary

This document provides a detailed, step-by-step plan to refactor the 5,030-line Code.js file into a well-structured, maintainable codebase. The refactoring is divided into 4 priority phases with estimated timelines and clear success criteria.

**Total Estimated Time:** 8-10 days
**Expected Benefits:**
- Maintainability: +400%
- Testability: +500%
- Development Speed: +30% (post-refactoring)
- Bug Reduction: -60%

---

## Phase 1: Critical Refactoring (Priority 1)
**Estimated Time:** 2-3 days
**Goal:** Fix architectural issues and reduce immediate technical debt

### Task 1.1: Set Up Module Structure
**Time:** 2 hours
**Goal:** Create the foundation for modular architecture

#### Steps:
1. Create a new directory structure:
   ```
   src/
   ├── modules/
   │   ├── ConfigurationService.js
   │   ├── LoggingService.js
   │   ├── FolderService.js
   │   ├── JiraService.js
   │   ├── UIService.js
   │   ├── AttachmentService.js
   │   └── BigQueryService.js
   ├── utils/
   │   ├── Constants.js
   │   ├── ErrorHandler.js
   │   └── UIHelpers.js
   └── Code.js (entry point)
   ```

2. Create empty module files with JSDoc headers:
   ```javascript
   /**
    * @fileoverview [Module Name] - [Brief Description]
    * @author Gmail Attachment Saver Team
    * @version 2.0.0
    */
   ```

3. Set up a build script (optional, for concatenation if needed):
   - Create `build.js` to concatenate modules into single Code.js for Apps Script
   - Or document manual concatenation process

#### Success Criteria:
- [ ] All 7 module files created
- [ ] Directory structure exists
- [ ] JSDoc headers added

---

### Task 1.2: Extract ConfigurationService Module
**Time:** 3 hours
**Goal:** Centralize all configuration and script properties

#### Steps:

1. **Create Constants.js first** (needed by ConfigurationService):
   ```javascript
   /**
    * @fileoverview Application-wide constants and configuration values
    */

   // === JIRA CONFIGURATION ===
   var JIRA = {
     DEFAULT_URL: 'https://support.transporeon.com',
     DEFAULT_PROJECT_PREFIX: 'CXPRODELIVERY',
     TEST_TICKET: 'CXPRODELIVERY-4750',

     ACTIVE_STATUSES: [
       'HYPERCARE',
       'HYPERCARE (WITH CHECK)',
       'Order received',
       'Test system available',
       'Project go-live/productive start',
       'System Design Assigned',
       'Implementation Assigned',
       'Implementation Started',
       'Requirements Clarified',
       'Handover Check Needed',
       'LIVE SYSTEM AVAILABLE'
     ],

     /**
      * Builds the default JQL query for fetching TPM projects
      * @return {string} JQL query string
      */
     buildDefaultJQL: function() {
       return [
         'project = ' + this.DEFAULT_PROJECT_PREFIX,
         'issuetype in (Project, "Project (Standard Solution)")',
         'status in (' + this.ACTIVE_STATUSES.map(function(s) {
           return '"' + s + '"';
         }).join(', ') + ')',
         '"Technical Project Manager" in (currentUser())'
       ].join(' AND ');
     }
   };

   // === UI LIMITS ===
   var UI_LIMITS = {
     ATTACHMENTS_PER_GROUP: 10,
     TOTAL_ATTACHMENTS_DISPLAY: 50,
     MAX_FOLDER_NAME_LENGTH: 200,
     MAX_FILE_NAME_LENGTH: 100,
     MAX_DROPDOWN_ITEMS: 100
   };

   // === CACHE CONFIGURATION ===
   var CACHE = {
     FOLDER_LOOKUP_SECONDS: 21600,  // 6 hours
     ATTACHMENT_SELECTIONS_SECONDS: 86400,  // 24 hours

     /**
      * Generates cache key for customer folder lookup
      * @param {string} customerId - Customer ID
      * @return {string} Cache key
      */
     getFolderCacheKey: function(customerId) {
       return 'customer_' + customerId;
     },

     /**
      * Generates cache key for project folder lookup
      * @param {string} ticketKey - Jira ticket key
      * @return {string} Cache key
      */
     getProjectCacheKey: function(ticketKey) {
       return 'project_' + ticketKey;
     }
   };

   // === FOLDER STRUCTURE ===
   var FOLDER_STRUCTURE = {
     PROJECT_SUBFOLDERS: [
       "01_System_Design",
       "02_Meet_Recordings",
       "03_Correspondence",
       "04_Project_Documentation"
     ],

     DOC_SUBFOLDERS: [
       "Project_Management",
       "Custom_Bundle",
       "Carrier_Onboarding"
     ],

     /**
      * Gets the parent folder name for documentation subfolders
      * @return {string} Parent folder name
      */
     getDocParentFolder: function() {
       return this.PROJECT_SUBFOLDERS[3]; // "04_Project_Documentation"
     }
   };

   // === USER-FACING MESSAGES ===
   var MESSAGES = {
     HINTS: {
       TICKET_NUMBER: "e.g., 6500 (will become " + JIRA.DEFAULT_PROJECT_PREFIX + "-6500)",
       FULL_TICKET: "e.g., " + JIRA.DEFAULT_PROJECT_PREFIX + "-6310"
     },

     ERRORS: {
       NO_SETTINGS: "Please configure your Jira settings first",
       NO_TICKET: "Please select a ticket or enter a manual ticket number",
       NO_GMAIL_CONTEXT: "Please open an email to save attachments",
       NO_ATTACHMENTS: "No attachments found in this thread",
       NO_SELECTION: "Please select at least one attachment",
       JIRA_CONNECTION_FAILED: "Failed to connect to Jira. Please check your settings.",
       FOLDER_NOT_FOUND: "Could not find the specified folder",
       FOLDER_ACCESS_DENIED: "Access denied to folder. Please check permissions."
     },

     SUCCESS: {
       SETTINGS_SAVED: "Settings saved successfully",
       ATTACHMENTS_SAVED: "Attachments saved successfully",
       FOLDER_CREATED: "Folder structure created successfully"
     }
   };

   // === LOG LEVELS ===
   var LOG_LEVELS = {
     ERROR: 0,
     WARN: 1,
     INFO: 2,
     DEBUG: 3
   };

   // === API TIMEOUTS ===
   var TIMEOUTS = {
     URL_FETCH_DEFAULT: 5000,      // 5 seconds
     JIRA_API: 10000,               // 10 seconds
     DRIVE_API: 30000,              // 30 seconds
     BIGQUERY_API: 15000            // 15 seconds
   };
   ```

2. **Create ConfigurationService.js**:
   Extract lines 1-253 from Code.js containing:
   - `setupScriptProperties()`
   - All `getXxx()` configuration functions
   - `setServiceAccountKey()`, `setServiceAccountEmail()`

   Template:
   ```javascript
   /**
    * @fileoverview Configuration and Script Properties management
    */

   /**
    * Sets up default script properties (one-time setup)
    * @return {Object} Result object with success status
    */
   function setupScriptProperties() {
     // Move code from Code.js lines 15-136
   }

   /**
    * Gets the default Jira URL from script properties
    * @return {string} Jira URL
    */
   function getDefaultJiraUrl() {
     // Move code from Code.js lines 138-144
     // Update to use JIRA.DEFAULT_URL constant
   }

   // ... continue for all config functions
   ```

3. **Update Code.js references**:
   - Replace `'https://support.transporeon.com'` with `JIRA.DEFAULT_URL`
   - Replace `'CXPRODELIVERY-'` with `JIRA.DEFAULT_PROJECT_PREFIX + '-'`
   - Replace magic numbers with constants from `UI_LIMITS`, `CACHE`, `TIMEOUTS`

#### Success Criteria:
- [ ] Constants.js contains all magic strings and numbers
- [ ] ConfigurationService.js contains all config functions (lines 1-253)
- [ ] No hardcoded URLs, project names, or limits in Code.js
- [ ] All functions reference constants correctly

---

### Task 1.3: Extract and Standardize LoggingService
**Time:** 2 hours
**Goal:** Create consistent logging infrastructure

#### Steps:

1. **Create LoggingService.js**:
   ```javascript
   /**
    * @fileoverview Centralized logging service with configurable levels
    */

   /**
    * Gets the current log level from user properties
    * @return {number} Current log level (0-3)
    */
   function getLogLevel() {
     // Move from Code.js lines 254-266
   }

   /**
    * Sets the log level in user properties
    * @param {number} level - Log level to set (0-3)
    */
   function setLogLevel(level) {
     // Move from Code.js lines 268-282
   }

   /**
    * Logs an error message (always visible)
    * @param {string} message - Error message
    * @param {Error} [error] - Optional error object
    */
   function logError(message, error) {
     var fullMessage = "❌ " + message;
     if (error) {
       fullMessage += ": " + error.message;
       if (error.stack) {
         console.error(fullMessage);
         console.error("Stack trace:", error.stack);
       } else {
         console.error(fullMessage);
       }
     } else {
       console.error(fullMessage);
     }
   }

   /**
    * Logs a warning message (visible at WARN level and above)
    * @param {string} message - Warning message
    */
   function logWarn(message) {
     if (getLogLevel() >= LOG_LEVELS.WARN) {
       console.log("⚠️ " + message);
     }
   }

   /**
    * Logs an info message (visible at INFO level and above)
    * @param {string} message - Info message
    */
   function logInfo(message) {
     if (getLogLevel() >= LOG_LEVELS.INFO) {
       console.log("ℹ️ " + message);
     }
   }

   /**
    * Logs a debug message (visible at DEBUG level only)
    * @param {string} message - Debug message
    */
   function logDebug(message) {
     if (getLogLevel() >= LOG_LEVELS.DEBUG) {
       console.log("🔍 " + message);
     }
   }

   /**
    * Enhanced verbose logging for debugging
    * @param {string} message - Message to log
    * @param {*} data - Optional data to stringify and log
    */
   function verboseLog(message, data) {
     // Move from Code.js lines 444-461
   }

   /**
    * Gets the name of a log level
    * @param {number} level - Log level number
    * @return {string} Log level name
    */
   function getLogLevelName(level) {
     // Move from Code.js lines 430-442
   }
   ```

2. **Replace all console.log/error calls throughout Code.js**:
   - `console.error("Error...")` → `logError("Error...")`
   - `console.log("Warning...")` → `logWarn("Warning...")`
   - `console.log("Info...")` → `logInfo("Info...")`
   - `console.log("Debug...")` → `logDebug("Debug...")`

3. **Create a script to automate replacement** (optional):
   ```javascript
   // replace-logging.js
   var patterns = [
     { find: /console\.error\("❌ ([^"]+)"\)/g, replace: 'logError("$1")' },
     { find: /console\.log\("⚠️ ([^"]+)"\)/g, replace: 'logWarn("$1")' },
     { find: /console\.log\("ℹ️ ([^"]+)"\)/g, replace: 'logInfo("$1")' },
     { find: /console\.log\("🔍 ([^"]+)"\)/g, replace: 'logDebug("$1")' }
   ];
   ```

#### Success Criteria:
- [ ] LoggingService.js created with all logging functions
- [ ] All 544 console.log/error calls replaced with logging functions
- [ ] Error logging includes optional error object parameter
- [ ] Log levels work correctly (tested with test script)

---

### Task 1.4: Create ErrorHandler Utility
**Time:** 3 hours
**Goal:** Standardize error handling across the application

#### Steps:

1. **Create ErrorHandler.js**:
   ```javascript
   /**
    * @fileoverview Centralized error handling utilities
    */

   /**
    * Error codes for consistent error handling
    */
   var ERROR_CODES = {
     // Configuration errors
     CONFIG_MISSING: 'CONFIG_MISSING',
     CONFIG_INVALID: 'CONFIG_INVALID',

     // Jira errors
     JIRA_CONNECTION_FAILED: 'JIRA_CONNECTION_FAILED',
     JIRA_AUTH_FAILED: 'JIRA_AUTH_FAILED',
     JIRA_TICKET_NOT_FOUND: 'JIRA_TICKET_NOT_FOUND',
     JIRA_PARSE_ERROR: 'JIRA_PARSE_ERROR',

     // Drive errors
     DRIVE_FOLDER_NOT_FOUND: 'DRIVE_FOLDER_NOT_FOUND',
     DRIVE_ACCESS_DENIED: 'DRIVE_ACCESS_DENIED',
     DRIVE_UPLOAD_FAILED: 'DRIVE_UPLOAD_FAILED',

     // Gmail errors
     GMAIL_NO_CONTEXT: 'GMAIL_NO_CONTEXT',
     GMAIL_NO_ATTACHMENTS: 'GMAIL_NO_ATTACHMENTS',
     GMAIL_FETCH_FAILED: 'GMAIL_FETCH_FAILED',

     // General errors
     VALIDATION_FAILED: 'VALIDATION_FAILED',
     UNKNOWN_ERROR: 'UNKNOWN_ERROR'
   };

   /**
    * Creates a standardized error result object
    * @param {string} code - Error code from ERROR_CODES
    * @param {string} message - User-friendly error message
    * @param {Object} [details] - Additional error details
    * @return {Object} Standardized error object
    */
   function createErrorResult(code, message, details) {
     logError(message, details ? details.error : null);

     return {
       success: false,
       error: {
         code: code,
         message: message,
         details: details || null,
         timestamp: new Date().toISOString()
       }
     };
   }

   /**
    * Creates a success result object
    * @param {*} data - Success data
    * @param {string} [message] - Optional success message
    * @return {Object} Standardized success object
    */
   function createSuccessResult(data, message) {
     if (message) {
       logInfo(message);
     }

     return {
       success: true,
       data: data,
       message: message || null,
       timestamp: new Date().toISOString()
     };
   }

   /**
    * Creates an error card for display in Gmail add-on
    * @param {string} functionName - Name of function where error occurred
    * @param {Error} error - Error object
    * @param {boolean} [showSettings] - Whether to show settings button
    * @return {Card} CardService card with error message
    */
   function createErrorCard(functionName, error, showSettings) {
     var card = CardService.newCardBuilder();
     var section = CardService.newCardSection()
       .setHeader("Error - " + functionName);

     section.addWidget(
       CardService.newTextParagraph()
         .setText("❌ " + error.message)
     );

     if (error.stack && getLogLevel() >= LOG_LEVELS.DEBUG) {
       section.addWidget(
         CardService.newTextParagraph()
           .setText("<font color='#666'><i>Stack trace logged to console</i></font>")
       );
     }

     // Add troubleshooting button
     if (showSettings !== false) {
       section.addWidget(
         CardService.newButtonSet()
           .addButton(
             CardService.newTextButton()
               .setText("⚙️ Check Settings")
               .setOnClickAction(
                 CardService.newAction().setFunctionName("showSettings")
               )
           )
       );
     }

     card.addSection(section);
     return card.build();
   }

   /**
    * Creates an error notification for action responses
    * @param {string} message - Error message to display
    * @return {ActionResponse} Action response with error notification
    */
   function createErrorNotification(message) {
     logError(message);

     return CardService.newActionResponseBuilder()
       .setNotification(
         CardService.newNotification()
           .setType(CardService.NotificationType.ERROR)
           .setText(message)
       )
       .build();
   }

   /**
    * Creates a success notification for action responses
    * @param {string} message - Success message to display
    * @return {ActionResponse} Action response with success notification
    */
   function createSuccessNotification(message) {
     logInfo(message);

     return CardService.newActionResponseBuilder()
       .setNotification(
         CardService.newNotification()
           .setType(CardService.NotificationType.INFO)
           .setText(message)
       )
       .build();
   }

   /**
    * Wraps a function with error boundary
    * @param {Function} fn - Function to wrap
    * @param {Function} [fallbackFn] - Optional fallback function
    * @return {Function} Wrapped function with error handling
    */
   function withErrorBoundary(fn, fallbackFn) {
     return function() {
       try {
         return fn.apply(this, arguments);
       } catch (error) {
         logError("Error in " + fn.name, error);

         if (fallbackFn) {
           return fallbackFn(error);
         }

         // Default fallback for card-building functions
         if (fn.name.startsWith('build') || fn.name.startsWith('show')) {
           return [createErrorCard(fn.name, error)];
         }

         // Re-throw for non-UI functions
         throw error;
       }
     };
   }
   ```

2. **Replace error handling patterns in Code.js**:

   Find all instances of:
   ```javascript
   } catch (error) {
     console.error("Error in [functionName]:", error);
     return { success: false, error: error.message };
   }
   ```

   Replace with:
   ```javascript
   } catch (error) {
     return createErrorResult(ERROR_CODES.UNKNOWN_ERROR, error.message, { error: error });
   }
   ```

3. **Update specific error handlers**:
   - Jira connection errors → Use `ERROR_CODES.JIRA_CONNECTION_FAILED`
   - Folder not found → Use `ERROR_CODES.DRIVE_FOLDER_NOT_FOUND`
   - Missing settings → Use `ERROR_CODES.CONFIG_MISSING`

#### Success Criteria:
- [ ] ErrorHandler.js created with all utility functions
- [ ] ERROR_CODES enum defined with all error types
- [ ] All try-catch blocks use createErrorResult()
- [ ] All error cards use createErrorCard()
- [ ] All notifications use createErrorNotification()/createSuccessNotification()

---

### Task 1.5: Refactor buildAddOn() Function
**Time:** 4 hours
**Goal:** Break 457-line function into manageable pieces

#### Steps:

1. **Create UIHelpers.js** with reusable UI components:
   ```javascript
   /**
    * @fileoverview Reusable UI component builders for Gmail add-on
    */

   /**
    * Creates a settings button widget
    * @return {Widget} Settings button widget
    */
   function createSettingsButton() {
     return CardService.newButtonSet()
       .addButton(
         CardService.newTextButton()
           .setText("⚙️")
           .setOnClickAction(
             CardService.newAction().setFunctionName("showSettings")
           )
       );
   }

   /**
    * Creates a standard header section with settings button
    * @return {CardSection} Header section
    */
   function createHeaderSection() {
     var section = CardService.newCardSection();
     section.addWidget(createSettingsButton());
     return section;
   }

   /**
    * Creates a text paragraph widget with optional styling
    * @param {string} text - Text content
    * @param {Object} [options] - Styling options
    * @return {Widget} Text paragraph widget
    */
   function createTextParagraph(text, options) {
     options = options || {};

     var styledText = text;
     if (options.color) {
       styledText = '<font color="' + options.color + '">' + text + '</font>';
     }
     if (options.italic) {
       styledText = '<i>' + styledText + '</i>';
     }
     if (options.bold) {
       styledText = '<b>' + styledText + '</b>';
     }

     return CardService.newTextParagraph().setText(styledText);
   }

   /**
    * Creates a dropdown selection widget
    * @param {string} fieldName - Field name for form data
    * @param {Array} items - Array of {text, value, selected} objects
    * @param {string} [title] - Optional title
    * @return {Widget} Selection input widget
    */
   function createDropdown(fieldName, items, title) {
     var dropdown = CardService.newSelectionInput()
       .setType(CardService.SelectionInputType.DROPDOWN)
       .setFieldName(fieldName);

     if (title) {
       dropdown.setTitle(title);
     }

     items.forEach(function(item) {
       dropdown.addItem(item.text, item.value, item.selected || false);
     });

     return dropdown;
   }
   ```

2. **Extract ticket selection logic** from buildAddOn():
   ```javascript
   /**
    * Creates the ticket selection section
    * @param {Object} e - Event object
    * @return {CardSection} Ticket selection section
    */
   function createTicketSelectionSection(e) {
     logDebug("Creating ticket selection section");

     var section = CardService.newCardSection()
       .setHeader("Jira Project");

     try {
       var projects = getMyJiraProjects();

       if (!projects || projects.length === 0) {
         return createNoProjectsSection();
       }

       var ticketItems = formatTicketsForDropdown(projects);
       var dropdown = createDropdown('selectedTicket', ticketItems, 'Select Ticket:');
       section.addWidget(dropdown);

       // Add "Show Details" button
       section.addWidget(
         CardService.newButtonSet()
           .addButton(
             CardService.newTextButton()
               .setText("Show Details")
               .setOnClickAction(
                 CardService.newAction()
                   .setFunctionName("showTicketDetails")
                   .setLoadIndicator(CardService.LoadIndicator.SPINNER)
               )
           )
       );

     } catch (error) {
       return createJiraErrorSection(error);
     }

     return section;
   }

   /**
    * Formats Jira tickets for dropdown display
    * @param {Array} tickets - Array of Jira tickets
    * @return {Array} Formatted items for dropdown
    */
   function formatTicketsForDropdown(tickets) {
     return tickets.map(function(ticket) {
       var displayText = formatCompactTicketDisplay(ticket);
       return {
         text: displayText,
         value: ticket.key,
         selected: false
       };
     });
   }

   /**
    * Creates a fallback section when no projects are found
    * @return {CardSection} No projects section
    */
   function createNoProjectsSection() {
     var section = CardService.newCardSection()
       .setHeader("No Projects Found");

     section.addWidget(
       createTextParagraph(
         "No active projects found. Please check your Jira settings or permissions.",
         { color: '#666', italic: true }
       )
     );

     section.addWidget(createSettingsButton());

     return section;
   }

   /**
    * Creates an error section for Jira connection issues
    * @param {Error} error - Error object
    * @return {CardSection} Error section
    */
   function createJiraErrorSection(error) {
     var section = CardService.newCardSection()
       .setHeader("Jira Connection Error");

     section.addWidget(
       createTextParagraph(
         MESSAGES.ERRORS.JIRA_CONNECTION_FAILED + " " + error.message,
         { color: '#cc0000' }
       )
     );

     section.addWidget(createSettingsButton());

     return section;
   }
   ```

3. **Refactor main buildAddOn() function**:
   ```javascript
   /**
    * Main entry point for Gmail add-on
    * @param {Object} e - Event object from Gmail
    * @return {Array<Card>} Array of cards to display
    */
   function buildAddOn(e) {
     logInfo("=== BUILD ADD-ON CALLED ===");

     try {
       // Check if user has configured settings
       if (!isConfigured()) {
         return [buildInitialSetupCard()];
       }

       var card = CardService.newCardBuilder();

       // Add header with settings button
       card.addSection(createHeaderSection());

       // Add ticket selection section
       card.addSection(createTicketSelectionSection(e));

       // Add attachment sections if Gmail context exists
       if (hasGmailContext(e)) {
         var attachmentSections = createAttachmentSections(e);
         attachmentSections.forEach(function(section) {
           card.addSection(section);
         });
       }

       // Add save section
       card.addSection(createSaveSection(e));

       return [card.build()];

     } catch (error) {
       return [createErrorCard('buildAddOn', error)];
     }
   }

   /**
    * Checks if user has configured required settings
    * @return {boolean} True if configured
    */
   function isConfigured() {
     var settings = getUserSettings();
     return !!(settings.jiraToken && settings.jiraUrl);
   }

   /**
    * Checks if event has Gmail context
    * @param {Object} e - Event object
    * @return {boolean} True if Gmail context exists
    */
   function hasGmailContext(e) {
     return !!(e && e.gmail && e.gmail.threadId);
   }

   /**
    * Creates the initial setup card for first-time users
    * @return {Card} Setup card
    */
   function buildInitialSetupCard() {
     logInfo("User not configured, showing setup card");
     return buildSettingsCard(true);
   }
   ```

#### Success Criteria:
- [ ] buildAddOn() reduced from 457 lines to < 50 lines
- [ ] UIHelpers.js created with reusable components
- [ ] 5+ helper functions extracted (createTicketSelectionSection, etc.)
- [ ] All extracted functions have JSDoc comments
- [ ] buildAddOn() function is readable and maintainable

---

### Task 1.6: Refactor showTicketDetails() Function
**Time:** 5 hours
**Goal:** Break 530-line function into manageable pieces

#### Steps:

1. **Extract attachment section builder** (this is the complex part):
   Create `AttachmentSectionBuilder.js` in utils/:
   ```javascript
   /**
    * @fileoverview Builder for creating attachment selection sections
    */

   /**
    * Creates attachment selection sections from email thread
    * @param {Object} e - Event object with Gmail context
    * @return {Array<CardSection>} Array of attachment sections
    */
   function createAttachmentSections(e) {
     logDebug("Creating attachment sections");

     if (!hasGmailContext(e)) {
       return [createNoContextSection()];
     }

     try {
       var threadId = e.gmail.threadId;
       var messageId = e.gmail.messageId;
       var accessToken = e.gmail.accessToken;

       var attachments = getAllAttachmentsFromThread(threadId, messageId, accessToken);

       if (attachments.length === 0) {
         return [createNoAttachmentsSection()];
       }

       var builder = new AttachmentSectionBuilder(attachments, threadId);
       return builder.buildSections();

     } catch (error) {
       logError("Error creating attachment sections", error);
       return [createAttachmentErrorSection(error)];
     }
   }

   /**
    * Builder class for attachment sections
    * @param {Array} attachments - Array of attachments
    * @param {string} threadId - Gmail thread ID
    * @constructor
    */
   function AttachmentSectionBuilder(attachments, threadId) {
     this.attachments = attachments;
     this.threadId = threadId;
     this.sections = [];
     this.totalWidgets = 0;
     this.totalSkipped = 0;
   }

   /**
    * Builds all attachment sections
    * @return {Array<CardSection>} Array of sections
    */
   AttachmentSectionBuilder.prototype.buildSections = function() {
     var grouped = this.groupAttachmentsByExtension();
     var selections = getStoredAttachmentSelections(this.threadId);

     var sortedExtensions = Object.keys(grouped).sort();
     var isFirstSection = true;

     for (var i = 0; i < sortedExtensions.length; i++) {
       var ext = sortedExtensions[i];

       if (this.hasRemainingBudget()) {
         var section = this.buildSection(grouped[ext], selections, isFirstSection);
         this.sections.push(section);
         isFirstSection = false;
       } else {
         this.totalSkipped += grouped[ext].count;
       }
     }

     this.logSkippedAttachments();
     return this.sections;
   };

   /**
    * Groups attachments by file extension
    * @return {Object} Grouped attachments
    */
   AttachmentSectionBuilder.prototype.groupAttachmentsByExtension = function() {
     var groups = {};

     for (var i = 0; i < this.attachments.length; i++) {
       var attachment = this.attachments[i];
       var ext = this.getAttachmentExtension(attachment);

       if (!groups[ext]) {
         groups[ext] = {
           extension: ext,
           count: 0,
           items: []
         };
       }

       groups[ext].items.push({
         attachment: attachment,
         index: i
       });
       groups[ext].count++;
     }

     return groups;
   };

   /**
    * Gets file extension from attachment
    * @param {Object} attachment - Attachment object
    * @return {string} File extension
    */
   AttachmentSectionBuilder.prototype.getAttachmentExtension = function(attachment) {
     if (attachment.type === 'google_docs') {
       return 'google-' + attachment.docType;
     }

     var fileName = attachment.name;
     var lastDot = fileName.lastIndexOf('.');

     return lastDot > -1 ?
       fileName.substring(lastDot + 1).toLowerCase() :
       'no-extension';
   };

   /**
    * Builds a single attachment section
    * @param {Object} group - Attachment group
    * @param {Object} selections - Stored selections
    * @param {boolean} isFirst - Is this the first section
    * @return {CardSection} Built section
    */
   AttachmentSectionBuilder.prototype.buildSection = function(group, selections, isFirst) {
     var header = this.buildSectionHeader(group, isFirst);
     var section = CardService.newCardSection()
       .setHeader(header)
       .setCollapsible(true)
       .setNumUncollapsibleWidgets(Math.min(group.count, 2));

     var itemsToShow = this.calculateItemsToShow(group.items.length);

     for (var i = 0; i < itemsToShow; i++) {
       var checkbox = this.createAttachmentCheckbox(group.items[i], selections);
       section.addWidget(checkbox);
       this.totalWidgets++;
     }

     var skipped = group.items.length - itemsToShow;
     if (skipped > 0) {
       section.addWidget(this.createSkippedMessage(group.extension, skipped));
       this.totalSkipped += skipped;
     }

     return section;
   };

   /**
    * Builds section header text
    * @param {Object} group - Attachment group
    * @param {boolean} isFirst - Is this the first section
    * @return {string} Header text
    */
   AttachmentSectionBuilder.prototype.buildSectionHeader = function(group, isFirst) {
     var extName = group.extension === 'no-extension' ?
       'Files without extension' :
       group.extension.toUpperCase() + ' files';

     var header = extName + " (" + group.count + ")";

     if (isFirst) {
       header = "📎 Select Attachments & Docs: " + header +
                " (Total: " + this.attachments.length + ")";
     }

     return header;
   };

   /**
    * Calculates how many items to show in section
    * @param {number} groupSize - Total items in group
    * @return {number} Number of items to show
    */
   AttachmentSectionBuilder.prototype.calculateItemsToShow = function(groupSize) {
     var perGroupLimit = Math.min(groupSize, UI_LIMITS.ATTACHMENTS_PER_GROUP);
     var remaining = UI_LIMITS.TOTAL_ATTACHMENTS_DISPLAY - this.totalWidgets;
     return Math.min(perGroupLimit, remaining);
   };

   /**
    * Checks if there's remaining budget for more widgets
    * @return {boolean} True if can add more widgets
    */
   AttachmentSectionBuilder.prototype.hasRemainingBudget = function() {
     return this.totalWidgets < UI_LIMITS.TOTAL_ATTACHMENTS_DISPLAY;
   };

   /**
    * Creates checkbox widget for attachment
    * @param {Object} item - Attachment item
    * @param {Object} selections - Stored selections
    * @return {Widget} Checkbox widget
    */
   AttachmentSectionBuilder.prototype.createAttachmentCheckbox = function(item, selections) {
     var attachment = item.attachment;
     var uniqueKey = attachment.name + "_" + item.index;
     var isSelected = selections && selections.hasOwnProperty(uniqueKey) ?
       selections[uniqueKey] : false;

     var displayText = this.getAttachmentDisplayText(attachment);

     return CardService.newSelectionInput()
       .setType(CardService.SelectionInputType.CHECK_BOX)
       .setFieldName("attachment_" + item.index)
       .addItem(displayText, item.index.toString(), isSelected);
   };

   /**
    * Gets display text for attachment
    * @param {Object} attachment - Attachment object
    * @return {string} Display text
    */
   AttachmentSectionBuilder.prototype.getAttachmentDisplayText = function(attachment) {
     if (attachment.type === 'google_docs') {
       var icon = getGoogleDocsIcon(attachment.docType);
       var typeName = attachment.docType.charAt(0).toUpperCase() +
                      attachment.docType.slice(1);
       return icon + " " + attachment.name + " (Google " + typeName + ")";
     }

     return attachment.name + " (" + formatFileSize(attachment.size) + ")";
   };

   /**
    * Creates "skipped" message widget
    * @param {string} extension - File extension
    * @param {number} count - Number skipped
    * @return {Widget} Text paragraph widget
    */
   AttachmentSectionBuilder.prototype.createSkippedMessage = function(extension, count) {
     var message = "... and " + count + " more " + extension.toUpperCase() + " files";
     return createTextParagraph(message, { color: '#888888', italic: true });
   };

   /**
    * Logs skipped attachments info
    */
   AttachmentSectionBuilder.prototype.logSkippedAttachments = function() {
     if (this.totalSkipped > 0) {
       logInfo("Skipped " + this.totalSkipped + " attachments due to widget limits");
     }
   };

   /**
    * Creates section when no Gmail context available
    * @return {CardSection} No context section
    */
   function createNoContextSection() {
     var section = CardService.newCardSection()
       .setHeader("No Email Selected");

     section.addWidget(
       createTextParagraph(
         MESSAGES.ERRORS.NO_GMAIL_CONTEXT,
         { color: '#666', italic: true }
       )
     );

     return section;
   }

   /**
    * Creates section when no attachments found
    * @return {CardSection} No attachments section
    */
   function createNoAttachmentsSection() {
     var section = CardService.newCardSection()
       .setHeader("No Attachments Found");

     section.addWidget(
       createTextParagraph(
         MESSAGES.ERRORS.NO_ATTACHMENTS,
         { color: '#666', italic: true }
       )
     );

     return section;
   }

   /**
    * Creates error section for attachment processing
    * @param {Error} error - Error object
    * @return {CardSection} Error section
    */
   function createAttachmentErrorSection(error) {
     var section = CardService.newCardSection()
       .setHeader("Error Loading Attachments");

     section.addWidget(
       createTextParagraph(
         "Failed to load attachments: " + error.message,
         { color: '#cc0000' }
       )
     );

     return section;
   }
   ```

2. **Refactor main showTicketDetails() function**:
   ```javascript
   /**
    * Shows detailed view for selected Jira ticket
    * @param {Object} e - Event object
    * @return {ActionResponse} Navigation response
    */
   function showTicketDetails(e) {
     logInfo("=== SHOW TICKET DETAILS CALLED ===");

     try {
       var ticketKey = extractTicketKey(e);

       if (!ticketKey || ticketKey === "manual") {
         return rebuildMainCard(e);
       }

       var ticket = findTicketByKey(ticketKey);

       if (!ticket) {
         logWarn("Ticket not found: " + ticketKey);
         return CardService.newActionResponseBuilder()
           .setNavigation(CardService.newNavigation().popCard())
           .build();
       }

       var card = buildTicketDetailsCard(e, ticket);

       return CardService.newActionResponseBuilder()
         .setNavigation(CardService.newNavigation().pushCard(card))
         .build();

     } catch (error) {
       return CardService.newActionResponseBuilder()
         .setNotification(
           CardService.newNotification()
             .setType(CardService.NotificationType.ERROR)
             .setText("Error: " + error.message)
         )
         .build();
     }
   }

   /**
    * Builds the ticket details card
    * @param {Object} e - Event object
    * @param {Object} ticket - Jira ticket object
    * @return {Card} Ticket details card
    */
   function buildTicketDetailsCard(e, ticket) {
     var card = CardService.newCardBuilder();

     // Add header
     card.addSection(createHeaderSection());

     // Add ticket dropdown (same as main card)
     card.addSection(createTicketDropdownSection(e, ticket));

     // Add ticket details section
     card.addSection(createTicketInfoSection(ticket));

     // Add attachment sections if Gmail context exists
     if (hasGmailContext(e)) {
       var attachmentSections = createAttachmentSections(e);
       attachmentSections.forEach(function(section) {
         card.addSection(section);
       });
     }

     // Add subfolder selection section
     card.addSection(createSubfolderSelectionSection(e));

     // Add save button section
     card.addSection(createSaveSection(e));

     return card.build();
   }

   /**
    * Creates ticket info display section
    * @param {Object} ticket - Jira ticket object
    * @return {CardSection} Ticket info section
    */
   function createTicketInfoSection(ticket) {
     var section = CardService.newCardSection()
       .setHeader("📋 Ticket Details");

     // Ticket key and status
     var statusEmoji = getStatusEmoji(ticket.status);
     var statusText = statusEmoji + " " + ticket.key + " - " + ticket.status;
     section.addWidget(createTextParagraph(statusText, { bold: true }));

     // Summary
     if (ticket.summary) {
       section.addWidget(
         createTextParagraph(ticket.summary, { color: '#666' })
       );
     }

     // Link to Jira
     var settings = getUserSettings();
     var jiraUrl = settings.jiraUrl + "/browse/" + ticket.key;
     section.addWidget(
       CardService.newTextButton()
         .setText("View in Jira")
         .setOpenLink(CardService.newOpenLink().setUrl(jiraUrl))
     );

     return section;
   }
   ```

3. **Extract remaining helper functions**:
   - `extractTicketKey(e)`
   - `findTicketByKey(ticketKey)`
   - `rebuildMainCard(e)`
   - `createTicketDropdownSection(e, ticket)`
   - `createSubfolderSelectionSection(e)`
   - `createSaveSection(e)`

#### Success Criteria:
- [ ] showTicketDetails() reduced from 530 lines to < 50 lines
- [ ] AttachmentSectionBuilder.js created with builder pattern
- [ ] Attachment nesting reduced from 7 levels to 2 levels
- [ ] 8+ helper functions extracted
- [ ] All functions tested and working

---

## Phase 2: High Priority Refactoring (Priority 2)
**Estimated Time:** 2-3 days
**Goal:** Remove duplication and improve code quality

### Task 2.1: Extract Remaining Modules
**Time:** 4 hours per module (3 days total for 6 modules)

Follow the same pattern as Phase 1, extracting:
- **FolderService.js** - Lines 462-928 (folder operations)
- **JiraService.js** - Lines 3981-4061 + scattered functions
- **AttachmentService.js** - Lines 2618-3056 + 4063-4457
- **BigQueryService.js** - Lines 4459-4917

For each module:
1. Create the module file
2. Move functions with proper JSDoc
3. Update references in Code.js
4. Test functionality
5. Commit changes

### Task 2.2: Consolidate CardService Patterns
**Time:** 3 hours

1. Add to UIHelpers.js:
   - `createButton(text, functionName, options)`
   - `createButtonSet(buttons)`
   - `createSection(header, widgets, options)`
   - `createDropdownSection(header, fieldName, items)`

2. Replace all repetitive CardService calls throughout code

### Task 2.3: Refactor saveSelectedAttachmentsToGDrive()
**Time:** 4 hours

Break into:
- `validateSaveRequest(e)`
- `processAttachmentSelections(e)`
- `getFolderForTicket(ticket)`
- `saveAttachmentsToFolder(attachments, folder)`
- `createSaveNotification(result)`

---

## Phase 3: Medium Priority (Priority 3)
**Estimated Time:** 3-4 days
**Goal:** Improve function complexity and add tests

### Task 3.1: Reduce All Functions to < 50 Lines
**Time:** 4 hours

Target functions over 50 lines:
- `completeJiraToFolderWorkflowWithCreate()` - 189 lines
- All remaining functions > 50 lines

Apply Extract Method pattern systematically.

### Task 3.2: Add JSDoc Comments
**Time:** 4 hours

Add complete JSDoc to all functions:
```javascript
/**
 * Brief description
 * @param {Type} paramName - Description
 * @return {Type} Description
 * @throws {Error} When condition
 */
```

### Task 3.3: Set Up Testing Framework
**Time:** 1 day

1. Choose testing framework (QUnit for Apps Script)
2. Create test/ directory
3. Write tests for:
   - Configuration functions
   - Folder parsing logic
   - Attachment grouping
   - Error handling

### Task 3.4: Write Unit Tests
**Time:** 2 days

Target 70% code coverage:
- Test all business logic functions
- Mock external APIs (Jira, Drive, Gmail)
- Test error conditions

---

## Phase 4: Final Polish (Priority 4)
**Estimated Time:** 2-3 days
**Goal:** Performance, documentation, deployment

### Task 4.1: Performance Optimization
**Time:** 1 day

1. Profile slow functions (folder search, Jira API calls)
2. Optimize caching strategy
3. Reduce redundant API calls
4. Consider batch operations

### Task 4.2: Create Build System
**Time:** 4 hours

1. Create build script to concatenate modules
2. Add source maps for debugging
3. Automate deployment to Apps Script
4. Document build process

### Task 4.3: Update Documentation
**Time:** 1 day

1. Update README with new architecture
2. Create ARCHITECTURE.md explaining module structure
3. Add code examples for common tasks
4. Update deployment guide with build steps

### Task 4.4: Final Testing
**Time:** 4 hours

1. End-to-end testing in Gmail
2. Test all user workflows
3. Verify error handling
4. Performance benchmarks

---

## Success Metrics

### Code Quality Metrics
- [ ] File size: 5,030 lines → 7 files (~700 lines each)
- [ ] Max function length: 530 lines → 50 lines
- [ ] Cyclomatic complexity: 30 → <10
- [ ] Code duplication: 40% → <10%
- [ ] Test coverage: 0% → 70%

### Performance Metrics
- [ ] Add-on load time: Measure baseline → 20% improvement
- [ ] Folder search: Measure baseline → 30% improvement via caching
- [ ] Memory usage: Monitor and optimize

### Maintainability Metrics
- [ ] Time to find function: 5 min → 30 sec
- [ ] Time to understand function: 10 min → 2 min
- [ ] Time to add new feature: 4 hours → 2 hours

---

## Risk Management

### Potential Risks:
1. **Apps Script limitations** - May not support multiple files
   - **Mitigation:** Create build script to concatenate

2. **Breaking existing functionality** - Refactoring introduces bugs
   - **Mitigation:** Test thoroughly after each phase, maintain backup

3. **Time estimates too optimistic** - Tasks take longer than expected
   - **Mitigation:** Prioritize critical tasks, accept incremental progress

4. **Merge conflicts** - If active development continues
   - **Mitigation:** Work in feature branch, frequent merges

---

## Daily Progress Checklist

Use this checklist to track daily progress:

### Week 1: Phase 1 (Critical)
- [ ] Day 1: Tasks 1.1, 1.2 (Module structure + Configuration)
- [ ] Day 2: Tasks 1.3, 1.4 (Logging + Error handling)
- [ ] Day 3: Task 1.5 (Refactor buildAddOn)
- [ ] Day 4-5: Task 1.6 (Refactor showTicketDetails)

### Week 2: Phase 2 (High Priority)
- [ ] Day 6-7: Task 2.1 (Extract FolderService, JiraService)
- [ ] Day 8: Task 2.1 continued (AttachmentService, BigQueryService)
- [ ] Day 9: Tasks 2.2, 2.3 (UI patterns + saveAttachments)
- [ ] Day 10: Buffer/catch-up day

### Week 3: Phase 3 (Medium Priority)
- [ ] Day 11: Task 3.1 (Reduce function complexity)
- [ ] Day 12: Task 3.2 (Add JSDoc)
- [ ] Day 13: Task 3.3 (Set up testing)
- [ ] Day 14-15: Task 3.4 (Write tests)

### Optional Week 4: Phase 4 (Polish)
- [ ] Day 16: Task 4.1 (Performance)
- [ ] Day 17: Task 4.2 (Build system)
- [ ] Day 18: Task 4.3 (Documentation)
- [ ] Day 19: Task 4.4 (Final testing)
- [ ] Day 20: Buffer/deployment

---

## Rollback Plan

If issues arise:

1. **Each task should be a separate git commit**
   - Easy to revert individual changes

2. **Keep original Code.js as Code.original.js**
   - Quick rollback if needed

3. **Test after each phase**
   - Catch issues early before they compound

4. **Deploy to test environment first**
   - Don't touch production until fully tested

---

## Notes

- This is an aggressive timeline. Adjust based on your availability.
- Focus on Phase 1 first - it provides the most value
- Phases 2-4 can be done incrementally over time
- Don't skip testing - it will save time in the long run
- Document as you go - future you will thank you

---

## Getting Started

To begin:
1. Create a new branch: `git checkout -b refactor/modular-architecture`
2. Start with Task 1.1 (Set up module structure)
3. Work through Phase 1 systematically
4. Commit after each completed task
5. Test frequently

Good luck! 🚀
