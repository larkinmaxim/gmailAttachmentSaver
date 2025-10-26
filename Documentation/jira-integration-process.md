# Jira Integration Process Flow

## Overview
The Jira Integration connects to your Jira instance to fetch active Technical Project Manager tickets, enabling seamless ticket selection for attachment organization.

## Process Flow

### 1. Settings Configuration
```
User Opens Settings → Enter Credentials → Test Connection → Save Configuration
```

**Technical Steps:**
1. **Credential Input**
   - Jira URL validation (must start with http/https)
   - API token input (masked for security)
   - Custom JQL query configuration

2. **Connection Testing**
   ```javascript
   // Test API endpoint
   var testUrl = settings.jiraUrl + '/rest/api/2/myself';
   var response = UrlFetchApp.fetch(testUrl, {
     headers: { 'Authorization': 'Bearer ' + settings.jiraToken }
   });
   ```

3. **Settings Storage**
   - Secure storage in `PropertiesService.getUserProperties()`
   - Token masking for display security
   - Validation and error handling

### 2. Ticket Retrieval Process
```
Load Settings → Build JQL Query → Execute API Call → Process Results → Cache Data
```

**Technical Steps:**
1. **JQL Query Construction**
   ```javascript
   var jqlQuery = settings.customJql || getDefaultJQL();
   // Default: Focus on TPM tickets in active states
   ```

2. **API Request Execution**
   ```javascript
   var payload = {
     jql: jqlQuery,
     fields: ['project', 'key', 'summary', 'status', 'issuetype'],
     maxResults: 1000,
     startAt: 0
   };
   ```

3. **Response Processing**
   - Parse JSON response
   - Extract ticket information
   - Group by projects
   - Sort tickets alphabetically

### 3. Default JQL Logic
```
Project Filter → Issue Type Filter → Status Filter → Assignment Filter
```

**Default Query Components:**
```jql
project = CXPRODELIVERY 
AND issuetype in (Project, "Project (Standard Solution)")
AND status in (HYPERCARE, "Order received", "Test system available", ...)
AND "Technical Project Manager" in (currentUser())
```

**Status Categories:**
- **Active Development:** Implementation Started, System Design Started
- **Waiting States:** Order received, Test system available
- **Support Phases:** HYPERCARE, HYPERCARE (WITH CHECK)
- **Go-Live:** Project go-live/productive start, LIVE SYSTEM AVAILABLE

### 4. Ticket Display Process
```
Raw Ticket Data → Format Display → Add Status Emojis → Create Dropdown Options
```

**Technical Steps:**
1. **Display Text Formatting**
   ```javascript
   function formatCompactTicketDisplay(issue) {
     var statusEmoji = getStatusEmoji(issue.status);
     var clientInfo = extractClientFromSummary(issue.summary);
     return statusEmoji + " " + issue.key + " - " + clientInfo;
   }
   ```

2. **Status Emoji Mapping**
   - 🔧 HYPERCARE
   - 📋 Order received  
   - 🧪 Test system available
   - 🚀 Project go-live
   - 👤 System Design Assigned
   - 🔨 Implementation Started
   - ✅ Requirements Clarified
   - 🟢 LIVE SYSTEM AVAILABLE

3. **Client Name Extraction**
   - Parse summary field for client information
   - Format: `{number} - {client name} | {project details}`
   - Truncate long client names for dropdown display

### 5. Dynamic Ticket Details
```
Ticket Selection → Fetch Details → Display Information → Update UI Context
```

**Technical Steps:**
1. **Selection Processing**
   - Capture dropdown onChange event
   - Extract selected ticket key
   - Find ticket in cached results

2. **Detail Display**
   ```javascript
   var detailsText = statusEmoji + " **" + selectedTicket.key + "**\n\n";
   detailsText += "📝 **Summary:**\n" + selectedTicket.summary + "\n\n";
   detailsText += "📊 **Status:** " + selectedTicket.status + "\n";
   detailsText += "🏷️ **Type:** " + selectedTicket.issueType;
   ```

3. **UI Context Update**
   - Maintain dropdown state
   - Update attachment selection context
   - Preserve form inputs

## Error Handling

### Connection Failures
- **Authentication Errors:** Invalid API token or expired credentials
- **Network Issues:** Timeout or connectivity problems
- **Permission Errors:** Insufficient Jira permissions

**Recovery Strategies:**
```javascript
if (responseCode !== 200) {
  return {
    projects: {}, 
    issues: [], 
    totalIssues: 0,
    error: "Connection failed: " + responseCode
  };
}
```

### API Rate Limiting
- **Detection:** HTTP 429 responses
- **Mitigation:** Exponential backoff retry logic
- **User Feedback:** Clear error messages

### Data Processing Errors
- **Malformed Responses:** JSON parsing failures
- **Missing Fields:** Handle optional ticket fields gracefully
- **Large Result Sets:** Pagination support for >1000 tickets

## Security Considerations

### Credential Management
1. **Token Storage:** Encrypted in Google Apps Script properties
2. **Display Masking:** Show only first/last 4 characters
3. **Transmission:** HTTPS-only API calls with Bearer token auth

### Permission Scope
- **Minimal Access:** Only read permissions required
- **User Context:** API calls made as authenticated user
- **Audit Trail:** All API calls logged for troubleshooting

## Performance Optimizations

### Caching Strategy
1. **Ticket Caching:** Store results during session
2. **Project Mapping:** Cache project metadata
3. **Smart Refresh:** Only reload when settings change

### API Efficiency
1. **Field Selection:** Request only required fields
2. **Result Limiting:** Default 1000 tickets max
3. **Compression:** Accept gzipped responses

### User Experience
1. **Progressive Loading:** Show UI while loading tickets
2. **Error Recovery:** Graceful fallback to manual entry
3. **Connection Testing:** Validate settings before use
