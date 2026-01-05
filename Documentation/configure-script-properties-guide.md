# How to Configure Script Properties for BigQuery Logging

This step-by-step guide explains how to configure the Script Properties required for BigQuery meeting transcript logging.

## Prerequisites

Before starting, you need:
- Access to the Google Apps Script editor for the addon
- Your GCP project ID (e.g., `transporeon-gmail-addon-prod`)
- Service Account JSON key file (downloaded from GCP Console)

---

## Step 1: Open the Apps Script Editor

1. Go to [script.google.com](https://script.google.com)
2. Find and open your **Jira Attachment Saver** project
3. You should see the `Code.js` file in the editor

---

## Step 2: Set the BigQuery Project ID

### Option A: Using the Setup Function (Recommended for first-time setup)

1. In the Apps Script editor, find the function `setupScriptProperties()` at the top of `Code.js`

2. Locate this line and change the placeholder:
   ```javascript
   'BIGQUERY_PROJECT_ID': 'transporeon-gmail-addon-prod',
   ```
   Change to:
   ```javascript
   'BIGQUERY_PROJECT_ID': 'transporeon-gmail-addon-prod',
   ```

3. Click **Run** → Select `setupScriptProperties` → Click **Run**

4. Check the execution log for confirmation:
   ```
   ✅ Script properties configured successfully:
     - BIGQUERY_PROJECT_ID: transporeon-gmail-addon-prod
   ```

### Option B: Using the Project Settings UI

1. In Apps Script editor, click **⚙️ Project Settings** (gear icon in left sidebar)
2. Scroll down to **Script Properties**
3. Click **Edit script properties**
4. Click **+ Add script property**
5. Add:
   - **Property:** `BIGQUERY_PROJECT_ID`
   - **Value:** `transporeon-gmail-addon-prod`
6. Click **Save script properties**

---

## Step 3: Set the Service Account Email

1. Open your Service Account JSON key file in a text editor

2. Find the `client_email` field:
   ```json
   {
     "client_email": "gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com",
     ...
   }
   ```

3. In Apps Script editor, open the **Run** menu or use the function dropdown

4. Run the following in the editor (create a temporary function or use the execution area):
   ```javascript
   function configureServiceAccount() {
     setServiceAccountEmail('gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com');
   }
   ```

5. Or use the **Script Properties UI** method:
   - Go to **⚙️ Project Settings** → **Script Properties**
   - Add property:
     - **Property:** `SERVICE_ACCOUNT_EMAIL`
     - **Value:** `gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com`

---

## Step 4: Set the Service Account Private Key

This is the most important step. The private key authenticates the addon to BigQuery.

### 4.1 Extract the Private Key

1. Open your Service Account JSON key file

2. Find the `private_key` field - it looks like this:
   ```json
   {
     "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBg...[long string]...Kj8=\n-----END PRIVATE KEY-----\n",
     ...
   }
   ```

3. Copy the **entire value** including:
   - The `-----BEGIN PRIVATE KEY-----`
   - All the encoded content
   - The `-----END PRIVATE KEY-----`
   - The `\n` characters (these are important!)

### 4.2 Set the Key Using the Helper Function

1. In Apps Script editor, create and run this temporary function:

   ```javascript
   function setMyServiceAccountKey() {
     var privateKey = '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASC...[paste your full key here]...\n-----END PRIVATE KEY-----\n';
     
     setServiceAccountKey(privateKey);
   }
   ```

2. **Important:** Replace the example key with YOUR actual private key from the JSON file

3. Run the function

4. Check the log for:
   ```
   ✅ Service Account Private Key has been securely stored
      Run testBigQueryConnection() to verify the configuration
   ```

### 4.3 Alternative: Direct Script Properties Entry

1. Go to **⚙️ Project Settings** → **Script Properties**
2. Click **+ Add script property**
3. Add:
   - **Property:** `SERVICE_ACCOUNT_PRIVATE_KEY`
   - **Value:** Paste the entire private key (with `-----BEGIN/END-----` lines)
4. Click **Save script properties**

**Note:** When pasting into the UI, the `\n` characters should be converted to actual newlines.

---

## Step 5: Enable BigQuery Logging

1. Go to **⚙️ Project Settings** → **Script Properties**

2. Ensure this property exists and is set to `true`:
   - **Property:** `BIGQUERY_LOGGING_ENABLED`
   - **Value:** `true`

---

## Step 6: Verify All Properties Are Set

### Required Script Properties Checklist

| Property | Example Value | Status |
|----------|---------------|--------|
| `BIGQUERY_PROJECT_ID` | `transporeon-gmail-addon-prod` | ☐ |
| `BIGQUERY_DATASET_ID` | `addon_logs` | ☐ |
| `BIGQUERY_TABLE_ID` | `meeting_transcript_logs` | ☐ |
| `BIGQUERY_LOGGING_ENABLED` | `true` | ☐ |
| `SERVICE_ACCOUNT_EMAIL` | `gmail-addon-bigquery@...iam.gserviceaccount.com` | ☐ |
| `SERVICE_ACCOUNT_PRIVATE_KEY` | `-----BEGIN PRIVATE KEY-----...` | ☐ |

### View All Properties

1. Go to **⚙️ Project Settings** → **Script Properties**
2. All properties should be listed there

Or run this function to check:
```javascript
function checkBigQueryConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  
  console.log('=== BigQuery Configuration ===');
  console.log('BIGQUERY_PROJECT_ID:', props['BIGQUERY_PROJECT_ID'] || 'NOT SET');
  console.log('BIGQUERY_DATASET_ID:', props['BIGQUERY_DATASET_ID'] || 'NOT SET');
  console.log('BIGQUERY_TABLE_ID:', props['BIGQUERY_TABLE_ID'] || 'NOT SET');
  console.log('BIGQUERY_LOGGING_ENABLED:', props['BIGQUERY_LOGGING_ENABLED'] || 'NOT SET');
  console.log('SERVICE_ACCOUNT_EMAIL:', props['SERVICE_ACCOUNT_EMAIL'] || 'NOT SET');
  console.log('SERVICE_ACCOUNT_PRIVATE_KEY:', props['SERVICE_ACCOUNT_PRIVATE_KEY'] ? '[SET - ' + props['SERVICE_ACCOUNT_PRIVATE_KEY'].length + ' chars]' : 'NOT SET');
}
```

---

## Step 7: Test the Configuration

Run the test function to verify everything works:

```javascript
testBigQueryConnection();
```

### Expected Successful Output

```
=== TESTING BIGQUERY CONNECTION (Service Account) ===

Configuration:
  - Project ID: transporeon-gmail-addon-prod
  - Dataset ID: addon_logs
  - Table ID: meeting_transcript_logs
  - Service Account: gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com
  - Private Key Set: Yes
  - Logging Enabled: true

Testing Service Account authentication...
✅ Service Account authenticated successfully

Testing table access...
✅ Successfully connected to BigQuery table
   Table: transporeon-gmail-addon-prod.addon_logs.meeting_transcript_logs
```

---

## Troubleshooting

### "BigQuery logging skipped - not enabled or not configured"

Run `checkBigQueryConfig()` to see which properties are missing.

### "Service Account authentication failed"

1. Verify the private key is complete (should be ~2,600+ characters)
2. Check that `\n` characters are preserved (or converted to actual newlines)
3. Run `resetBigQueryAuth()` to clear cached tokens
4. Re-run `setServiceAccountKey()` with the correct key

### "Table not found"

The BigQuery table doesn't exist. Create it using the SQL in `bigquery-setup.md`.

### Private Key Format Issues

The private key should:
- Start with `-----BEGIN PRIVATE KEY-----`
- End with `-----END PRIVATE KEY-----`
- Contain line breaks (either `\n` or actual newlines)

Example of correct format:
```
-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC7...
...multiple lines of base64 encoded content...
...Kj8xYz==
-----END PRIVATE KEY-----
```

---

## Security Notes

1. **Never share the private key** - it grants full access to write to BigQuery
2. **Delete the JSON key file** after configuring Script Properties
3. **Script Properties are secure** - only script editors can view them, not addon users
4. **Rotate keys periodically** - create new keys in GCP Console and update Script Properties

---

## Quick Setup Commands

For experienced users, run these in sequence:

```javascript
// 1. Set service account email
setServiceAccountEmail('gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com');

// 2. Set private key (paste your actual key)
setServiceAccountKey('-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n');

// 3. Test connection
testBigQueryConnection();
```

