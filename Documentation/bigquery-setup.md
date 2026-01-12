# BigQuery Setup for Meeting Transcript Logging

This document describes how to set up BigQuery with Service Account authentication to log meeting transcript saves from the Gmail addon.

## Architecture Overview

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Gmail Addon    │────▶│ Service Account  │────▶│    BigQuery     │
│  (User Action)  │     │  (OAuth2 Auth)   │     │  (Data Insert)  │
└─────────────────┘     └──────────────────┘     └─────────────────┘
```

**Key Benefits of Service Account approach:**
- Users don't need individual BigQuery permissions
- All logs are written under a single service identity
- Centralized access control and audit trail
- No additional OAuth prompts for end users

## Prerequisites

- Google Cloud Platform (GCP) project with billing enabled
- BigQuery API enabled in your GCP project
- IAM Admin permissions to create service accounts
- Access to Google Apps Script editor

---

## Step 1: Create a Service Account

1. Go to [GCP Console → IAM & Admin → Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)

2. Click **"+ CREATE SERVICE ACCOUNT"**

3. Fill in the details:
   - **Name:** `gmail-addon-bigquery`
   - **ID:** `gmail-addon-bigquery` (auto-generated)
   - **Description:** `Service account for Gmail addon BigQuery logging`

4. Click **"CREATE AND CONTINUE"**

5. Grant the role:
   - **Role:** `BigQuery Data Editor`
   - Click **"CONTINUE"**

6. Click **"DONE"**

## Step 2: Create and Download the Service Account Key

1. Find your new service account in the list
2. Click on the service account email
3. Go to **"KEYS"** tab
4. Click **"ADD KEY" → "Create new key"**
5. Select **"JSON"** format
6. Click **"CREATE"**
7. **Save the downloaded JSON file securely** - you'll need it in Step 5

The JSON file contains:
```json
{
  "type": "service_account",
  "project_id": "transporeon-gmail-addon-prod",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n",
  "client_email": "gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com",
  ...
}
```

## Step 3: Create the BigQuery Dataset

Run this command in Cloud Shell or use the BigQuery Console:

```sql
-- Create the dataset (run once)
CREATE SCHEMA IF NOT EXISTS `transporeon-gmail-addon-prod.addon_logs`
OPTIONS (
  description = 'Logs from Gmail Addon for meeting transcript tracking',
  location = 'EU'  -- Change to your preferred region (US, EU, etc.)
);
```

Or via `bq` command line:

```bash
bq mk --location=EU --dataset transporeon-gmail-addon-prod:addon_logs
```

## Step 4: Create the Meeting Transcript Logs Table

Run this SQL in BigQuery Console:

```sql
CREATE TABLE IF NOT EXISTS `transporeon-gmail-addon-prod.addon_logs.meeting_transcript_logs` (
  event_id STRING NOT NULL,
  timestamp TIMESTAMP NOT NULL,
  user_email STRING NOT NULL,
  jira_ticket STRING NOT NULL,
  meeting_name STRING NOT NULL,
  gdrive_folder_id STRING NOT NULL,
  gdrive_file_id STRING NOT NULL,
  direct_link STRING NOT NULL,
  file_type STRING,
  source_email_subject STRING
)
PARTITION BY DATE(timestamp)
CLUSTER BY jira_ticket, user_email
OPTIONS (
  description = 'Logs of meeting transcripts saved via Gmail addon to 02_Meet_Recordings folder',
  labels = [('app', 'gmail_addon'), ('type', 'audit_log')]
);
```

### Table Schema Details

| Column | Type | Description |
|--------|------|-------------|
| `event_id` | STRING | Unique identifier for each save event (UUID) |
| `timestamp` | TIMESTAMP | When the save occurred (ISO 8601 format) |
| `user_email` | STRING | Email of the user who saved the file |
| `jira_ticket` | STRING | Jira project ticket (e.g., CXPRODELIVERY-1234) |
| `meeting_name` | STRING | Name of the saved meeting transcript file |
| `gdrive_folder_id` | STRING | Google Drive folder ID where file was saved |
| `gdrive_file_id` | STRING | Google Drive file ID of the saved file |
| `direct_link` | STRING | Direct URL to open the file |
| `file_type` | STRING | Type of Google file (docs, sheets, slides, etc.) |
| `source_email_subject` | STRING | Subject line of the email containing the transcript link |

## Step 5: Configure Script Properties

Open your service account JSON key file and extract the required values.

### 5.1 Set the Service Account Email

In Google Apps Script editor, run:

```javascript
// Copy the "client_email" value from your JSON key file
setServiceAccountEmail('gmail-addon-bigquery@transporeon-gmail-addon-prod.iam.gserviceaccount.com');
```

### 5.2 Set the Private Key

In Google Apps Script editor, run:

```javascript
// Copy the entire "private_key" value from your JSON key file
// Include the -----BEGIN PRIVATE KEY----- and -----END PRIVATE KEY----- lines
setServiceAccountKey('-----BEGIN PRIVATE KEY-----\nMIIEvQIBAD...your-key-here...\n-----END PRIVATE KEY-----\n');
```

**Security Note:** The private key is stored in Script Properties, which are only accessible to script editors, not end users.

### 5.3 Set BigQuery Configuration

Run `setupScriptProperties()` or manually set these in Script Properties:

| Property | Value | Description |
|----------|-------|-------------|
| `BIGQUERY_PROJECT_ID` | `transporeon-gmail-addon-prod` | Your GCP project ID |
| `BIGQUERY_DATASET_ID` | `addon_logs` | Dataset name (default) |
| `BIGQUERY_TABLE_ID` | `meeting_transcript_logs` | Table name (default) |
| `BIGQUERY_LOGGING_ENABLED` | `true` | Enable logging |

## Step 6: Test the Configuration

Run the test function in Google Apps Script editor:

```javascript
// Run this function to test BigQuery connectivity
testBigQueryConnection();
```

Expected output for successful configuration:

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

## Sample Queries

### View Recent Transcript Saves

```sql
SELECT 
  timestamp,
  jira_ticket,
  meeting_name,
  direct_link,
  source_email_subject
FROM `transporeon-gmail-addon-prod.addon_logs.meeting_transcript_logs`
ORDER BY timestamp DESC
LIMIT 100;
```

### Count Saves by Jira Project

```sql
SELECT 
  jira_ticket,
  COUNT(*) as transcript_count,
  MIN(timestamp) as first_save,
  MAX(timestamp) as last_save
FROM `transporeon-gmail-addon-prod.addon_logs.meeting_transcript_logs`
GROUP BY jira_ticket
ORDER BY transcript_count DESC;
```

### Daily Save Activity

```sql
SELECT 
  DATE(timestamp) as save_date,
  COUNT(*) as saves,
  COUNT(DISTINCT jira_ticket) as unique_projects
FROM `transporeon-gmail-addon-prod.addon_logs.meeting_transcript_logs`
WHERE timestamp >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 30 DAY)
GROUP BY save_date
ORDER BY save_date DESC;
```

---

## Troubleshooting

### Error: "BigQuery logging skipped - not enabled or not configured"

Checklist:
- [ ] `BIGQUERY_PROJECT_ID` is set to your actual GCP project ID (not placeholder)
- [ ] `BIGQUERY_LOGGING_ENABLED` is set to `'true'`
- [ ] `SERVICE_ACCOUNT_EMAIL` is set correctly
- [ ] `SERVICE_ACCOUNT_PRIVATE_KEY` is set (run `setServiceAccountKey()`)

### Error: "Service Account authentication failed"

- Verify the private key is complete (including BEGIN/END lines)
- Check the service account email matches your GCP project
- Ensure the service account exists and is not disabled
- Try running `resetBigQueryAuth()` to clear token cache

### Error: "Table not found" (404)

- Ensure the dataset and table were created correctly
- Verify the `BIGQUERY_DATASET_ID` and `BIGQUERY_TABLE_ID` match your BigQuery setup
- Check the service account has access to the dataset

### Error: "Permission denied" (403)

- Grant `BigQuery Data Editor` role to the service account
- Verify the role is applied at dataset or project level

### Reset Authentication

If you need to force re-authentication:

```javascript
resetBigQueryAuth();
```

---

## Security Best Practices

1. **Protect the Service Account Key**
   - Never commit the JSON key file to version control
   - Delete the JSON file after configuring Script Properties
   - Rotate keys periodically via GCP Console

2. **Minimal Permissions**
   - Service account only has `BigQuery Data Editor` role
   - Role is scoped to the specific dataset, not entire project

3. **Audit Trail**
   - All inserts are logged with the service account identity
   - BigQuery audit logs show all data access

---

## Cost Considerations

BigQuery streaming inserts are charged per successfully inserted row:
- Current pricing: $0.01 per 200 MB of data inserted
- Meeting transcript logs are small (~500 bytes per row)
- Typical usage: Very minimal cost (likely less than $1/month for thousands of saves)

Storage costs:
- First 10 GB per month is free
- Beyond that: ~$0.02 per GB per month
- Partitioning by date helps manage storage and query costs

---

## Quick Reference: Setup Functions

| Function | Purpose |
|----------|---------|
| `setServiceAccountEmail(email)` | Store service account email |
| `setServiceAccountKey(privateKey)` | Store service account private key |
| `testBigQueryConnection()` | Test full configuration |
| `resetBigQueryAuth()` | Clear OAuth token cache |
| `setupScriptProperties()` | Set all default properties |
