# Google Sheets Email Automation

A Google Apps Script that automatically sends personalized emails from a Gmail draft to recipients listed in Google Sheets.

## Features

* Automatic email sending when rows are added or pasted
* Gmail draft-based email template
* Dynamic `{{name}}` personalization
* Supports multiple email addresses per row
* HTML and plain-text email support
* Sends draft attachments
* Supports additional Google Drive attachments
* Automatically uses CC and BCC from the draft
* Supports signatures included in the draft
* Email format validation
* Duplicate email prevention
* Sent, Failed, Partial, and Duplicate status tracking
* Continues processing when individual emails fail
* Gmail daily sending-quota checking
* Configurable delay between emails
* Lock protection against simultaneous executions
* Automatically creates Email, Name, and Status headers
* Processes only rows with blank Status
* Logs sending and error information
* Advanced Gmail API support for reading draft CC/BCC

## Setup

1. Create a Google Sheet with **Email, Name, Status** columns.
2. Open **Extensions → Apps Script**.
3. Paste the `Code.gs` code.
4. Update `DRAFT_SUBJECT` and `SENDER_NAME`.
5. Create a Gmail draft with the matching subject.
6. Add your email content, `{{name}}`, signature, CC/BCC, and attachments if needed.
7. Run `createTrigger()` once and authorize the required permissions.
8. Add or paste recipient rows into the sheet.

The script automatically processes rows with a blank **Status** and updates the status after sending.
