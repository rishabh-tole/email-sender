# Google Sheets Email Sender Framework

A reusable, non-technical-friendly email sending system built on **Google Sheets + Google Apps Script**.

This tool lets a team:

- store reusable HTML email templates
- paste recipient lists into a sheet
- personalize emails with placeholders
- test drafts safely
- send test emails to a dummy inbox
- send full campaigns
- attach files from Google Drive
- track sends, drafts, errors, and logs

The system is designed to be easy for normal users while still being structured enough to scale to multiple campaigns.

---

# Table of Contents

1. [Overview](#overview)
2. [How to Use](#how-to-use)
   - [Initial Setup](#initial-setup)
   - [Daily Workflow](#daily-workflow)
   - [Testing Workflow](#testing-workflow)
   - [Full Send Workflow](#full-send-workflow)
   - [Attachments Workflow](#attachments-workflow)
   - [Common User Tasks](#common-user-tasks)
3. [How It Works](#how-it-works)
   - [Tab Structure](#tab-structure)
   - [Core Logic](#core-logic)
   - [Eligibility Logic](#eligibility-logic)
   - [Template Rendering](#template-rendering)
   - [Greeting Logic](#greeting-logic)
   - [Draft and Send Logic](#draft-and-send-logic)
   - [Logging](#logging)
   - [Statuses and Tracking Fields](#statuses-and-tracking-fields)
4. [Sheet Structure](#sheet-structure)
   - [MAIN DASHBOARD](#main-dashboard)
   - [HTML TEMPLATES](#html-templates)
   - [RECIPIENTS](#recipients)
   - [LOGS](#logs)
5. [Placeholder Reference](#placeholder-reference)
6. [Menu Actions](#menu-actions)
7. [Recommended Best Practices](#recommended-best-practices)
8. [Troubleshooting](#troubleshooting)
9. [Future Improvements](#future-improvements)

---

# Overview

This project turns a Google Sheet into a lightweight email campaign tool.

It is meant to be:

- **easy to use**
- **safe to test**
- **reusable**
- **structured**
- **transparent**

It avoids the need for non-technical users to edit code or understand email automation systems.

Instead, users mostly interact with four tabs:

- `MAIN DASHBOARD`
- `HTML TEMPLATES`
- `RECIPIENTS`
- `LOGS`

And a custom menu:

- `Load Active Template`
- `Validate Recipients`
- `Preview Top Recipient`
- `TEST DRAFT`
- `TEST SEND`
- `SEND FULL`
- `Refresh Dashboard Stats`
- `Reset Processing Fields`

---

# How to Use

## Initial Setup

This section explains how to set the tool up for the first time.

### Step 1: Create the spreadsheet

Create a Google Sheet with exactly these tabs:

1. `MAIN DASHBOARD`
2. `HTML TEMPLATES`
3. `RECIPIENTS`
4. `LOGS`

These names must match the Apps Script code.

---

### Step 2: Set up the `MAIN DASHBOARD` tab

This sheet is the control center.

It stores:
- campaign settings
- template settings
- sending settings
- attachment settings
- summary metrics

Populate it with a `Setting / Value / Notes` structure.

Example settings include:

- `campaign_name`
- `org_name`
- `active_template_name`
- `subject_template`
- `greeting_mode`
- `custom_greeting_template`
- `greeting_fallback`
- `sender_name`
- `reply_to_email`
- `dummy_send_email`
- `redirect_test_drafts`
- `test_send_subject_prefix`
- `test_draft_subject_prefix`
- `test_send_count`
- `send_limit`
- `pause_between_sends_ms`
- `draft_instead_of_send_full`
- `continue_on_error`
- `skip_already_sent`
- `skip_invalid_rows`
- `block_unresolved_placeholders`
- `dedupe_by_email`
- `require_email`
- `require_template`
- `require_subject`
- `require_sender_name`
- `attachment_1_name` through `attachment_10_name`
- `attachment_1_file_id` through `attachment_10_file_id`
- `attachment_1_enabled` through `attachment_10_enabled`
- `last_loaded_template`
- `last_test_draft_run`
- `last_test_send_run`
- `last_full_run`
- `last_error_summary`

This sheet should also contain a metrics section that displays:

- total recipients
- eligible recipients
- invalid recipients
- test drafted recipients
- full drafted recipients
- sent recipients
- failed recipients
- skipped recipients
- duplicate email count
- active template loaded
- enabled attachment count

---

### Step 3: Set up the `HTML TEMPLATES` tab

This is the reusable template library.

Each row is one template.

Recommended columns:

- `template_name`
- `is_active`
- `template_category`
- `subject_template`
- `greeting_template`
- `html_body`
- `text_body`
- `signature_override`
- `notes`
- `version`
- `created_at`
- `updated_at`

At minimum, each usable template should include:

- a unique template name
- a subject template
- a greeting template
- an HTML body

Example:

- `template_name`: `Professor Outreach`
- `subject_template`: `Penn Masala visiting {{city}}`
- `greeting_template`: `Hello {{first_name}},`
- `html_body`: full HTML email body

---

### Step 4: Set up the `RECIPIENTS` tab

This is where contacts are pasted.

Recommended columns:

- `email`
- `first_name`
- `last_name`
- `full_name`
- `prefix`
- `title`
- `organization`
- `college_name`
- `department`
- `city`
- `state`
- `country`
- `contact_type`
- `source`
- `greeting_override`
- `subject_override`
- `custom_1`
- `custom_2`
- `custom_3`
- `custom_4`
- `custom_5`
- `skip`
- `validation_status`
- `draft_status`
- `send_status`
- `attempt_count`
- `last_attempt_at`
- `sent_at`
- `message_id`
- `error_message`
- `notes`

For a simpler version, the minimum useful recipient fields are:

- `email`
- `first_name`
- `last_name`
- `college_name`
- `city`
- `contact_type`
- `greeting_override`
- `subject_override`
- `skip`

The rest are tracking fields and are generally maintained by the script.

---

### Step 5: Set up the `LOGS` tab

This stores an audit trail of all actions.

Recommended columns:

- `timestamp`
- `run_id`
- `campaign_name`
- `action_type`
- `template_name`
- `recipient_email`
- `recipient_name`
- `result`
- `details`
- `subject_used`
- `greeting_used`
- `attachments_used`
- `message_id`
- `error_message`
- `performed_by`

This tab is essential for debugging and tracking campaign activity.

---

### Step 6: Install the Apps Script

Open:

`Extensions → Apps Script`

Paste the project’s script into the Apps Script editor and save it.

Then refresh the Google Sheet.

A custom menu called `Email Sender` should appear.

---

### Step 7: Authorize the script

Run a safe command first, such as:

- `Refresh Dashboard Stats`

Google will ask for permissions.

Approve them so the script can:
- read and write the spreadsheet
- create Gmail drafts
- send Gmail messages
- read Google Drive attachments

---

## Daily Workflow

This section explains the normal workflow for sending a campaign.

### Step 1: Choose your campaign settings

In `MAIN DASHBOARD`, fill in:

- `campaign_name`
- `org_name`
- `active_template_name`
- `sender_name`
- `reply_to_email`
- `dummy_send_email`
- `test_send_count`

Optional:
- `subject_template` if you want to override the template’s subject
- `greeting_mode` if you want greeting behavior controlled globally
- attachment fields if you want files attached
- send behavior options

---

### Step 2: Confirm the template exists

Go to `HTML TEMPLATES` and confirm the row for the chosen `active_template_name` exists.

It should have:
- a valid template name
- a subject template
- a greeting template
- HTML content

---

### Step 3: Paste or update recipients

Go to `RECIPIENTS`.

Paste one recipient per row.

At minimum, each usable row should include:
- an email
- any personalization fields your template needs

If you want to exclude a row, set:

- `skip = TRUE`

No technical status like `READY` is required in the simplified workflow.

---

### Step 4: Load the active template

Use:

`Email Sender → Load Active Template`

This confirms the active template exists and prepares the dashboard for the run.

---

### Step 5: Refresh dashboard stats

Use:

`Email Sender → Refresh Dashboard Stats`

This updates:
- recipient counts
- eligibility counts
- sent counts
- failed counts
- attachment counts

This is your quick sanity check before sending.

---

### Step 6: Validate recipients

Use:

`Email Sender → Validate Recipients`

This checks:
- missing emails
- invalid email format
- duplicates, if enabled
- skipped rows
- basic eligibility readiness

After validation, review:

- `validation_status`
- `error_message`

in the `RECIPIENTS` tab.

---

### Step 7: Preview a sample recipient

Use:

`Email Sender → Preview Top Recipient`

This renders the first eligible row and lets you inspect:
- the recipient email
- the final subject
- the final HTML

This is one of the best safety checks before testing or sending.

---

## Testing Workflow

There are two test modes.

---

### TEST DRAFT

Use:

`Email Sender → TEST DRAFT`

This action:

- takes the top `N` eligible recipients, where `N = test_send_count`
- renders personalized emails for each
- creates Gmail drafts
- does not send anything

If `redirect_test_drafts = TRUE`, drafts are created to the dummy email.

If `redirect_test_drafts = FALSE`, drafts are created to the actual recipient addresses.

#### Use TEST DRAFT when:
- you want to inspect subject lines in Gmail drafts
- you want to verify formatting
- you want to confirm placeholders are being filled correctly
- you want a safe preview without sending anything

#### After TEST DRAFT:
Open Gmail drafts and check:
- greeting
- body formatting
- links
- signature
- attachments
- subject line
- overall visual quality

---

### TEST SEND

Use:

`Email Sender → TEST SEND`

This action:

- takes the top `N` eligible recipients
- renders personalized emails
- sends real emails
- but sends **all of them only to `dummy_send_email`**

This means you get a true inbox test without emailing the actual recipients.

#### Use TEST SEND when:
- you want to see how the email arrives in a real inbox
- you want to verify HTML rendering in an inbox
- you want to check attachment behavior
- you want to see how the email looks on desktop or mobile
- you want to confirm deliverability and reply-to behavior

#### After TEST SEND:
Open the dummy inbox and inspect:
- inbox rendering
- subject line
- spacing
- attachments
- links
- header behavior
- whether the email feels polished

---

## Full Send Workflow

Once both test modes look good, you can send the real campaign.

### Step 1: Final checks
Before sending fully, confirm:

- template is correct
- recipients are validated
- preview looks correct
- test drafts looked correct
- test sends looked correct
- dummy inbox check passed
- attachments are enabled correctly

---

### Step 2: Run full send

Use:

`Email Sender → SEND FULL`

This action:

- processes all eligible recipients up to `send_limit`
- renders personalized emails
- sends to the actual recipient emails
- logs every action
- records status fields

If `draft_instead_of_send_full = TRUE`, this action will create drafts instead of sending.

---

### Step 3: Review results

After the full send, check:

#### In `RECIPIENTS`:
- `send_status`
- `sent_at`
- `error_message`

#### In `LOGS`:
- action history
- failures
- details per row
- template used
- greeting used
- attachments used

---

## Attachments Workflow

Attachments are configured from `MAIN DASHBOARD`.

Each attachment slot includes:

- `attachment_X_name`
- `attachment_X_file_id`
- `attachment_X_enabled`

where `X` is from `1` to `10`.

### To add an attachment:
1. Upload the file to Google Drive.
2. Copy the file ID from the Drive URL.
3. Paste the file ID into the matching `attachment_X_file_id`.
4. Set `attachment_X_enabled = TRUE`.

### Notes:
- only enabled attachments are included
- invalid file IDs can cause failures
- attachments apply globally to the campaign unless future per-template logic is added

---

## Common User Tasks

### Task: Send a new campaign
1. Set campaign details in `MAIN DASHBOARD`
2. Set the active template name
3. Paste recipients
4. Validate recipients
5. Preview top recipient
6. Run TEST DRAFT
7. Run TEST SEND
8. Run SEND FULL

---

### Task: Skip one recipient
Set:
- `skip = TRUE`

for that row in `RECIPIENTS`

---

### Task: Change greeting style
Update:
- `greeting_mode`
- `custom_greeting_template`
- `greeting_fallback`

in `MAIN DASHBOARD`

Or use:
- `greeting_template` in `HTML TEMPLATES`

Or override per row using:
- `greeting_override`

---

### Task: Change sender signature
Either:
- include `{{sender_name}}` and `{{org_name}}` directly in the HTML body
- or use a reusable signature block if your script supports `{{signature}}`

---

### Task: Send as drafts only
Set:
- `draft_instead_of_send_full = TRUE`

Then use:
- `SEND FULL`

This creates drafts instead of sending.

---

# How It Works

## Tab Structure

The system is built around four main tabs.

### `MAIN DASHBOARD`
Stores configuration and summary metrics.

### `HTML TEMPLATES`
Stores reusable subject, greeting, and HTML templates.

### `RECIPIENTS`
Stores recipient rows and per-row personalization data.

### `LOGS`
Stores an audit trail of every major action.

---

## Core Logic

The script works by:

1. reading the current dashboard config
2. loading the selected template
3. reading recipients
4. determining which recipients are eligible
5. rendering personalized subject lines and HTML bodies
6. drafting or sending emails
7. writing tracking data to the recipient rows
8. writing logs to the logs sheet
9. refreshing summary metrics

---

## Eligibility Logic

In the simplified non-technical workflow, a row is eligible if:

- `email` is not blank
- `email` is valid
- `skip` is not true
- `validation_status` is not `INVALID` if invalid rows are skipped
- `send_status` is not `SENT` if already-sent rows are skipped

This means users generally do **not** need to mark rows with a special status like `READY`.

That makes the system much easier to use.

---

## Template Rendering

Templates use placeholder replacement.

Example placeholders:

- `{{first_name}}`
- `{{last_name}}`
- `{{full_name}}`
- `{{college_name}}`
- `{{city}}`
- `{{contact_type}}`
- `{{org_name}}`
- `{{sender_name}}`
- `{{campaign_name}}`
- `{{reply_to_email}}`
- `{{greeting}}`
- `{{signature}}`
- `{{custom_1}}` through `{{custom_5}}`

The rendering engine replaces placeholders with values from:
- the recipient row
- dashboard settings
- computed values like greeting and date

If unresolved placeholders remain and blocking is enabled, the row fails instead of being drafted or sent.

---

## Greeting Logic

Greeting behavior can come from several places.

Priority order should generally be:

1. `greeting_override` from the recipient row
2. global greeting mode from `MAIN DASHBOARD`
3. template greeting from `HTML TEMPLATES`
4. fallback greeting

Common greeting modes:

- `TEMPLATE`
- `Hello First`
- `Hi First`
- `Dear First`
- `Dear Last`
- `Dear Prefix Last`
- `Custom`
- `None`

Examples:
- `Hello Rishabh,`
- `Hi Maya,`
- `Dear Professor Patel,`
- `Dear Smith,`

If name data is missing, fallback behavior applies.

---

## Draft and Send Logic

There are three main execution modes.

### TEST DRAFT
Creates drafts only.

### TEST SEND
Sends real emails to the dummy inbox only.

### SEND FULL
Sends to real recipients, or creates full drafts if configured to do so.

Each of these modes:
- renders the email
- collects enabled attachments
- writes logs
- updates row tracking fields

---

## Logging

Every important action writes a row into `LOGS`.

This includes:
- validation
- draft creation
- test send
- full send
- failures
- skipped rows

Logging provides:
- transparency
- auditability
- debugging support
- campaign history

---

## Statuses and Tracking Fields

The system may use these tracking fields in `RECIPIENTS`:

- `validation_status`
- `draft_status`
- `send_status`
- `attempt_count`
- `last_attempt_at`
- `sent_at`
- `message_id`
- `error_message`

Typical values include:

### `validation_status`
- `VALID`
- `INVALID`
- `SKIPPED`

### `draft_status`
- `TEST_DRAFTED`
- `FULL_DRAFTED`
- `FAILED`

### `send_status`
- `TEST_SENT`
- `SENT`
- `FAILED`

These fields are meant to be maintained by the script, not edited manually by normal users.

---

# Sheet Structure

## MAIN DASHBOARD

This tab stores all configuration and summary metrics.

### Common fields
- campaign name
- active template name
- sender name
- reply-to email
- dummy send email
- greeting mode
- greeting fallback
- attachment settings
- send settings
- safety settings

### Common metrics
- total recipients
- eligible recipients
- invalid recipients
- sent recipients
- failed recipients
- attachment count

---

## HTML TEMPLATES

This tab stores template rows.

### Important fields
- template name
- subject template
- greeting template
- HTML body
- text body
- signature override

Each row is a reusable email template.

---

## RECIPIENTS

This tab stores recipient data.

### User-edited fields
- email
- name fields
- org fields
- personalization fields
- skip

### Script-maintained fields
- validation status
- draft status
- send status
- attempts
- timestamps
- message id
- error message

---

## LOGS

This tab stores an event history.

### Common events
- template loaded
- recipients validated
- draft created
- test sent
- email sent
- row failed

This tab should generally never be manually edited during normal use.

---

# Placeholder Reference

Below are common placeholders supported by the system.

## Recipient placeholders
- `{{email}}`
- `{{first_name}}`
- `{{last_name}}`
- `{{full_name}}`
- `{{prefix}}`
- `{{title}}`
- `{{organization}}`
- `{{college_name}}`
- `{{department}}`
- `{{city}}`
- `{{state}}`
- `{{country}}`
- `{{contact_type}}`
- `{{source}}`
- `{{custom_1}}`
- `{{custom_2}}`
- `{{custom_3}}`
- `{{custom_4}}`
- `{{custom_5}}`

## Dashboard placeholders
- `{{campaign_name}}`
- `{{org_name}}`
- `{{sender_name}}`
- `{{reply_to_email}}`

## Computed placeholders
- `{{greeting}}`
- `{{signature}}`
- `{{today_date}}`

---

# Menu Actions

The custom menu should include actions like:

- `Refresh Dashboard Stats`
- `Load Active Template`
- `Validate Recipients`
- `Preview Top Recipient`
- `TEST DRAFT`
- `TEST SEND`
- `SEND FULL`
- `Reset Processing Fields`

### What each does

#### Refresh Dashboard Stats
Recomputes summary numbers in the dashboard.

#### Load Active Template
Confirms the chosen template exists and is available.

#### Validate Recipients
Checks recipient rows for issues and updates validation fields.

#### Preview Top Recipient
Renders a preview for the first eligible row.

#### TEST DRAFT
Creates drafts for the top N eligible rows.

#### TEST SEND
Sends top N rendered emails to the dummy inbox.

#### SEND FULL
Sends or drafts the full campaign.

#### Reset Processing Fields
Clears script-maintained statuses and timestamps.

---

# Recommended Best Practices

## Always test before full send
Never go straight to `SEND FULL`.

Use:
1. preview
2. test draft
3. test send
4. full send

---

## Keep a real dummy inbox
Always maintain a working:
- `dummy_send_email`

This is critical for safe testing.

---

## Do not manually edit script-owned fields
Avoid editing:
- `validation_status`
- `draft_status`
- `send_status`
- `attempt_count`
- `sent_at`
- `message_id`
- `error_message`

Edit only user-facing fields unless intentionally resetting.

---

## Use placeholders consistently
If your template includes:
- `{{city}}`
- `{{college_name}}`

make sure those columns exist and are populated where needed.

---

## Use `skip = TRUE` for exceptions
Instead of deleting rows, use the skip field to exclude recipients.

This keeps your data intact and makes audits easier.

---

## Keep logs
Do not delete `LOGS` during a campaign.

Logs are the best debugging tool in the system.

---

# Troubleshooting

## “No eligible recipients”
This usually means one or more of the following:
- no email is present
- email format is invalid
- row is skipped
- row is already sent
- row is marked invalid and invalid rows are being skipped

Check:
- `email`
- `skip`
- `validation_status`
- `send_status`

---

## “Template not found”
Make sure:
- `active_template_name` in `MAIN DASHBOARD`
exactly matches
- `template_name` in `HTML TEMPLATES`

Spacing and capitalization matter.

---

## Greeting looks wrong
Check:
- `greeting_override`
- `greeting_mode`
- `custom_greeting_template`
- `greeting_template`
- fallback settings

---

## Missing placeholders in the final email
This usually means:
- the template uses a placeholder that does not exist in the recipient or dashboard data
- the field is blank
- placeholder names do not exactly match the column/config keys

---

## Attachments not working
Check:
- attachment is enabled
- file ID is correct
- file exists in Google Drive
- the script has Drive permissions

---

## Duplicate recipients flagged
If `dedupe_by_email = TRUE`, duplicate emails will be marked invalid during validation.

Turn it off only if you intentionally want duplicates.

---

## Test draft works but test send fails
Check:
- `dummy_send_email`
- Gmail sending permissions
- attachment issues
- unresolved placeholder blocking
- quota-related Gmail issues

---

# Future Improvements

Potential future features include:

- sidebar UI in Sheets
- true HTML preview dialog
- template picker dropdown
- per-template attachment sets
- per-recipient attachment controls
- better placeholder audit tools
- retry failed rows button
- archive completed campaigns
- analytics dashboard
- environment separation for dev/staging/prod
- sheet protection for script-owned columns
- richer signature builder
- CC/BCC support
- unsubscribe footer support
- scheduled sends

---

# Summary

This framework is meant to make campaign emailing inside Google Sheets:

- safe
- reusable
- personalized
- understandable
- manageable by non-technical users

The intended user workflow is:

1. choose a template
2. paste recipients
3. validate
4. preview
5. test draft
6. test send
7. send full
8. inspect logs

That is the full operating model of the system.
