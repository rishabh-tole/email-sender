/***********************
 * SHEETS EMAIL SENDER *
 ***********************/

const SHEET_DASHBOARD = 'MAIN DASHBOARD';
const SHEET_TEMPLATES = 'HTML TEMPLATES';
const SHEET_RECIPIENTS = 'RECIPIENTS';
const SHEET_LOGS = 'LOGS';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Sender')
    .addItem('Refresh Dashboard Stats', 'refreshDashboardStats')
    .addItem('Load Active Template', 'loadActiveTemplate')
    .addItem('Validate Recipients', 'validateRecipients')
    .addItem('Preview Top Recipient', 'previewTopRecipient')
    .addSeparator()
    .addItem('TEST DRAFT', 'testDraft')
    .addItem('TEST SEND', 'testSend')
    .addItem('SEND FULL', 'sendFull')
    .addSeparator()
    .addItem('Reset Processing Fields', 'resetProcessingFields')
    .addToUi();
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetOrThrow_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Missing sheet: ${name}`);
  }
  return sheet;
}

function getDashboardConfig_() {
  const sheet = getSheetOrThrow_(SHEET_DASHBOARD);
  const values = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const val = values[i][1];
    if (key) config[key] = val;
  }
  return config;
}

function setDashboardValue_(key, value) {
  const sheet = getSheetOrThrow_(SHEET_DASHBOARD);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

function getTemplates_() {
  const sheet = getSheetOrThrow_(SHEET_TEMPLATES);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((h, idx) => row[h] = values[i][idx]);
    row._rowNumber = i + 1;
    rows.push(row);
  }
  return rows;
}

function getTemplateByName_(templateName) {
  const templates = getTemplates_();
  const t = templates.find(row =>
    String(row.template_name || '').trim() === String(templateName || '').trim()
  );
  if (!t) return null;
  return t;
}

function getRecipientsData_() {
  const sheet = getSheetOrThrow_(SHEET_RECIPIENTS);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1).map((arr, idx) => {
    const row = {};
    headers.forEach((h, j) => row[h] = arr[j]);
    row._rowNumber = idx + 2;
    return row;
  });

  return { headers, rows };
}

function updateRecipientFields_(rowNumber, updates) {
  const sheet = getSheetOrThrow_(SHEET_RECIPIENTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  Object.keys(updates).forEach(key => {
    const col = headers.indexOf(key);
    if (col !== -1) {
      sheet.getRange(rowNumber, col + 1).setValue(updates[key]);
    }
  });
}

function appendLog_(entry) {
  const sheet = getSheetOrThrow_(SHEET_LOGS);
  const email = Session.getActiveUser().getEmail() || 'unknown';

  sheet.appendRow([
    new Date(),
    entry.run_id || '',
    entry.campaign_name || '',
    entry.action_type || '',
    entry.template_name || '',
    entry.recipient_id || '',
    entry.recipient_email || '',
    entry.recipient_name || '',
    entry.result || '',
    entry.details || '',
    entry.subject_used || '',
    entry.greeting_used || '',
    entry.attachments_used || '',
    entry.message_id || '',
    entry.error_message || '',
    email,
  ]);
}

function bool_(v) {
  if (typeof v === 'boolean') return v;
  return String(v).trim().toUpperCase() === 'TRUE';
}

function safeString_(v) {
  return v == null ? '' : String(v);
}

function todayString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function renderTemplateString_(template, data) {
  let out = safeString_(template);
  out = out.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => {
    const val = data[key];
    return val == null ? '' : String(val);
  });
  return out;
}

function containsUnresolvedPlaceholders_(str) {
  return /\{\{\s*[a-zA-Z0-9_]+\s*\}\}/.test(safeString_(str));
}

function getSignatureHtml_(config, template) {
  if (template && template.signature_override) {
    return String(template.signature_override);
  }
  return '';
}

function buildGreeting_(recipient, config, template) {
  if (recipient.greeting_override) return String(recipient.greeting_override);

  const mode = safeString_(config.greeting_mode).trim();
  const firstName = safeString_(recipient.first_name).trim();
  const lastName = safeString_(recipient.last_name).trim();
  const prefix = safeString_(recipient.prefix).trim();
  const fallback = safeString_(config.greeting_fallback).trim() || 'Hello';

  if (mode === 'TEMPLATE') {
    const g = safeString_(template && template.greeting_template);
    if (g) {
      const rendered = renderTemplateString_(g, {
        ...recipient,
        sender_name: config.sender_name,
        org_name: config.org_name,
        campaign_name: config.campaign_name,
        reply_to_email: config.reply_to_email,
        today_date: todayString_(),
      });
      return rendered || `${fallback},`;
    }
    return `${fallback},`;
  }

  if (mode === 'Hello First') return firstName ? `Hello ${firstName},` : `${fallback},`;
  if (mode === 'Hi First') return firstName ? `Hi ${firstName},` : `${fallback},`;
  if (mode === 'Dear First') return firstName ? `Dear ${firstName},` : `Dear,`;
  if (mode === 'Dear Last') return lastName ? `Dear ${lastName},` : `Dear,`;
  if (mode === 'Dear Prefix Last') {
    if (prefix && lastName) return `Dear ${prefix} ${lastName},`;
    if (lastName) return `Dear ${lastName},`;
    return `Dear,`;
  }
  if (mode === 'Custom') {
    const custom = safeString_(config.custom_greeting_template);
    if (!custom) return `${fallback},`;
    return renderTemplateString_(custom, recipient);
  }
  if (mode === 'None') return '';

  return `${fallback},`;
}

function collectEnabledAttachments_(config) {
  const attachments = [];
  const names = [];

  for (let i = 1; i <= 10; i++) {
    const enabled = bool_(config[`attachment_${i}_enabled`]);
    const fileId = safeString_(config[`attachment_${i}_file_id`]).trim();
    const label = safeString_(config[`attachment_${i}_name`]).trim();

    if (enabled && fileId) {
      const file = DriveApp.getFileById(fileId);
      attachments.push(file.getBlob());
      names.push(label || file.getName());
    }
  }

  return { attachments, names };
}

function getRecipientDataContext_(recipient, config, template, greeting) {
  return {
    ...recipient,
    sender_name: safeString_(config.sender_name),
    org_name: safeString_(config.org_name),
    campaign_name: safeString_(config.campaign_name),
    reply_to_email: safeString_(config.reply_to_email),
    greeting,
    signature: getSignatureHtml_(config, template),
    today_date: todayString_(),
  };
}

function buildRenderedEmail_(recipient, config, template) {
  const greeting = buildGreeting_(recipient, config, template);
  const context = getRecipientDataContext_(recipient, config, template, greeting);

  const subjectTemplate =
    safeString_(recipient.subject_override).trim() ||
    safeString_(config.subject_template).trim() ||
    safeString_(template.subject_template).trim();

  const htmlBodyTemplate = safeString_(template.html_body);
  const textBodyTemplate = safeString_(template.text_body);

  const renderedSubject = renderTemplateString_(subjectTemplate, context);
  const renderedHtml = renderTemplateString_(htmlBodyTemplate, context);
  const renderedText = renderTemplateString_(textBodyTemplate, context);

  return {
    greeting,
    subject: renderedSubject,
    htmlBody: renderedHtml,
    textBody: renderedText,
    context,
  };
}

function isEmailValid_(email) {
  const e = safeString_(email).trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
}

function getEligibleRecipients_(config) {
  const { rows } = getRecipientsData_();
  const skipAlreadySent = bool_(config.skip_already_sent);

  return rows.filter(r => {
    const email = safeString_(r.email).trim();
    const sendStatus = safeString_(r.send_status).trim().toUpperCase();
    const validationStatus = safeString_(r.validation_status).trim().toUpperCase();
    const skip = safeString_(r.skip).trim().toUpperCase();

    if (!email) return false;
    if (!isEmailValid_(email)) return false;
    if (skip === 'TRUE' || skip === 'YES' || skip === '1') return false;
    if (skipAlreadySent && sendStatus === 'SENT') return false;
    if (bool_(config.skip_invalid_rows) && validationStatus === 'INVALID') return false;

    return true;
  });
}

function loadActiveTemplate() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();

    if (!templateName) throw new Error('MAIN DASHBOARD: active_template_name is blank.');

    const template = getTemplateByName_(templateName);
    if (!template) throw new Error(`Template not found: ${templateName}`);

    if (!safeString_(config.subject_template).trim() && safeString_(template.subject_template).trim()) {
      setDashboardValue_('subject_template', template.subject_template);
    }

    setDashboardValue_('last_loaded_template', `${templateName} @ ${new Date()}`);
    refreshDashboardStats();
    ui.alert(`Loaded template: ${templateName}`);
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`Load Active Template failed:\n${err.message}`);
  }
}

function validateRecipients() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();
    const requireTemplate = bool_(config.require_template);
    const requireEmail = bool_(config.require_email);
    const dedupeByEmail = bool_(config.dedupe_by_email);

    let template = null;
    if (templateName) {
      template = getTemplateByName_(templateName);
    }
    if (requireTemplate && !template) {
      throw new Error('Active template missing or not found.');
    }

    const { rows } = getRecipientsData_();
    const seen = new Set();
    let duplicateCount = 0;
    let invalidCount = 0;
    let validCount = 0;

    rows.forEach(r => {
      const email = safeString_(r.email).trim();
      const skip = safeString_(r.skip).trim().toUpperCase();

      if (!email && !skip) {
        updateRecipientFields_(r._rowNumber, {
          validation_status: '',
          error_message: '',
        });
        return;
      }

      if (skip === 'TRUE' || skip === 'YES' || skip === '1') {
        updateRecipientFields_(r._rowNumber, {
          validation_status: 'SKIPPED',
          error_message: '',
        });

        appendLog_({
          campaign_name: config.campaign_name,
          action_type: 'VALIDATE_ROW',
          template_name: templateName,
          recipient_email: r.email,
          recipient_name: r.full_name || `${safeString_(r.first_name)} ${safeString_(r.last_name)}`.trim(),
          result: 'SKIPPED',
          details: 'Row marked skip',
        });
        return;
      }

      let validation = 'VALID';
      let error = '';

      if (requireEmail && !email) {
        validation = 'INVALID';
        error = 'Missing email';
      } else if (email && !isEmailValid_(email)) {
        validation = 'INVALID';
        error = 'Invalid email format';
      } else if (dedupeByEmail && email) {
        const lower = email.toLowerCase();
        if (seen.has(lower)) {
          validation = 'INVALID';
          error = 'Duplicate email';
          duplicateCount++;
        } else {
          seen.add(lower);
        }
      }

      if (validation === 'INVALID') invalidCount++;
      if (validation === 'VALID') validCount++;

      updateRecipientFields_(r._rowNumber, {
        validation_status: validation,
        error_message: error,
      });

      appendLog_({
        campaign_name: config.campaign_name,
        action_type: 'VALIDATE_ROW',
        template_name: templateName,
        recipient_email: r.email,
        recipient_name: r.full_name || `${safeString_(r.first_name)} ${safeString_(r.last_name)}`.trim(),
        result: validation,
        details: error,
      });
    });

    setDashboardValue_('last_error_summary', '');
    refreshDashboardStats();
    ui.alert(`Validation complete.\nValid rows: ${validCount}\nInvalid rows: ${invalidCount}\nDuplicate emails: ${duplicateCount}`);
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`Validate Recipients failed:\n${err.message}`);
  }
}

function previewTopRecipient() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();
    const template = getTemplateByName_(templateName);

    if (!template) throw new Error(`Template not found: ${templateName}`);

    const eligible = getEligibleRecipients_(config).filter(r => {
      if (bool_(config.skip_invalid_rows)) {
        return safeString_(r.validation_status).trim().toUpperCase() !== 'INVALID';
      }
      return true;
    });

    if (!eligible.length) throw new Error('No eligible recipients found.');

    const recipient = eligible[0];
    const rendered = buildRenderedEmail_(recipient, config, template);

    ui.alert(
      'Preview Top Recipient',
      `TO: ${recipient.email}\n\nSUBJECT:\n${rendered.subject}\n\nHTML:\n${rendered.htmlBody}`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`Preview failed:\n${err.message}`);
  }
}

function testDraft() {
  const ui = SpreadsheetApp.getUi();
  const runId = `TEST-DRAFT-${Date.now()}`;

  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();
    const template = getTemplateByName_(templateName);

    if (!template) throw new Error(`Template not found: ${templateName}`);

    const testCount = Number(config.test_send_count || 0);
    if (!testCount || testCount < 1) throw new Error('test_send_count must be >= 1');

    const redirectDrafts = bool_(config.redirect_test_drafts);
    const dummySendEmail = safeString_(config.dummy_send_email).trim();

    if (redirectDrafts && !dummySendEmail) {
      throw new Error('redirect_test_drafts is TRUE but dummy_send_email is blank.');
    }

    const { attachments, names } = collectEnabledAttachments_(config);

    const eligible = getEligibleRecipients_(config).filter(r => {
      if (bool_(config.skip_invalid_rows)) {
        return safeString_(r.validation_status).trim().toUpperCase() !== 'INVALID';
      }
      return true;
    }).slice(0, testCount);

    if (!eligible.length) throw new Error('No eligible recipients for TEST DRAFT.');

    let drafted = 0;

    eligible.forEach(recipient => {
      try {
        const rendered = buildRenderedEmail_(recipient, config, template);

        if (bool_(config.require_subject) && !safeString_(rendered.subject).trim()) {
          throw new Error('Rendered subject is blank');
        }

        if (
          bool_(config.block_unresolved_placeholders) &&
          (containsUnresolvedPlaceholders_(rendered.subject) || containsUnresolvedPlaceholders_(rendered.htmlBody))
        ) {
          throw new Error('Unresolved placeholders remain in rendered output');
        }

        const to = redirectDrafts ? dummySendEmail : safeString_(recipient.email).trim();
        const subject = `${safeString_(config.test_draft_subject_prefix).trim()} ${rendered.subject}`.trim();

        const htmlBody = redirectDrafts
          ? `<p><strong>INTENDED RECIPIENT:</strong> ${recipient.email}</p>${rendered.htmlBody}`
          : rendered.htmlBody;

        const draft = GmailApp.createDraft(
          to,
          subject,
          rendered.textBody || 'HTML email draft',
          {
            htmlBody,
            attachments,
            replyTo: safeString_(config.reply_to_email).trim() || undefined,
            name: safeString_(config.sender_name).trim() || undefined,
          }
        );

        updateRecipientFields_(recipient._rowNumber, {
          draft_status: 'TEST_DRAFTED',
          attempt_count: Number(recipient.attempt_count || 0) + 1,
          last_attempt_at: new Date(),
          error_message: '',
        });

        appendLog_({
          run_id: runId,
          campaign_name: config.campaign_name,
          action_type: 'TEST_DRAFT_CREATE',
          template_name: templateName,
          recipient_id: recipient.recipient_id,
          recipient_email: recipient.email,
          recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
          result: 'SUCCESS',
          details: `Draft created to ${to}`,
          subject_used: subject,
          greeting_used: rendered.greeting,
          attachments_used: names.join('; '),
          message_id: draft.getMessage().getId(),
        });

        drafted++;
      } catch (rowErr) {
        updateRecipientFields_(recipient._rowNumber, {
          draft_status: 'FAILED',
          error_message: rowErr.message,
          attempt_count: Number(recipient.attempt_count || 0) + 1,
          last_attempt_at: new Date(),
        });

        appendLog_({
          run_id: runId,
          campaign_name: config.campaign_name,
          action_type: 'TEST_DRAFT_CREATE',
          template_name: templateName,
          recipient_id: recipient.recipient_id,
          recipient_email: recipient.email,
          recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
          result: 'FAILED',
          error_message: rowErr.message,
        });

        if (!bool_(config.continue_on_error)) throw rowErr;
      }
    });

    setDashboardValue_('last_test_draft_run', new Date());
    setDashboardValue_('last_error_summary', '');
    refreshDashboardStats();
    ui.alert(`TEST DRAFT complete.\nDrafts created: ${drafted}`);
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`TEST DRAFT failed:\n${err.message}`);
  }
}

function testSend() {
  const ui = SpreadsheetApp.getUi();
  const runId = `TEST-SEND-${Date.now()}`;

  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();
    const template = getTemplateByName_(templateName);

    if (!template) throw new Error(`Template not found: ${templateName}`);

    const testCount = Number(config.test_send_count || 0);
    if (!testCount || testCount < 1) throw new Error('test_send_count must be >= 1');

    const dummySendEmail = safeString_(config.dummy_send_email).trim();
    if (!dummySendEmail) throw new Error('dummy_send_email is blank.');

    const { attachments, names } = collectEnabledAttachments_(config);

    const eligible = getEligibleRecipients_(config).filter(r => {
      if (bool_(config.skip_invalid_rows)) {
        return safeString_(r.validation_status).trim().toUpperCase() !== 'INVALID';
      }
      return true;
    }).slice(0, testCount);

    if (!eligible.length) throw new Error('No eligible recipients for TEST SEND.');

    let sent = 0;

    eligible.forEach(recipient => {
      try {
        const rendered = buildRenderedEmail_(recipient, config, template);

        if (bool_(config.require_subject) && !safeString_(rendered.subject).trim()) {
          throw new Error('Rendered subject is blank');
        }

        if (
          bool_(config.block_unresolved_placeholders) &&
          (containsUnresolvedPlaceholders_(rendered.subject) || containsUnresolvedPlaceholders_(rendered.htmlBody))
        ) {
          throw new Error('Unresolved placeholders remain in rendered output');
        }

        const subject = `${safeString_(config.test_send_subject_prefix).trim()} ${rendered.subject}`.trim();
        const htmlBody =
          `<p><strong>INTENDED RECIPIENT:</strong> ${recipient.email}</p>` +
          `<p><strong>RECIPIENT NAME:</strong> ${(recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`).trim()}</p>` +
          rendered.htmlBody;

        GmailApp.sendEmail(
          dummySendEmail,
          subject,
          rendered.textBody || 'HTML email',
          {
            htmlBody,
            attachments,
            replyTo: safeString_(config.reply_to_email).trim() || undefined,
            name: safeString_(config.sender_name).trim() || undefined,
          }
        );

        updateRecipientFields_(recipient._rowNumber, {
          send_status: 'TEST_SENT',
          attempt_count: Number(recipient.attempt_count || 0) + 1,
          last_attempt_at: new Date(),
          error_message: '',
        });

        appendLog_({
          run_id: runId,
          campaign_name: config.campaign_name,
          action_type: 'TEST_SEND',
          template_name: templateName,
          recipient_id: recipient.recipient_id,
          recipient_email: recipient.email,
          recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
          result: 'SUCCESS',
          details: `Test email sent to dummy address ${dummySendEmail}`,
          subject_used: subject,
          greeting_used: rendered.greeting,
          attachments_used: names.join('; '),
        });

        sent++;
      } catch (rowErr) {
        updateRecipientFields_(recipient._rowNumber, {
          send_status: 'FAILED',
          error_message: rowErr.message,
          attempt_count: Number(recipient.attempt_count || 0) + 1,
          last_attempt_at: new Date(),
        });

        appendLog_({
          run_id: runId,
          campaign_name: config.campaign_name,
          action_type: 'TEST_SEND',
          template_name: templateName,
          recipient_id: recipient.recipient_id,
          recipient_email: recipient.email,
          recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
          result: 'FAILED',
          error_message: rowErr.message,
        });

        if (!bool_(config.continue_on_error)) throw rowErr;
      }
    });

    setDashboardValue_('last_test_send_run', new Date());
    setDashboardValue_('last_error_summary', '');
    refreshDashboardStats();
    ui.alert(`TEST SEND complete.\nEmails sent to dummy inbox: ${sent}`);
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`TEST SEND failed:\n${err.message}`);
  }
}

function sendFull() {
  const ui = SpreadsheetApp.getUi();
  const runId = `FULL-${Date.now()}`;

  try {
    const config = getDashboardConfig_();
    const templateName = safeString_(config.active_template_name).trim();
    const template = getTemplateByName_(templateName);

    if (!template) throw new Error(`Template not found: ${templateName}`);

    const limit = Number(config.send_limit || 0) || 1000;
    const pauseMs = Number(config.pause_between_sends_ms || 0) || 0;
    const draftInstead = bool_(config.draft_instead_of_send_full);

    const { attachments, names } = collectEnabledAttachments_(config);

    let eligible = getEligibleRecipients_(config).filter(r => {
      if (bool_(config.skip_invalid_rows)) {
        return safeString_(r.validation_status).trim().toUpperCase() !== 'INVALID';
      }
      return true;
    });

    eligible = eligible.slice(0, limit);

    if (!eligible.length) throw new Error('No eligible recipients for full send.');

    let successCount = 0;
    let failCount = 0;

    eligible.forEach((recipient, idx) => {
      try {
        const rendered = buildRenderedEmail_(recipient, config, template);

        if (bool_(config.require_subject) && !safeString_(rendered.subject).trim()) {
          throw new Error('Rendered subject is blank');
        }

        if (
          bool_(config.block_unresolved_placeholders) &&
          (containsUnresolvedPlaceholders_(rendered.subject) || containsUnresolvedPlaceholders_(rendered.htmlBody))
        ) {
          throw new Error('Unresolved placeholders remain in rendered output');
        }

        const to = safeString_(recipient.email).trim();
        if (!isEmailValid_(to)) throw new Error('Invalid recipient email');

        if (draftInstead) {
          const draft = GmailApp.createDraft(
            to,
            rendered.subject,
            rendered.textBody || 'HTML email draft',
            {
              htmlBody: rendered.htmlBody,
              attachments,
              replyTo: safeString_(config.reply_to_email).trim() || undefined,
              name: safeString_(config.sender_name).trim() || undefined,
            }
          );

          updateRecipientFields_(recipient._rowNumber, {
            draft_status: 'FULL_DRAFTED',
            send_status: '',
            attempt_count: Number(recipient.attempt_count || 0) + 1,
            last_attempt_at: new Date(),
            error_message: '',
            message_id: draft.getMessage().getId(),
          });

          appendLog_({
            run_id: runId,
            campaign_name: config.campaign_name,
            action_type: 'FULL_DRAFT_CREATE',
            template_name: templateName,
            recipient_id: recipient.recipient_id,
            recipient_email: recipient.email,
            recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
            result: 'SUCCESS',
            details: `Full draft created to ${to}`,
            subject_used: rendered.subject,
            greeting_used: rendered.greeting,
            attachments_used: names.join('; '),
            message_id: draft.getMessage().getId(),
          });
        } else {
          GmailApp.sendEmail(
            to,
            rendered.subject,
            rendered.textBody || 'HTML email',
            {
              htmlBody: rendered.htmlBody,
              attachments,
              replyTo: safeString_(config.reply_to_email).trim() || undefined,
              name: safeString_(config.sender_name).trim() || undefined,
            }
          );

          updateRecipientFields_(recipient._rowNumber, {
            send_status: 'SENT',
            draft_status: '',
            attempt_count: Number(recipient.attempt_count || 0) + 1,
            last_attempt_at: new Date(),
            sent_at: new Date(),
            error_message: '',
          });

          appendLog_({
            run_id: runId,
            campaign_name: config.campaign_name,
            action_type: 'SEND_EMAIL',
            template_name: templateName,
            recipient_id: recipient.recipient_id,
            recipient_email: recipient.email,
            recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
            result: 'SUCCESS',
            details: `Email sent to ${to}`,
            subject_used: rendered.subject,
            greeting_used: rendered.greeting,
            attachments_used: names.join('; '),
          });
        }

        successCount++;
        if (pauseMs > 0 && idx < eligible.length - 1) Utilities.sleep(pauseMs);
      } catch (rowErr) {
        failCount++;
        updateRecipientFields_(recipient._rowNumber, {
          send_status: 'FAILED',
          attempt_count: Number(recipient.attempt_count || 0) + 1,
          last_attempt_at: new Date(),
          error_message: rowErr.message,
        });

        appendLog_({
          run_id: runId,
          campaign_name: config.campaign_name,
          action_type: draftInstead ? 'FULL_DRAFT_CREATE' : 'SEND_EMAIL',
          template_name: templateName,
          recipient_id: recipient.recipient_id,
          recipient_email: recipient.email,
          recipient_name: recipient.full_name || `${safeString_(recipient.first_name)} ${safeString_(recipient.last_name)}`.trim(),
          result: 'FAILED',
          error_message: rowErr.message,
        });

        if (!bool_(config.continue_on_error)) throw rowErr;
      }
    });

    setDashboardValue_('last_full_run', new Date());
    setDashboardValue_('last_error_summary', failCount ? `${failCount} row(s) failed` : '');
    refreshDashboardStats();

    ui.alert(
      draftInstead ? 'Full Draft Creation Complete' : 'Full Send Complete',
      `Success: ${successCount}\nFailed: ${failCount}`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`SEND FULL failed:\n${err.message}`);
  }
}

function refreshDashboardStats() {
  const dashboard = getSheetOrThrow_(SHEET_DASHBOARD);
  const { rows } = getRecipientsData_();

  const total = rows.filter(r => safeString_(r.email).trim()).length;
  const ready = getEligibleRecipients_(getDashboardConfig_()).length;
  const invalid = rows.filter(r => safeString_(r.validation_status).trim().toUpperCase() === 'INVALID').length;
  const testDrafted = rows.filter(r => safeString_(r.draft_status).trim().toUpperCase() === 'TEST_DRAFTED').length;
  const fullDrafted = rows.filter(r => safeString_(r.draft_status).trim().toUpperCase() === 'FULL_DRAFTED').length;
  const sent = rows.filter(r => safeString_(r.send_status).trim().toUpperCase() === 'SENT').length;
  const failed = rows.filter(r =>
    safeString_(r.send_status).trim().toUpperCase() === 'FAILED' ||
    safeString_(r.draft_status).trim().toUpperCase() === 'FAILED'
  ).length;
  const skipped = rows.filter(r => {
    const skip = safeString_(r.skip).trim().toUpperCase();
    return skip === 'TRUE' || skip === 'YES' || skip === '1';
  }).length;

  const emailMap = {};
  let duplicates = 0;
  rows.forEach(r => {
    const e = safeString_(r.email).trim().toLowerCase();
    if (!e) return;
    emailMap[e] = (emailMap[e] || 0) + 1;
  });
  Object.keys(emailMap).forEach(e => {
    if (emailMap[e] > 1) duplicates += emailMap[e] - 1;
  });

  const config = getDashboardConfig_();
  const templateName = safeString_(config.active_template_name).trim();
  const activeTemplateLoaded = getTemplateByName_(templateName) ? 'TRUE' : 'FALSE';

  let enabledAttachmentCount = 0;
  for (let i = 1; i <= 10; i++) {
    if (bool_(config[`attachment_${i}_enabled`]) && safeString_(config[`attachment_${i}_file_id`]).trim()) {
      enabledAttachmentCount++;
    }
  }

  const metricMap = {
    total_recipients: total,
    ready_recipients: ready,
    invalid_recipients: invalid,
    test_drafted_recipients: testDrafted,
    full_drafted_recipients: fullDrafted,
    sent_recipients: sent,
    failed_recipients: failed,
    skipped_recipients: skipped,
    duplicate_email_count: duplicates,
    active_template_loaded: activeTemplateLoaded,
    enabled_attachment_count: enabledAttachmentCount,
  };

  const values = dashboard.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][4] || '').trim();
    if (key && key in metricMap) {
      dashboard.getRange(i + 1, 6).setValue(metricMap[key]);
    }
  }
}

function resetProcessingFields() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = getSheetOrThrow_(SHEET_RECIPIENTS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      ui.alert('No recipient data found.');
      return;
    }

    const headers = data[0].map(h => String(h).trim());
    const resetCols = [
      'validation_status',
      'draft_status',
      'send_status',
      'attempt_count',
      'last_attempt_at',
      'sent_at',
      'message_id',
      'error_message',
    ];

    for (let r = 2; r <= data.length; r++) {
      resetCols.forEach(colName => {
        const colIndex = headers.indexOf(colName);
        if (colIndex !== -1) {
          sheet.getRange(r, colIndex + 1).clearContent();
        }
      });
    }

    refreshDashboardStats();
    ui.alert('Processing fields reset.');
  } catch (err) {
    setDashboardValue_('last_error_summary', err.message);
    ui.alert(`Reset failed:\n${err.message}`);
  }
}
