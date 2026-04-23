/**
 * Free Mail Merge
 *
 * A free Google Apps Script that turns a Sheet + a Gmail draft into a mail merge tool.
 * Drop-in replacement for YAMM. Paste this file in Extensions → Apps Script.
 *
 * Setup:
 *   1. Edit the 3 constants below (FROM_NAME, REPLY_TO, TEST_EMAIL_DEFAULT).
 *   2. Save (Cmd+S). Reload the Sheet.
 *   3. Menu "Free Mail Merge" appears.
 *
 * Usage:
 *   - Create a draft in Gmail with placeholders {{first_name}}, {{company}}, etc.
 *   - Menu → Pick template draft → select your draft.
 *   - Menu → Send test → receive a real email to yourself.
 *   - Menu → Send batch → fire the campaign.
 *
 * Repo: https://github.com/diegotorreslopez81/free-mail-merge
 */

// ----- EDIT THESE 3 CONSTANTS BEFORE USING -----
const FROM_NAME = "Your Name · Your Company";       // display name in recipient's inbox
const REPLY_TO = "you@yourdomain.com";              // Reply-To + From alias (must be a Send-As alias if different from login)
const TEST_EMAIL_DEFAULT = "you@yourdomain.com";    // default destination for "Send test"
// -----------------------------------------------
// NOTE: the leads tab is chosen from the menu ("🗂  Pick leads sheet"), not hardcoded.
//       If nothing is picked, the script falls back to SHEET_NAME_FALLBACK below.
const SHEET_NAME_FALLBACK = "Leads";

const BATCH_SIZE_DEFAULT = 30;          // default batch size (Apps Script has a 6-min timeout)
const DELAY_MIN_MS = 5000;              // 5 s
const DELAY_MAX_MS = 15000;             // 15 s (keep short to stay under the script limit)
const PROP_DRAFT_ID = "TEMPLATE_DRAFT_ID";
const PROP_SHEET_ID = "LEADS_SHEET_ID";
const PROP_SCHEDULE_LIMIT = "SCHEDULE_LIMIT";
const PROP_SCHEDULE_DAILY_LIMIT = "SCHEDULE_DAILY_LIMIT";

// Column indexes (1-based). Adjust to match your sheet layout.
const COL = {
  email: 1, first_name: 2, last_name: 3, company: 4, sector_huma: 5,
  title: 6, segment: 7, city: 8, num_employees: 9, linkedin_url: 10,
  website: 11, lk_contacted: 12, sent_at: 13, sent_status: 14, error: 15,
  replied_at: 16
};

// ---------- Menu ----------


function onOpen() {
  SpreadsheetApp.getUi().createMenu("✉️ Free Mail Merge")
    .addItem("🗂  Pick leads sheet", "chooseSheet")
    .addItem("🎯 Pick template draft", "chooseTemplate")
    .addItem("ℹ️  Show current config", "showTemplate")
    .addSeparator()
    .addItem("🔍 Dry run (preview 3 in logs)", "dryRun")
    .addItem("📨 Send test email", "sendTest")
    .addSeparator()
    .addItem("📧 Send batch (skip LK-contacted)", "sendBatchSkipLK")
    .addItem("📧 Send batch (include LK-contacted)", "sendBatchAll")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("⏰ Schedule")
      .addItem("📅 Schedule one-time batch", "scheduleOneTime")
      .addItem("🔁 Schedule daily batch", "scheduleDaily")
      .addSeparator()
      .addItem("📋 List scheduled jobs", "listSchedules")
      .addItem("🗑️  Cancel all scheduled jobs", "cancelSchedules"))
    .addSeparator()
    .addItem("💬 Check replies", "checkReplies")
    .addSeparator()
    .addItem("↩️  Reset all sends", "resetSent")
    .addItem("ℹ️  Campaign status", "showStatus")
    .addToUi();
}

// ---------- Leads sheet selection ----------

function chooseSheet() {
  const sheets = SpreadsheetApp.getActive().getSheets();
  const items = sheets.map(function(s) {
    return { id: s.getSheetId(), label: s.getName() + "  (" + s.getLastRow() + " rows)" };
  });

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:16px;">' +
    '<h3 style="margin:0 0 12px 0;">Pick leads sheet</h3>' +
    '<p style="color:#555;margin:0 0 16px 0;font-size:13px;">Pick the tab with your leads. The script stores the tab ID (not the name), so you can rename it freely afterwards.</p>' +
    '<select id="sel" style="width:100%;padding:8px;font-size:14px;margin-bottom:16px;">' +
      items.map(function(it){ return '<option value="' + it.id + '">' + it.label.replace(/"/g, "&quot;").replace(/</g,"&lt;") + '</option>'; }).join("") +
    '</select>' +
    '<div style="text-align:right;">' +
      '<button onclick="google.script.host.close()" style="padding:8px 14px;margin-right:8px;">Cancel</button>' +
      '<button onclick="save()" style="padding:8px 14px;background:#DA291C;color:white;border:none;border-radius:4px;cursor:pointer;">Save</button>' +
    '</div>' +
    '<script>' +
    'function save() {' +
    '  const id = document.getElementById("sel").value;' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).saveLeadsSheet(id);' +
    '}' +
    '</script>' +
    '</div>'
  ).setWidth(480).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, "Leads sheet");
}

function saveLeadsSheet(sheetId) {
  PropertiesService.getDocumentProperties().setProperty(PROP_SHEET_ID, String(sheetId));
  return true;
}

// ---------- Template draft selection ----------

function chooseTemplate() {
  const drafts = GmailApp.getDrafts();
  if (!drafts.length) {
    SpreadsheetApp.getUi().alert("No drafts in Gmail. Create one first.");
    return;
  }
  const items = drafts.slice(0, 30).map(function(d, i) {
    const m = d.getMessage();
    const subj = (m.getSubject() || "(no subject)").substring(0, 80);
    const date = Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    return {id: d.getId(), label: (i + 1) + ". [" + date + "] " + subj};
  });

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:16px;">' +
    '<h3 style="margin:0 0 12px 0;">Pick template draft</h3>' +
    '<p style="color:#555;margin:0 0 16px 0;font-size:13px;">Pick the Gmail draft that will be used as the template. Placeholders like {{first_name}}, {{company}}, etc. will be replaced per row.</p>' +
    '<select id="sel" style="width:100%;padding:8px;font-size:14px;margin-bottom:16px;">' +
      items.map(function(it){ return '<option value="' + it.id + '">' + it.label.replace(/"/g, "&quot;").replace(/</g,"&lt;") + '</option>'; }).join("") +
    '</select>' +
    '<div style="text-align:right;">' +
      '<button onclick="google.script.host.close()" style="padding:8px 14px;margin-right:8px;">Cancel</button>' +
      '<button onclick="save()" style="padding:8px 14px;background:#DA291C;color:white;border:none;border-radius:4px;cursor:pointer;">Save</button>' +
    '</div>' +
    '<script>' +
    'function save() {' +
    '  const id = document.getElementById("sel").value;' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).saveTemplateDraft(id);' +
    '}' +
    '</script>' +
    '</div>'
  ).setWidth(520).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, "Template");
}

function saveTemplateDraft(draftId) {
  PropertiesService.getDocumentProperties().setProperty(PROP_DRAFT_ID, draftId);
  return true;
}

function showTemplate() {
  const lines = [];

  // Leads sheet
  const sheet = _getLeadsSheet();
  lines.push("Leads sheet: " + (sheet ? sheet.getName() + "  (" + (sheet.getLastRow() - 1) + " rows)" : "NOT PICKED — use '🗂 Pick leads sheet'"));

  // Template draft
  const id = PropertiesService.getDocumentProperties().getProperty(PROP_DRAFT_ID);
  if (!id) {
    lines.push("Template: NOT PICKED — use '🎯 Pick template draft'");
  } else {
    try {
      const m = GmailApp.getDraft(id).getMessage();
      lines.push("Template: " + m.getSubject() + "  (" + m.getAttachments().length + " attachments)");
    } catch (e) {
      lines.push("Template: draft no longer exists, pick another");
    }
  }

  SpreadsheetApp.getUi().alert("Current config:\n\n" + lines.join("\n"));
}

function sendTest() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Send test", "To which email? (default " + TEST_EMAIL_DEFAULT + ")", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const testEmail = (r.getResponseText() || "").trim() || TEST_EMAIL_DEFAULT;

  const tmpl = _getTemplateDraft();
  if (!tmpl.ok) { ui.alert(tmpl.error); return; }

  const sheet = _getLeadsSheet();
  if (!sheet) { ui.alert("No leads sheet picked. Use '🗂 Pick leads sheet' first."); return; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
  let sample = null;
  for (const row of data) {
    if (row[COL.email - 1] && !row[COL.sent_at - 1]) {
      sample = {
        first_name: row[COL.first_name - 1] || "Test",
        company: row[COL.company - 1] || "Empresa Test",
        sector_huma: row[COL.sector_huma - 1] || "pimes com la teva",
      };
      break;
    }
  }
  if (!sample) sample = { first_name: "Test", company: "Empresa Test", sector_huma: "pimes com la teva" };

  const merged = _mergeTemplate(tmpl.data, sample);
  try {
    GmailApp.sendEmail(testEmail, "[TEST] " + merged.subject, merged.plainBody, {
      htmlBody: merged.htmlBody,
      attachments: tmpl.data.attachments,
      name: FROM_NAME,
      from: REPLY_TO,
      replyTo: REPLY_TO,
    });
    ui.alert("✅ Test sent to " + testEmail + "\n\nUsing lead data: " + sample.first_name + " · " + sample.company + "\n\nVerify:\n- From: " + REPLY_TO + "\n- HTML signature renders\n- Attachment present\n- Placeholders replaced\n- Links work");
  } catch (e) {
    ui.alert("❌ Error: " + e);
  }
}

// ---------- Entry points ----------

function dryRun() {
  _run({ dryRun: true, limit: 3, skipLK: true });
}

function sendBatchSkipLK() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Send batch", "How many emails? (recommended max 30-40 per run)", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const limit = parseInt(r.getResponseText(), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, skipLK: true });
}

function sendBatchAll() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Send batch (include LK)", "How many emails?", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const limit = parseInt(r.getResponseText(), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, skipLK: false });
}

// ---------- Core ----------

function _run(opts) {
  const sheet = _getLeadsSheet();
  if (!sheet) {
    const msg = "No leads sheet picked. Use 'Free Mail Merge → 🗂 Pick leads sheet' first.";
    if (opts.silent) { Logger.log(msg); } else { SpreadsheetApp.getUi().alert(msg); }
    return;
  }

  const tmpl = _getTemplateDraft();
  if (!tmpl.ok) { SpreadsheetApp.getUi().alert(tmpl.error); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No leads in the sheet."); return; }

  const range = sheet.getRange(2, 1, lastRow - 1, 15);
  const data = range.getValues();

  let sent = 0;
  const errors = [];

  for (let i = 0; i < data.length && sent < opts.limit; i++) {
    const row = data[i];
    const sheetRow = i + 2;
    const email = String(row[COL.email - 1] || "").trim();
    const firstName = String(row[COL.first_name - 1] || "").trim();
    const company = String(row[COL.company - 1] || "").trim();
    const sectorHuma = String(row[COL.sector_huma - 1] || "").trim();
    const lkContacted = String(row[COL.lk_contacted - 1] || "").trim().toUpperCase();
    const alreadySent = String(row[COL.sent_at - 1] || "").trim();

    if (!email || !email.includes("@")) continue;
    if (alreadySent) continue;            // skip already sent
    if (opts.skipLK && lkContacted === "YES") continue;

    const merged = _mergeTemplate(tmpl.data, { first_name: firstName, company, sector_huma: sectorHuma });

    if (opts.dryRun) {
      Logger.log("=== DRY RUN [" + (sent + 1) + "] ===");
      Logger.log("To: " + email);
      Logger.log("Subject: " + merged.subject);
      Logger.log(merged.plainBody.substring(0, 600) + "...\n");
    } else {
      try {
        GmailApp.sendEmail(email, merged.subject, merged.plainBody, {
          htmlBody: merged.htmlBody,
          attachments: tmpl.data.attachments,
          name: FROM_NAME,
          from: REPLY_TO,          // force sender to the Send-as alias
          replyTo: REPLY_TO,
        });
        const ts = new Date().toISOString();
        sheet.getRange(sheetRow, COL.sent_at).setValue(ts);
        sheet.getRange(sheetRow, COL.sent_status).setValue("sent");
        sheet.getRange(sheetRow, COL.error).setValue("");
      } catch (e) {
        sheet.getRange(sheetRow, COL.sent_at).setValue(new Date().toISOString());
        sheet.getRange(sheetRow, COL.sent_status).setValue("error");
        sheet.getRange(sheetRow, COL.error).setValue(String(e).substring(0, 500));
        errors.push(email + ": " + e);
      }
      // random delay to be gentle
      const delay = DELAY_MIN_MS + Math.floor(Math.random() * (DELAY_MAX_MS - DELAY_MIN_MS));
      Utilities.sleep(delay);
    }
    sent++;
  }

  const msg = opts.dryRun
    ? "Dry run OK. " + sent + " emails rendered. Check Executions → Logs."
    : "Sent: " + sent + (errors.length ? "\nErrors: " + errors.length : "");
  if (opts.silent) {
    Logger.log(msg);
  } else {
    SpreadsheetApp.getUi().alert(msg);
  }
}

function _getLeadsSheet() {
  const ss = SpreadsheetApp.getActive();
  const savedId = PropertiesService.getDocumentProperties().getProperty(PROP_SHEET_ID);
  if (savedId) {
    const match = ss.getSheets().find(function(s) { return String(s.getSheetId()) === String(savedId); });
    if (match) return match;
  }
  // Fallback: try the constant name (for brand-new installs)
  const byName = ss.getSheetByName(SHEET_NAME_FALLBACK);
  if (byName) return byName;
  return null;
}

function _getTemplateDraft() {
  const id = PropertiesService.getDocumentProperties().getProperty(PROP_DRAFT_ID);
  if (!id) {
    return { ok: false, error: "No template draft picked.\nUse 'Free Mail Merge → 🎯 Pick template draft' first." };
  }
  try {
    const m = GmailApp.getDraft(id).getMessage();
    return { ok: true, data: {
      subject: m.getSubject() || "",
      htmlBody: m.getBody(),
      plainBody: m.getPlainBody(),
      attachments: m.getAttachments(),
    }};
  } catch (e) {
    return { ok: false, error: "The template draft no longer exists.\nPick another with 'Free Mail Merge → 🎯 Pick template draft'." };
  }
}

function _mergeTemplate(tmpl, vars) {
  const subst = (s) => (s || "")
    .replace(/\{\{first_name\}\}/g, vars.first_name || "")
    .replace(/\{\{company\}\}/g, vars.company || "")
    .replace(/\{\{sector_huma\}\}/g, vars.sector_huma || "");
  return {
    subject: subst(tmpl.subject),
    htmlBody: subst(tmpl.htmlBody),
    plainBody: subst(tmpl.plainBody),
  };
}

// ---------- Utilities ----------

function resetSent() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert("Reset all sends", "Clear sent_at/sent_status/error for ALL leads? Next batch will resend them.", ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const sheet = _getLeadsSheet();
  if (!sheet) { ui.alert("No leads sheet picked."); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, COL.sent_at, lastRow - 1, 3).clearContent();
  }
  ui.alert("Reset done.");
}

function showStatus() {
  const sheet = _getLeadsSheet();
  if (!sheet) { SpreadsheetApp.getUi().alert("No leads sheet picked."); return; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
  let total = 0, sent = 0, err = 0, pending = 0, lkPending = 0, noLKPending = 0;
  for (const row of data) {
    const email = row[COL.email - 1];
    if (!email) continue;
    total++;
    const status = String(row[COL.sent_status - 1] || "").trim();
    const lk = String(row[COL.lk_contacted - 1] || "").trim().toUpperCase();
    if (status === "sent") sent++;
    else if (status === "error") err++;
    else {
      pending++;
      if (lk === "YES") lkPending++;
      else noLKPending++;
    }
  }
  SpreadsheetApp.getUi().alert(
    "Total leads: " + total + "\n" +
    "  Sent: " + sent + "\n" +
    "  Errors: " + err + "\n" +
    "  Pending: " + pending + "\n" +
    "    No LK: " + noLKPending + "\n" +
    "    LK: " + lkPending
  );
}

// ---------- Scheduling ----------

const SCHEDULED_HANDLER = "_scheduledSend";
const SCHEDULED_DAILY_HANDLER = "_scheduledDailySend";

function scheduleOneTime() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt("Schedule one-time batch",
    "When? Use format YYYY-MM-DD HH:MM (24h, script timezone: " + Session.getScriptTimeZone() + ")\n\nExample: 2026-04-24 09:30",
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;

  const when = (r1.getResponseText() || "").trim();
  const date = _parseScheduleDate(when);
  if (!date) { ui.alert("Invalid date format. Use YYYY-MM-DD HH:MM"); return; }
  if (date.getTime() <= Date.now()) { ui.alert("That time is in the past."); return; }

  const r2 = ui.prompt("Batch size", "How many emails to send at " + when + "? (default " + BATCH_SIZE_DEFAULT + ")", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const limit = parseInt(r2.getResponseText(), 10) || BATCH_SIZE_DEFAULT;

  PropertiesService.getDocumentProperties().setProperty(PROP_SCHEDULE_LIMIT, String(limit));

  ScriptApp.newTrigger(SCHEDULED_HANDLER).timeBased().at(date).create();
  ui.alert("✅ Scheduled " + limit + " emails for " + when + ".\n\nThe script will run unattended at that time. Make sure a template draft is picked.");
}

function scheduleDaily() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt("Schedule daily batch",
    "At what hour should it run every day? (0-23, script timezone: " + Session.getScriptTimeZone() + ")\n\nExample: 9",
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const hour = parseInt(r1.getResponseText(), 10);
  if (isNaN(hour) || hour < 0 || hour > 23) { ui.alert("Invalid hour. Use 0-23."); return; }

  const r2 = ui.prompt("Daily batch size", "How many emails per day? (default " + BATCH_SIZE_DEFAULT + ")", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const limit = parseInt(r2.getResponseText(), 10) || BATCH_SIZE_DEFAULT;

  PropertiesService.getDocumentProperties().setProperty(PROP_SCHEDULE_DAILY_LIMIT, String(limit));

  // Wipe any previous daily triggers so we don't duplicate
  const existing = ScriptApp.getProjectTriggers();
  for (const t of existing) {
    if (t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER) ScriptApp.deleteTrigger(t);
  }

  ScriptApp.newTrigger(SCHEDULED_DAILY_HANDLER).timeBased().everyDays(1).atHour(hour).create();
  ui.alert("✅ Daily batch of " + limit + " emails scheduled at " + hour + ":00.\n\nCancel anytime via 'Cancel all scheduled jobs'.");
}

function listSchedules() {
  const triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === SCHEDULED_HANDLER || t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER;
  });
  if (!triggers.length) { SpreadsheetApp.getUi().alert("No scheduled jobs."); return; }

  const props = PropertiesService.getDocumentProperties();
  const oneTimeLimit = props.getProperty(PROP_SCHEDULE_LIMIT) || BATCH_SIZE_DEFAULT;
  const dailyLimit = props.getProperty(PROP_SCHEDULE_DAILY_LIMIT) || BATCH_SIZE_DEFAULT;

  const lines = triggers.map(function(t) {
    const fn = t.getHandlerFunction();
    if (fn === SCHEDULED_DAILY_HANDLER) return "🔁 Daily — " + dailyLimit + " emails/day";
    return "📅 One-time — " + oneTimeLimit + " emails (trigger id " + t.getUniqueId() + ")";
  });
  SpreadsheetApp.getUi().alert("Scheduled jobs:\n\n" + lines.join("\n"));
}

function cancelSchedules() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert("Cancel all scheduled jobs", "Remove every one-time and daily scheduled batch?", ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  let n = 0;
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === SCHEDULED_HANDLER || t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER) {
      ScriptApp.deleteTrigger(t); n++;
    }
  }
  ui.alert("Cancelled " + n + " scheduled job(s).");
}

// Trigger handlers — run unattended, no UI.
function _scheduledSend() {
  const limit = parseInt(PropertiesService.getDocumentProperties().getProperty(PROP_SCHEDULE_LIMIT), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, skipLK: true, silent: true });
  // One-time trigger: remove itself so it doesn't linger.
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === SCHEDULED_HANDLER) ScriptApp.deleteTrigger(t);
  }
}

function _scheduledDailySend() {
  const limit = parseInt(PropertiesService.getDocumentProperties().getProperty(PROP_SCHEDULE_DAILY_LIMIT), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, skipLK: true, silent: true });
}

function _parseScheduleDate(s) {
  // Accepts "YYYY-MM-DD HH:MM". Returns Date in script timezone or null.
  const m = /^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})$/.exec(s);
  if (!m) return null;
  const d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), parseInt(m[4]), parseInt(m[5]), 0);
  return isNaN(d.getTime()) ? null : d;
}

// ---------- Reply tracking ----------

function checkReplies() {
  const sheet = _getLeadsSheet();
  if (!sheet) { SpreadsheetApp.getUi().alert("No leads sheet picked."); return; }

  // Make sure the replied_at header exists in column 16
  const header = sheet.getRange(1, COL.replied_at).getValue();
  if (!header) sheet.getRange(1, COL.replied_at).setValue("replied_at");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No leads."); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const emailToRow = {};
  for (let i = 0; i < data.length; i++) {
    const email = String(data[i][COL.email - 1] || "").trim().toLowerCase();
    const sentAt = data[i][COL.sent_at - 1];
    const alreadyReplied = data[i][COL.replied_at - 1];
    if (!email || !sentAt || alreadyReplied) continue;
    emailToRow[email] = i + 2; // sheet row
  }
  const pending = Object.keys(emailToRow);
  if (!pending.length) { SpreadsheetApp.getUi().alert("Nothing to check (no sent rows without a reply)."); return; }

  // Pull recent inbox threads. Limit to last 14 days.
  const query = "in:inbox newer_than:14d";
  const threads = GmailApp.search(query, 0, 200);
  let marked = 0;

  for (const t of threads) {
    const msgs = t.getMessages();
    for (const m of msgs) {
      const from = (m.getFrom() || "").toLowerCase();
      // Extract email inside "Name <email@x>" or bare "email@x"
      const match = from.match(/<([^>]+)>/);
      const addr = (match ? match[1] : from).trim();
      if (!addr || !emailToRow[addr]) continue;
      const row = emailToRow[addr];
      const ts = Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      sheet.getRange(row, COL.replied_at).setValue(ts);
      delete emailToRow[addr];
      marked++;
      break;
    }
  }
  SpreadsheetApp.getUi().alert("Replies detected: " + marked + "\nRows updated in column " + COL.replied_at + " (replied_at).");
}
