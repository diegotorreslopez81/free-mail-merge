/**
 * Free Mail Merge
 *
 * A free Google Apps Script that turns a Sheet + a Gmail draft into a mail merge tool.
 * Drop-in replacement for YAMM. Paste this file in Extensions → Apps Script.
 *
 * Setup:
 *   1. Paste this file in Extensions → Apps Script → save (Cmd+S) → reload Sheet.
 *   2. Menu "✉️ Free Mail Merge" appears.
 *   3. Click "⚙️ Settings" once to enter your From name, alias and test email.
 *
 * Usage:
 *   - Create a draft in Gmail using placeholders like {{first_name}}, {{company}} (any column header).
 *   - Menu → 🗂 Pick leads sheet.
 *   - Menu → 🎯 Pick template draft.
 *   - Menu → 📨 Send test email.
 *   - Menu → 📧 Send batch (or ⏰ Schedule).
 *
 * Repo: https://github.com/diegotorreslopez81/free-mail-merge
 */

// ----- Fallback defaults. Real values are configured from the menu (⚙️ Settings). -----
const FROM_NAME = "Your Name · Your Company";       // display name in recipient's inbox
const REPLY_TO = "you@yourdomain.com";              // Reply-To + From alias (must be a Send-As alias if different from login)
const TEST_EMAIL_DEFAULT = "you@yourdomain.com";    // default destination for "Send test"
const SHEET_NAME_FALLBACK = "Leads";                // fallback tab name if none is picked
// ---------------------------------------------------------------------------------------

const SCHEDULE_TAB_NAME = "_Schedule";              // auto-generated, read-only tab
const BATCH_SIZE_DEFAULT = 30;                      // Apps Script has a 6-min timeout
const DELAY_MIN_MS = 5000;                          // 5 s random delay floor
const DELAY_MAX_MS = 15000;                         // 15 s random delay ceiling
const QUOTA_SAFETY_MARGIN = 3;                      // stop a batch when this many quota slots remain

// DocumentProperties keys (never rename — would drop user settings on update)
const PROP_DRAFT_ID = "TEMPLATE_DRAFT_ID";
const PROP_SHEET_ID = "LEADS_SHEET_ID";
const PROP_FROM_NAME = "SETTING_FROM_NAME";
const PROP_REPLY_TO = "SETTING_REPLY_TO";
const PROP_TEST_EMAIL = "SETTING_TEST_EMAIL";
const PROP_INJECT_UNSUB = "SETTING_INJECT_UNSUB";  // "1" = on, "0" = off. Default on.
const PROP_SCHEDULE_LIMIT = "SCHEDULE_LIMIT";
const PROP_SCHEDULE_DAILY_LIMIT = "SCHEDULE_DAILY_LIMIT";
const PROP_SCHED_META_PREFIX = "SCHED_META_";       // per-trigger metadata
const PROP_VERIFIER_URL = "SETTING_VERIFIER_URL";   // optional SMTP verifier endpoint (smtp_verifier.py --serve)

const SUPPRESSION_TAB_NAME = "_Suppression";        // auto-generated unsubscribe list

// Column layout. Columns A-L are user lead data (auto-discovered from the row 1
// headers for placeholder replacement). Columns M-U are auto-written by the script.
const COL = {
  email: 1,            // column A must be email
  sent_at: 13,
  sent_status: 14,
  error: 15,
  replied_at: 16,
  opened_at: 17,
  clicked_at: 18,
  unsubscribed_at: 19,
  bounced_at: 20,
  status: 21           // YAMM-style emoji summary of the row
};

const TRACKING_HEADERS = [
  { col: 13, name: "sent_at" },
  { col: 14, name: "sent_status" },
  { col: 15, name: "error" },
  { col: 16, name: "replied_at" },
  { col: 17, name: "opened_at" },
  { col: 18, name: "clicked_at" },
  { col: 19, name: "unsubscribed_at" },
  { col: 20, name: "bounced_at" },
  { col: 21, name: "status" }
];

// ---------- Menu ----------


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("✉️ Free Mail Merge")
    .addItem("🚀 Setup", "openSetup")
    .addSeparator()
    .addItem("📨 Send test email", "sendTest")
    .addItem("📧 Send batch", "sendBatch")
    .addSubMenu(ui.createMenu("⏰ Schedule")
      .addItem("📅 One-time batch", "scheduleOneTime")
      .addItem("🔁 Daily batch", "scheduleDaily")
      .addSeparator()
      .addItem("📋 Refresh schedule tab", "refreshScheduleTab")
      .addItem("🗑️  Cancel all scheduled", "cancelSchedules"))
    .addSeparator()
    .addItem("📊 Status", "showStatus")
    .addItem("🔄 Refresh lead statuses", "refreshStatuses")
    .addSeparator()
    .addItem("✅ Verify emails", "verifyEmails")
    .addSeparator()
    .addSubMenu(ui.createMenu("🛠️  More")
      .addItem("⚙️  Settings", "openSettings")
      .addItem("🗂  Pick leads sheet", "chooseSheet")
      .addItem("🎯 Pick template draft", "chooseTemplate")
      .addItem("🌐 Setup tracking web app", "showWebAppSetup")
      .addSeparator()
      .addItem("🧹 Reset errored rows", "resetErrors")
      .addItem("↩️  Reset all sends", "resetSent")
      .addSeparator()
      .addItem("🔍 Dry run (preview 3 in logs)", "dryRun"))
    .addToUi();
}

// ---------- Settings ----------

function _getSetting(propKey, fallback) {
  const v = PropertiesService.getDocumentProperties().getProperty(propKey);
  return (v == null || v === "") ? fallback : v;
}

function openSettings() {
  const fromName = _getSetting(PROP_FROM_NAME, FROM_NAME);
  const replyTo = _getSetting(PROP_REPLY_TO, REPLY_TO);
  const testEmail = _getSetting(PROP_TEST_EMAIL, TEST_EMAIL_DEFAULT);
  const injectUnsub = _getSetting(PROP_INJECT_UNSUB, "1") !== "0";
  const verifierUrl = _getSetting(PROP_VERIFIER_URL, "");

  const esc = function(s) { return String(s || "").replace(/"/g, "&quot;").replace(/</g, "&lt;"); };

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:20px;color:#111;">' +
    '<h2 style="margin:0 0 4px 0;">✉️ Settings</h2>' +
    '<p style="color:#666;margin:0 0 20px 0;font-size:13px;">These replace the defaults defined at the top of Code.gs. Stored in the spreadsheet, not in code, so pasting a new version of the script won\'t overwrite them.</p>' +

    '<label style="display:block;font-size:12px;color:#333;margin-bottom:4px;font-weight:600;">FROM name (display name)</label>' +
    '<input id="fromName" type="text" value="' + esc(fromName) + '" style="width:100%;padding:8px;font-size:14px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;">' +

    '<label style="display:block;font-size:12px;color:#333;margin-bottom:4px;font-weight:600;">Reply-To / From alias</label>' +
    '<input id="replyTo" type="email" value="' + esc(replyTo) + '" style="width:100%;padding:8px;font-size:14px;border:1px solid #ccc;border-radius:4px;margin-bottom:4px;box-sizing:border-box;">' +
    '<p style="color:#888;margin:0 0 16px 0;font-size:11px;">Must exist as "Send mail as" in Gmail Settings → Accounts if different from the login address.</p>' +

    '<label style="display:block;font-size:12px;color:#333;margin-bottom:4px;font-weight:600;">Default test email destination</label>' +
    '<input id="testEmail" type="email" value="' + esc(testEmail) + '" style="width:100%;padding:8px;font-size:14px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;">' +

    '<label style="display:flex;align-items:flex-start;gap:8px;font-size:13px;color:#333;margin-bottom:4px;cursor:pointer;">' +
      '<input id="injectUnsub" type="checkbox" ' + (injectUnsub ? "checked" : "") + ' style="margin-top:3px;">' +
      '<span><strong>Append automatic unsubscribe footer</strong><br>' +
      '<span style="color:#888;font-size:12px;">Adds "Don\'t want to hear from us? Unsubscribe here" at the end of each email. Turn off if your template already has its own unsubscribe instruction.</span></span>' +
    '</label>' +
    '<div style="margin-bottom:20px;"></div>' +

    '<label style="display:block;font-size:12px;color:#333;margin-bottom:4px;font-weight:600;">SMTP verifier endpoint URL <span style="color:#888;font-weight:400;">(optional)</span></label>' +
    '<input id="verifierUrl" type="url" value="' + esc(verifierUrl) + '" placeholder="https://verifier.example.com" style="width:100%;padding:8px;font-size:14px;border:1px solid #ccc;border-radius:4px;margin-bottom:4px;box-sizing:border-box;">' +
    '<p style="color:#888;margin:0 0 16px 0;font-size:11px;">If you run smtp_verifier.py from this repo as an HTTPS server, paste its base URL. Leave empty to use the free DNS-only check (drops fake domains).</p>' +


    '<div style="text-align:right;">' +
      '<button onclick="google.script.host.close()" style="padding:8px 14px;margin-right:8px;background:#f5f5f5;border:1px solid #ccc;border-radius:4px;cursor:pointer;">Cancel</button>' +
      '<button onclick="save()" style="padding:8px 14px;background:#DA291C;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:600;">Save</button>' +
    '</div>' +

    '<script>' +
    'function save() {' +
    '  const v = {' +
    '    fromName: document.getElementById("fromName").value,' +
    '    replyTo: document.getElementById("replyTo").value,' +
    '    testEmail: document.getElementById("testEmail").value,' +
    '    injectUnsub: document.getElementById("injectUnsub").checked ? "1" : "0",' +
    '    verifierUrl: document.getElementById("verifierUrl").value' +
    '  };' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).saveSettings(v);' +
    '}' +
    '</script>' +
    '</div>'
  ).setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Settings");
}

function saveSettings(v) {
  const props = PropertiesService.getDocumentProperties();
  if (v && v.fromName !== undefined) props.setProperty(PROP_FROM_NAME, v.fromName);
  if (v && v.replyTo !== undefined) props.setProperty(PROP_REPLY_TO, v.replyTo);
  if (v && v.testEmail !== undefined) props.setProperty(PROP_TEST_EMAIL, v.testEmail);
  if (v && v.injectUnsub !== undefined) props.setProperty(PROP_INJECT_UNSUB, v.injectUnsub);
  if (v && v.verifierUrl !== undefined) props.setProperty(PROP_VERIFIER_URL, v.verifierUrl);
  return true;
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

function sendTest() {
  const ui = SpreadsheetApp.getUi();
  const defaultTo = _getSetting(PROP_TEST_EMAIL, TEST_EMAIL_DEFAULT);
  const r = ui.prompt("Send test", "To which email? (default " + defaultTo + ")", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const testEmail = (r.getResponseText() || "").trim() || defaultTo;

  const tmpl = _getTemplateDraft();
  if (!tmpl.ok) { ui.alert(tmpl.error); return; }

  const sheet = _getLeadsSheet();
  if (!sheet) { ui.alert("No leads sheet picked. Use '🗂 Pick leads sheet' first."); return; }

  const headers = _getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  let vars = null;
  if (lastRow >= 2) {
    const all = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (const row of all) {
      if (row[COL.email - 1] && !row[COL.sent_at - 1]) {
        vars = _buildVars(headers, row);
        break;
      }
    }
  }
  if (!vars) vars = _sampleVars(headers);

  const merged = _mergeTemplate(tmpl.data, vars);
  const fromName = _getSetting(PROP_FROM_NAME, FROM_NAME);
  const replyTo = _getSetting(PROP_REPLY_TO, REPLY_TO);
  const webAppUrl = _getWebAppUrl();
  const injected = _injectTracking(merged.htmlBody, merged.plainBody, testEmail.toLowerCase(), webAppUrl);
  try {
    GmailApp.sendEmail(testEmail, "[TEST] " + merged.subject, injected.plain, {
      htmlBody: injected.html,
      attachments: tmpl.data.attachments,
      name: fromName,
      from: replyTo,
      replyTo: replyTo,
    });
    const trackNote = webAppUrl
      ? "\n\nTracking is active. Open the email, click a link, try the unsubscribe — then run '📊 Tracking status'."
      : "\n\n⚠️  Web app not deployed. This test went out WITHOUT tracking.";
    ui.alert("✅ Test sent to " + testEmail + "\n\nFrom alias: " + replyTo + "\n\nVerify: HTML signature, attachment, placeholders all replaced, links clickable." + trackNote);
  } catch (e) {
    ui.alert("❌ Error: " + e + "\n\nCommon causes: the alias '" + replyTo + "' isn't configured as 'Send mail as' in Gmail Settings → Accounts.");
  }
}

// ---------- Entry points ----------

function dryRun() {
  _run({ dryRun: true, limit: 3 });
}

function sendBatch() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Send batch", "How many emails? (recommended max 30-40 per run, Apps Script has a 6-min cap)", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const limit = parseInt(r.getResponseText(), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit });
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
  if (!tmpl.ok) {
    if (opts.silent) { Logger.log(tmpl.error); } else { SpreadsheetApp.getUi().alert(tmpl.error); }
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    const msg = "No leads in the sheet.";
    if (opts.silent) { Logger.log(msg); } else { SpreadsheetApp.getUi().alert(msg); }
    return;
  }

  const headers = _getHeaders(sheet);
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const fromName = _getSetting(PROP_FROM_NAME, FROM_NAME);
  const replyTo = _getSetting(PROP_REPLY_TO, REPLY_TO);
  const webAppUrl = _getWebAppUrl();

  if (!opts.dryRun) {
    _ensureTrackingHeaders(sheet);
    _ensureSuppressionTab();
  }

  let sent = 0;
  let skipped_quota = false;
  let skipped_suppressed = 0;
  const errors = [];

  // Quota awareness: stop before we get pushed off the cliff.
  let remaining = Infinity;
  try { remaining = MailApp.getRemainingDailyQuota(); } catch (e) { /* ignore */ }

  for (let i = 0; i < data.length && sent < opts.limit; i++) {
    const row = data[i];
    const sheetRow = i + 2;
    const email = String(row[COL.email - 1] || "").trim();
    const alreadySent = String(row[COL.sent_at - 1] || "").trim();

    if (!email || !email.includes("@")) continue;
    if (alreadySent) continue;
    if (_isSuppressed(email)) { skipped_suppressed++; continue; }

    if (!opts.dryRun && remaining <= QUOTA_SAFETY_MARGIN) {
      skipped_quota = true;
      break;
    }

    const vars = _buildVars(headers, row);
    const merged = _mergeTemplate(tmpl.data, vars);

    // Inject tracking (pixel + link wrapping + unsubscribe footer). Skipped if
    // no Web App URL is available.
    const injected = _injectTracking(merged.htmlBody, merged.plainBody, email.toLowerCase(), webAppUrl);

    if (opts.dryRun) {
      Logger.log("=== DRY RUN [" + (sent + 1) + "] ===");
      Logger.log("To: " + email);
      Logger.log("Subject: " + merged.subject);
      Logger.log(injected.plain.substring(0, 600) + "...\n");
    } else {
      try {
        GmailApp.sendEmail(email, merged.subject, injected.plain, {
          htmlBody: injected.html,
          attachments: tmpl.data.attachments,
          name: fromName,
          from: replyTo,
          replyTo: replyTo,
        });
        remaining = Math.max(0, remaining - 1);
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
      _updateStatusIcon(sheet, sheetRow);
      const delay = DELAY_MIN_MS + Math.floor(Math.random() * (DELAY_MAX_MS - DELAY_MIN_MS));
      Utilities.sleep(delay);
    }
    sent++;
  }

  let msg;
  if (opts.dryRun) {
    msg = "Dry run OK. " + sent + " emails rendered. Check Executions → Logs.";
  } else {
    msg = "Sent: " + sent;
    if (errors.length) msg += "\nErrors: " + errors.length + " (see 'error' column)";
    if (skipped_suppressed) msg += "\nSkipped (suppression list): " + skipped_suppressed;
    if (skipped_quota) msg += "\n\n⚠️ Stopped early: Gmail daily quota reached (" + QUOTA_SAFETY_MARGIN + " slot safety margin). Remaining: " + remaining + ".";
    if (!webAppUrl) msg += "\n\nℹ️  Web app not deployed. Emails went out WITHOUT tracking (no pixel, no unsubscribe link).";
  }
  if (opts.silent) { Logger.log(msg); } else { SpreadsheetApp.getUi().alert(msg); }
}

// ---------- Header + vars helpers ----------

function _getHeaders(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return row.map(function(v) { return String(v || "").trim(); });
}

function _buildVars(headers, row) {
  const vars = {};
  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (!key) continue;
    vars[key] = row[i] == null ? "" : String(row[i]);
  }
  return vars;
}

function _sampleVars(headers) {
  // Dummy values when no real lead is available (e.g., send test on empty sheet).
  const vars = {};
  for (const h of headers) { if (h) vars[h] = "[" + h + "]"; }
  return vars;
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
  // Dynamic: replace {{any_key}} where `any_key` is any column header from the leads sheet.
  const subst = function(s) {
    return (s || "").replace(/\{\{\s*([^}\s]+)\s*\}\}/g, function(match, key) {
      return Object.prototype.hasOwnProperty.call(vars, key) ? vars[key] : match;
    });
  };
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

function resetErrors() {
  const ui = SpreadsheetApp.getUi();
  const sheet = _getLeadsSheet();
  if (!sheet) { ui.alert("No leads sheet picked."); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert("No leads."); return; }

  const range = sheet.getRange(2, COL.sent_at, lastRow - 1, 3);
  const values = range.getValues();
  let cleared = 0;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1] || "").trim() === "error") {
      values[i] = ["", "", ""];
      cleared++;
    }
  }
  if (!cleared) { ui.alert("No errored rows found."); return; }
  const r = ui.alert("Reset errored rows", "Clear sent_at/status/error on " + cleared + " rows that previously errored? Next batch will retry them.", ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  range.setValues(values);
  ui.alert("Cleared " + cleared + " errored rows.");
}

function showStatus() {
  const ss = SpreadsheetApp.getActive();
  const sheet = _getLeadsSheet();

  // Config
  const fromName = _getSetting(PROP_FROM_NAME, FROM_NAME);
  const replyTo = _getSetting(PROP_REPLY_TO, REPLY_TO);
  const testEmail = _getSetting(PROP_TEST_EMAIL, TEST_EMAIL_DEFAULT);
  const injectUnsub = _getSetting(PROP_INJECT_UNSUB, "1") !== "0";
  const draftId = PropertiesService.getDocumentProperties().getProperty(PROP_DRAFT_ID);
  let draftSubject = "(none)";
  if (draftId) {
    try { draftSubject = GmailApp.getDraft(draftId).getMessage().getSubject(); } catch (e) { draftSubject = "(draft no longer exists)"; }
  }
  const sheetName = sheet ? sheet.getName() : "(not picked)";
  const webAppUrl = _getWebAppUrl();
  let quota = "";
  try { quota = String(MailApp.getRemainingDailyQuota()); } catch (e) { quota = "(n/a)"; }

  // Campaign aggregates
  let total = 0, sent = 0, err = 0, pending = 0;
  let opened = 0, clicked = 0, replied = 0, unsub = 0, bounced = 0;
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, COL.bounced_at).getValues();
      for (const row of data) {
        if (!row[COL.email - 1]) continue;
        total++;
        const st = String(row[COL.sent_status - 1] || "").trim();
        if (st === "sent") sent++;
        else if (st === "error") err++;
        else pending++;
        if (row[COL.opened_at - 1]) opened++;
        if (row[COL.clicked_at - 1]) clicked++;
        if (row[COL.replied_at - 1]) replied++;
        if (row[COL.unsubscribed_at - 1]) unsub++;
        if (row[COL.bounced_at - 1]) bounced++;
      }
    }
  }
  const pct = (n) => (sent ? Math.round(100 * n / sent) + "%" : "-");

  const suppressTab = ss.getSheetByName(SUPPRESSION_TAB_NAME);
  const supCount = suppressTab ? Math.max(0, suppressTab.getLastRow() - 1) : 0;

  const scheduleTab = ss.getSheetByName(SCHEDULE_TAB_NAME);
  const scheduledCount = ScriptApp.getProjectTriggers().filter(function(t){
    return t.getHandlerFunction() === SCHEDULED_HANDLER || t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER;
  }).length;

  const esc = (s) => String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const row = (label, value) => '<tr><td style="color:#888;padding:2px 12px 2px 0;white-space:nowrap;">' + label + '</td><td style="padding:2px 0;">' + esc(value) + '</td></tr>';

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:20px;color:#222;">' +

    '<h3 style="margin:0 0 4px 0;">📊 Status</h3>' +
    '<p style="color:#888;margin:0 0 20px 0;font-size:12px;">All the state of this campaign in one place.</p>' +

    '<h4 style="margin:0 0 6px 0;font-size:13px;color:#555;">Configuration</h4>' +
    '<table style="font-size:13px;margin-bottom:20px;">' +
      row("From name",  fromName) +
      row("Reply-To",   replyTo) +
      row("Test email", testEmail) +
      row("Unsub footer", injectUnsub ? "injected (Unsubscribe here)" : "disabled") +
      row("Leads tab",  sheetName) +
      row("Template",   draftSubject) +
    '</table>' +

    '<h4 style="margin:0 0 6px 0;font-size:13px;color:#555;">Campaign</h4>' +
    '<table style="font-size:13px;margin-bottom:20px;">' +
      row("Total leads", total) +
      row("✉️  Sent",    sent) +
      row("❌ Errors",   err) +
      row("⏳ Pending",  pending) +
      row("👁  Opened",  opened + "  (" + pct(opened) + " of sent)") +
      row("🔗 Clicked",  clicked + "  (" + pct(clicked) + " of sent)") +
      row("💬 Replied",  replied + "  (" + pct(replied) + " of sent)") +
      row("🚫 Unsub",    unsub) +
      row("📬 Bounced",  bounced) +
    '</table>' +

    '<h4 style="margin:0 0 6px 0;font-size:13px;color:#555;">Tracking & infra</h4>' +
    '<table style="font-size:13px;margin-bottom:20px;">' +
      row("Web app", webAppUrl ? "✅ deployed" : "⚠️ not deployed") +
      (webAppUrl ? row("URL", webAppUrl) : "") +
      row("Gmail quota today", quota) +
      row("Suppression list", supCount + " address" + (supCount === 1 ? "" : "es")) +
      row("Scheduled jobs", scheduledCount + (scheduleTab ? "  (see '" + SCHEDULE_TAB_NAME + "' tab)" : "")) +
    '</table>' +

    '<div style="text-align:right;">' +
      '<button onclick="google.script.host.close()" style="padding:8px 16px;background:#DA291C;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:600;">Close</button>' +
    '</div>' +
    '</div>'
  ).setWidth(560).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Status");
}

// ---------- Status icon column (YAMM-style per-lead visual) ----------

function _updateStatusIcon(sheet, sheetRow) {
  if (!sheet || !sheetRow) return;
  try {
    const cells = sheet.getRange(sheetRow, COL.sent_at, 1, COL.status - COL.sent_at + 1).getValues()[0];
    // indexes within `cells` (0-based): sent_at, sent_status, error, replied_at, opened_at, clicked_at, unsubscribed_at, bounced_at, status
    const iconFor = {
      0: { val: cells[0], emoji: "✉️" },  // sent_at → sent
      3: { val: cells[3], emoji: "💬" },  // replied_at
      4: { val: cells[4], emoji: "👁" },  // opened_at
      5: { val: cells[5], emoji: "🔗" },  // clicked_at
      6: { val: cells[6], emoji: "🚫" },  // unsubscribed_at
      7: { val: cells[7], emoji: "📬" },  // bounced_at
    };
    const icons = [];
    if (String(cells[1] || "") === "error") icons.push("❌");
    else if (cells[0]) icons.push("✉️");
    if (cells[4]) icons.push("👁");
    if (cells[5]) icons.push("🔗");
    if (cells[3]) icons.push("💬");
    if (cells[6]) icons.push("🚫");
    if (cells[7]) icons.push("📬");
    sheet.getRange(sheetRow, COL.status).setValue(icons.join(" "));
  } catch (e) { Logger.log("_updateStatusIcon: " + e); }
}

function refreshAllStatusIcons(sheet) {
  sheet = sheet || _getLeadsSheet();
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  _ensureTrackingHeaders(sheet);
  const width = COL.status - COL.sent_at + 1;
  const data = sheet.getRange(2, COL.sent_at, lastRow - 1, width).getValues();
  const out = data.map(function(cells) {
    const icons = [];
    if (String(cells[1] || "") === "error") icons.push("❌");
    else if (cells[0]) icons.push("✉️");
    if (cells[4]) icons.push("👁");
    if (cells[5]) icons.push("🔗");
    if (cells[3]) icons.push("💬");
    if (cells[6]) icons.push("🚫");
    if (cells[7]) icons.push("📬");
    return [icons.join(" ")];
  });
  sheet.getRange(2, COL.status, out.length, 1).setValues(out);
  return out.length;
}

function refreshStatuses() {
  // One-shot refresh: scans inbox for replies + bounces, then rebuilds the icon column.
  const sheet = _getLeadsSheet();
  if (!sheet) { SpreadsheetApp.getUi().alert("No leads sheet picked."); return; }
  let repliesMarked = 0, bouncesMarked = 0;
  try { repliesMarked = _scanReplies(sheet); } catch (e) { Logger.log(e); }
  try { bouncesMarked = _scanBounces(sheet); } catch (e) { Logger.log(e); }
  const iconsRefreshed = refreshAllStatusIcons(sheet);
  SpreadsheetApp.getUi().alert(
    "Lead statuses refreshed.\n\n" +
    "💬 New replies detected: " + repliesMarked + "\n" +
    "📬 New bounces detected: " + bouncesMarked + "\n" +
    "🔄 Rows with icons: " + iconsRefreshed
  );
}

// ---------- Setup checklist ----------

function openSetup() {
  const props = PropertiesService.getDocumentProperties();
  const hasSettings = !!(props.getProperty(PROP_FROM_NAME) && props.getProperty(PROP_REPLY_TO));
  const hasSheet = !!_getLeadsSheet();
  const hasTemplate = !!props.getProperty(PROP_DRAFT_ID);
  const hasWebApp = !!_getWebAppUrl();

  const check = (ok) => ok ? '<span style="color:#10B981;">✓</span>' : '<span style="color:#aaa;">○</span>';
  const btn = (label, fn, primary) => '<button onclick="run(\'' + fn + '\')" style="padding:6px 12px;font-size:13px;background:' + (primary ? "#DA291C" : "#f5f5f5") + ';color:' + (primary ? "white" : "#222") + ';border:1px solid ' + (primary ? "#DA291C" : "#ccc") + ';border-radius:4px;cursor:pointer;">' + label + '</button>';
  const row = (ok, title, desc, fn) => '<tr>' +
    '<td style="padding:10px 10px 10px 0;vertical-align:top;font-size:18px;">' + check(ok) + '</td>' +
    '<td style="padding:10px 0;vertical-align:top;">' +
      '<div style="font-weight:600;font-size:14px;">' + title + '</div>' +
      '<div style="color:#666;font-size:12px;margin-top:2px;">' + desc + '</div>' +
    '</td>' +
    '<td style="padding:10px 0 10px 12px;vertical-align:top;text-align:right;">' + btn(ok ? "Edit" : "Setup", fn, !ok) + '</td>' +
  '</tr>';

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:20px;color:#222;">' +
    '<h2 style="margin:0 0 4px 0;">🚀 Setup</h2>' +
    '<p style="color:#666;margin:0 0 16px 0;font-size:13px;">Complete the four steps below. Order matters only between sheet and template. Come back anytime to edit.</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
      row(hasSettings, "Settings",        "From name, reply-to alias, default test email.", "openSettings") +
      row(hasSheet,    "Leads sheet",     "Which tab holds your leads.",                    "chooseSheet") +
      row(hasTemplate, "Template draft",  "Which Gmail draft to use as template.",          "chooseTemplate") +
      row(hasWebApp,   "Tracking web app","Optional. Enables opens, clicks, unsubscribe.",   "showWebAppSetup") +
    '</table>' +
    '<div style="text-align:right;margin-top:20px;">' +
      '<button onclick="google.script.host.close()" style="padding:8px 16px;background:#222;color:white;border:none;border-radius:4px;cursor:pointer;">Done</button>' +
    '</div>' +
    '<script>' +
    'function run(fn){ google.script.run.withSuccessHandler(function(){}).dispatch(fn); google.script.host.close(); }' +
    '</script>' +
    '</div>'
  ).setWidth(560).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, "Setup");
}

// Dispatcher used by the Setup modal so a single onclick can call any setup action.
function dispatch(fn) {
  switch (fn) {
    case "openSettings":    openSettings();    break;
    case "chooseSheet":     chooseSheet();     break;
    case "chooseTemplate":  chooseTemplate();  break;
    case "showWebAppSetup": showWebAppSetup(); break;
  }
  return true;
}

// ---------- Scheduling ----------

const SCHEDULED_HANDLER = "_scheduledSend";
const SCHEDULED_DAILY_HANDLER = "_scheduledDailySend";

function scheduleOneTime() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt("Schedule one-time batch",
    "When? Use format YYYY-MM-DD HH:MM (24h, script timezone: " + Session.getScriptTimeZone() + ")\n\nExample: 2026-04-25 09:30",
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

  const trig = ScriptApp.newTrigger(SCHEDULED_HANDLER).timeBased().at(date).create();
  _saveScheduleMeta(trig.getUniqueId(), {
    type: "one-time",
    nextFire: date.toISOString(),
    limit,
    createdAt: new Date().toISOString(),
  });
  refreshScheduleTab();
  ui.alert("✅ Scheduled " + limit + " emails for " + when + ".\n\nSee the '" + SCHEDULE_TAB_NAME + "' tab for the full schedule.");
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

  // Wipe any previous daily triggers so we don't duplicate.
  const existing = ScriptApp.getProjectTriggers();
  for (const t of existing) {
    if (t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER) {
      _deleteScheduleMeta(t.getUniqueId());
      ScriptApp.deleteTrigger(t);
    }
  }

  const trig = ScriptApp.newTrigger(SCHEDULED_DAILY_HANDLER).timeBased().everyDays(1).atHour(hour).create();
  _saveScheduleMeta(trig.getUniqueId(), {
    type: "daily",
    hour,
    nextFire: _nextDailyFire(hour).toISOString(),
    limit,
    createdAt: new Date().toISOString(),
  });
  refreshScheduleTab();
  ui.alert("✅ Daily batch of " + limit + " emails scheduled at " + hour + ":00.\n\nSee the '" + SCHEDULE_TAB_NAME + "' tab for the full schedule.");
}

function cancelSchedules() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert("Cancel all scheduled jobs", "Remove every one-time and daily scheduled batch?", ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  let n = 0;
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === SCHEDULED_HANDLER || t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER) {
      _deleteScheduleMeta(t.getUniqueId());
      ScriptApp.deleteTrigger(t);
      n++;
    }
  }
  refreshScheduleTab();
  ui.alert("Cancelled " + n + " scheduled job(s).");
}

// Trigger handlers — run unattended, no UI.
function _scheduledSend(e) {
  const limit = parseInt(PropertiesService.getDocumentProperties().getProperty(PROP_SCHEDULE_LIMIT), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, silent: true });
  // One-time trigger: remove itself + its metadata.
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === SCHEDULED_HANDLER) {
      _deleteScheduleMeta(t.getUniqueId());
      ScriptApp.deleteTrigger(t);
    }
  }
  try { refreshScheduleTab(); } catch (err) { Logger.log(err); }
}

function _scheduledDailySend(e) {
  const limit = parseInt(PropertiesService.getDocumentProperties().getProperty(PROP_SCHEDULE_DAILY_LIMIT), 10) || BATCH_SIZE_DEFAULT;
  _run({ dryRun: false, limit, silent: true });
  // Advance nextFire in metadata to tomorrow for the daily trigger.
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER) {
      const meta = _getScheduleMeta(t.getUniqueId());
      if (meta && typeof meta.hour === "number") {
        meta.nextFire = _nextDailyFire(meta.hour).toISOString();
        _saveScheduleMeta(t.getUniqueId(), meta);
      }
    }
  }
  try { refreshScheduleTab(); } catch (err) { Logger.log(err); }
}

function _parseScheduleDate(s) {
  // Accepts "YYYY-MM-DD HH:MM". Returns Date in script timezone or null.
  const m = /^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})$/.exec(s);
  if (!m) return null;
  const d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), parseInt(m[4]), parseInt(m[5]), 0);
  return isNaN(d.getTime()) ? null : d;
}

function _nextDailyFire(hour) {
  const now = new Date();
  const next = new Date();
  next.setHours(hour, 0, 0, 0);
  if (next.getTime() <= now.getTime()) next.setDate(next.getDate() + 1);
  return next;
}

// Metadata helpers — one entry per trigger, keyed by uniqueId.
function _saveScheduleMeta(triggerId, meta) {
  PropertiesService.getDocumentProperties().setProperty(PROP_SCHED_META_PREFIX + triggerId, JSON.stringify(meta));
}
function _getScheduleMeta(triggerId) {
  const raw = PropertiesService.getDocumentProperties().getProperty(PROP_SCHED_META_PREFIX + triggerId);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function _deleteScheduleMeta(triggerId) {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_SCHED_META_PREFIX + triggerId);
}

// ---------- Schedule tab (read-only) ----------

function refreshScheduleTab() {
  const ss = SpreadsheetApp.getActive();
  let tab = ss.getSheetByName(SCHEDULE_TAB_NAME);
  if (!tab) {
    tab = ss.insertSheet(SCHEDULE_TAB_NAME);
    tab.setTabColor("#DA291C");
  }

  tab.clear();

  const title = [["✉️ Free Mail Merge — scheduled jobs"]];
  tab.getRange(1, 1, 1, 1).setValues(title).setFontWeight("bold").setFontSize(14);
  tab.getRange(2, 1, 1, 1).setValues([["Auto-generated. Do not edit — changes will be overwritten."]]).setFontStyle("italic").setFontColor("#888");

  const headers = [["Type", "Next fire", "Batch size", "Handler", "Trigger ID", "Created"]];
  tab.getRange(4, 1, 1, headers[0].length).setValues(headers).setFontWeight("bold").setBackground("#f0f0f0");

  const triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === SCHEDULED_HANDLER || t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER;
  });

  if (!triggers.length) {
    tab.getRange(5, 1, 1, 1).setValues([["(no jobs scheduled)"]]).setFontColor("#888");
  } else {
    const tz = Session.getScriptTimeZone();
    const rows = triggers.map(function(t) {
      const meta = _getScheduleMeta(t.getUniqueId()) || {};
      const isDaily = t.getHandlerFunction() === SCHEDULED_DAILY_HANDLER;
      let nextFire = "";
      if (isDaily && typeof meta.hour === "number") {
        nextFire = Utilities.formatDate(_nextDailyFire(meta.hour), tz, "yyyy-MM-dd HH:mm") + "  (then daily)";
      } else if (meta.nextFire) {
        nextFire = Utilities.formatDate(new Date(meta.nextFire), tz, "yyyy-MM-dd HH:mm");
      } else {
        nextFire = "(unknown — created before this version)";
      }
      const createdAt = meta.createdAt ? Utilities.formatDate(new Date(meta.createdAt), tz, "yyyy-MM-dd HH:mm") : "";
      const limit = meta.limit || PropertiesService.getDocumentProperties().getProperty(isDaily ? PROP_SCHEDULE_DAILY_LIMIT : PROP_SCHEDULE_LIMIT) || BATCH_SIZE_DEFAULT;
      return [
        isDaily ? "🔁 Daily" : "📅 One-time",
        nextFire,
        limit,
        t.getHandlerFunction(),
        t.getUniqueId(),
        createdAt,
      ];
    });
    tab.getRange(5, 1, rows.length, rows[0].length).setValues(rows);
  }

  tab.autoResizeColumns(1, 6);

  // Best-effort protection: warning-only so the user can unprotect if they want.
  try {
    const protection = tab.protect().setDescription("Auto-managed by Free Mail Merge");
    protection.setWarningOnly(true);
  } catch (e) { /* ignore in contexts where protect isn't allowed */ }
}

// ============================================================================
// TRACKING (opens, clicks, unsubscribe, bounces) — requires Web App deployment
// ============================================================================

// ---------- Web App endpoint ----------
// Apps Script calls doGet(e) when the deployed Web App URL is hit.
// We dispatch by ?t= parameter: open | click | unsub.
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const t = params.t || "";
  const recipient = (params.r || "").toLowerCase().trim();

  try {
    if (t === "open") {
      _trackEvent(recipient, COL.opened_at, "");
      return _transparentPixel();
    }
    if (t === "click") {
      const url = params.u || "";
      _trackEvent(recipient, COL.clicked_at, url);
      return _redirect(url);
    }
    if (t === "unsub") {
      _trackEvent(recipient, COL.unsubscribed_at, "unsubscribed");
      _addToSuppression(recipient, "user-unsubscribe");
      return _unsubConfirmation(recipient);
    }
  } catch (err) {
    Logger.log("doGet error: " + err);
  }
  return HtmlService.createHtmlOutput("<p>Free Mail Merge tracking endpoint.</p>").setTitle("Free Mail Merge");
}

function _transparentPixel() {
  // Apps Script can't return raw binary; email clients don't mind receiving
  // empty text content as long as the HTTP 200 fires — the pixel is a tracking
  // beacon, not a visible graphic. Hidden by width:1/height:1/display:none.
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

function _redirect(url) {
  if (!url || !/^https?:\/\//i.test(url)) {
    return HtmlService.createHtmlOutput("<p>Invalid URL.</p>");
  }
  const safe = url.replace(/"/g, "&quot;");
  const html = '<!doctype html><html><head><meta http-equiv="refresh" content="0;url=' + safe + '"><title>Redirecting...</title></head>' +
    '<body><p>Redirecting to <a href="' + safe + '">' + safe + '</a>...</p>' +
    '<script>window.location.replace("' + safe + '");</script></body></html>';
  return HtmlService.createHtmlOutput(html);
}

function _unsubConfirmation(email) {
  const safe = (email || "").replace(/</g, "&lt;");
  const html = '<!doctype html><html><head><title>Unsubscribed</title></head>' +
    '<body style="font-family:-apple-system,Helvetica,Arial,sans-serif;max-width:520px;margin:80px auto;padding:24px;color:#222;">' +
    '<h2 style="margin:0 0 12px 0;">✅ Unsubscribed</h2>' +
    '<p>The address <strong>' + safe + '</strong> has been added to our suppression list. You will not receive further emails from this campaign.</p>' +
    '<p style="color:#888;font-size:13px;margin-top:32px;">If this was a mistake, reply to any previous email and we\'ll restore you.</p>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html);
}

// ---------- Tracking helpers ----------

function _getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url || "";
  } catch (e) { return ""; }
}

function _trackEvent(recipient, colIndex, value) {
  if (!recipient) return;
  const sheet = _getLeadsSheet();
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  _ensureTrackingHeaders(sheet);
  const emails = sheet.getRange(2, COL.email, lastRow - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    const rowEmail = String(emails[i][0] || "").toLowerCase().trim();
    if (rowEmail && rowEmail === recipient) {
      const sheetRow = i + 2;
      const cur = sheet.getRange(sheetRow, colIndex).getValue();
      const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      // Write timestamp (first event wins for opens/clicks; always overwrite for unsub)
      if (!cur || colIndex === COL.unsubscribed_at) {
        sheet.getRange(sheetRow, colIndex).setValue(ts + (value ? "  |  " + value : ""));
      } else if (colIndex === COL.clicked_at && value) {
        // Click: append new URL to the cell if different
        const s = String(cur);
        if (s.indexOf(value) === -1) sheet.getRange(sheetRow, colIndex).setValue(s + "\n" + ts + "  |  " + value);
      }
      _updateStatusIcon(sheet, sheetRow);
      return;
    }
  }
}

function _ensureTrackingHeaders(sheet) {
  for (const h of TRACKING_HEADERS) {
    const cur = sheet.getRange(1, h.col).getValue();
    if (!cur) sheet.getRange(1, h.col).setValue(h.name);
  }
}

function _injectTracking(htmlBody, plainBody, recipient, webAppUrl) {
  if (!webAppUrl) return { html: htmlBody, plain: plainBody };

  const encR = encodeURIComponent(recipient);
  const injectUnsub = _getSetting(PROP_INJECT_UNSUB, "1") !== "0";

  // 1) Wrap <a href=""> links (skip mailto/tel/#anchor and our own web app)
  const html1 = (htmlBody || "").replace(/<a\s+([^>]*?)href=(["'])([^"']+)\2([^>]*)>/gi, function(match, pre, q, url, post) {
    if (/^mailto:|^tel:|^#|^javascript:/i.test(url)) return match;
    if (url.indexOf(webAppUrl) === 0) return match;
    const wrapped = webAppUrl + "?t=click&r=" + encR + "&u=" + encodeURIComponent(url);
    return "<a " + pre + "href=" + q + wrapped + q + post + ">";
  });

  // 2) Open-pixel at the end of body (or of html if no body tag)
  const pixel = '<img src="' + webAppUrl + '?t=open&r=' + encR + '" width="1" height="1" alt="" border="0" style="display:block;width:1px;height:1px;border:0;" />';
  const html2 = /<\/body>/i.test(html1) ? html1.replace(/<\/body>/i, pixel + "</body>") : (html1 + pixel);

  // 3) Unsubscribe footer — only if user hasn't disabled it (e.g. their template
  //    already carries its own unsubscribe instruction).
  let html3 = html2;
  let plain = plainBody || "";
  if (injectUnsub) {
    const unsubUrl = webAppUrl + "?t=unsub&r=" + encR;
    const footer = '<p style="color:#999;font-size:11px;margin-top:40px;border-top:1px solid #eee;padding-top:12px;">' +
      'Don\'t want to hear from us? <a href="' + unsubUrl + '" style="color:#999;">Unsubscribe here</a>.' +
      '</p>';
    html3 = /<\/body>/i.test(html2) ? html2.replace(/<\/body>/i, footer + "</body>") : (html2 + footer);
    plain = plain + "\n\n---\nUnsubscribe: " + unsubUrl;
  }
  return { html: html3, plain: plain };
}

// ---------- Suppression list ----------

function _ensureSuppressionTab() {
  const ss = SpreadsheetApp.getActive();
  let tab = ss.getSheetByName(SUPPRESSION_TAB_NAME);
  if (!tab) {
    tab = ss.insertSheet(SUPPRESSION_TAB_NAME);
    tab.setTabColor("#666666");
    tab.getRange(1, 1, 1, 3).setValues([["email", "unsubscribed_at", "source"]]).setFontWeight("bold").setBackground("#f0f0f0");
  }
  return tab;
}

function _addToSuppression(email, source) {
  if (!email) return;
  const tab = _ensureSuppressionTab();
  // Dedup
  const lastRow = tab.getLastRow();
  if (lastRow > 1) {
    const existing = tab.getRange(2, 1, lastRow - 1, 1).getValues();
    for (const row of existing) {
      if (String(row[0] || "").toLowerCase().trim() === email.toLowerCase().trim()) return;
    }
  }
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  tab.appendRow([email.toLowerCase().trim(), ts, source || ""]);
}

function _isSuppressed(email) {
  if (!email) return false;
  const ss = SpreadsheetApp.getActive();
  const tab = ss.getSheetByName(SUPPRESSION_TAB_NAME);
  if (!tab) return false;
  const lastRow = tab.getLastRow();
  if (lastRow < 2) return false;
  const emails = tab.getRange(2, 1, lastRow - 1, 1).getValues();
  const target = email.toLowerCase().trim();
  for (const row of emails) {
    if (String(row[0] || "").toLowerCase().trim() === target) return true;
  }
  return false;
}

// ---------- Bounces ----------

function _scanBounces(sheet) {
  _ensureTrackingHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const all = sheet.getRange(2, 1, lastRow - 1, COL.bounced_at).getValues();
  const emailToRow = {};
  for (let i = 0; i < all.length; i++) {
    const e = String(all[i][0] || "").toLowerCase().trim();
    const bounced = all[i][COL.bounced_at - 1];
    if (e && !bounced) emailToRow[e] = i + 2;
  }

  const threads = GmailApp.search("from:(mailer-daemon OR postmaster) in:inbox newer_than:30d");
  let marked = 0;
  for (const t of threads) {
    for (const m of t.getMessages()) {
      const body = (m.getPlainBody() || "");
      const found = body.match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/gi) || [];
      for (const candidate of found) {
        const key = candidate.toLowerCase();
        if (emailToRow[key]) {
          const row = emailToRow[key];
          const ts = Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          sheet.getRange(row, COL.bounced_at).setValue(ts);
          _addToSuppression(key, "bounce");
          _updateStatusIcon(sheet, row);
          delete emailToRow[key];
          marked++;
          break;
        }
      }
    }
  }
  return marked;
}

function checkBounces() {
  const sheet = _getLeadsSheet();
  if (!sheet) { SpreadsheetApp.getUi().alert("No leads sheet picked."); return; }
  const marked = _scanBounces(sheet);
  SpreadsheetApp.getUi().alert("Bounces detected: " + marked + "\n" +
    (marked ? "Marked in 'bounced_at' column and added to suppression list." : "Inbox scanned, no unseen bounces match your leads."));
}

// ---------- Web App UX ----------

function showWebAppSetup() {
  const existing = _getWebAppUrl();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;padding:20px;color:#222;max-width:560px;">' +
    '<h2 style="margin:0 0 8px 0;">🌐 Deploy tracking web app</h2>' +
    '<p style="color:#666;margin:0 0 16px 0;font-size:13px;">Tracking (opens, clicks, unsubscribe) needs a public URL that your emails can call. Apps Script gives you one when you deploy this script as a Web App. Takes 30 seconds, done once.</p>' +
    '<ol style="line-height:1.6;font-size:14px;">' +
    '<li>Open <strong>Extensions → Apps Script</strong>.</li>' +
    '<li>Click <strong>Deploy → New deployment</strong>.</li>' +
    '<li>Gear icon (⚙️) → pick <strong>Web app</strong>.</li>' +
    '<li>Description: <em>Free Mail Merge tracker</em>.</li>' +
    '<li>Execute as: <strong>Me</strong>.</li>' +
    '<li>Who has access: <strong>Anyone</strong>. (Required so recipients\' email clients can load the pixel.)</li>' +
    '<li>Click <strong>Deploy</strong> and authorize.</li>' +
    '<li>Copy the URL. Come back and reload this Sheet.</li>' +
    '</ol>' +
    '<p style="margin:16px 0 0 0;font-size:13px;color:' + (existing ? '#10B981' : '#DA291C') + ';"><strong>' +
    (existing ? '✅ Currently deployed: ' + existing : '⚠️ No web app detected yet.') +
    '</strong></p>' +
    '<div style="text-align:right;margin-top:20px;">' +
    '<button onclick="google.script.host.close()" style="padding:8px 14px;background:#DA291C;color:white;border:none;border-radius:4px;cursor:pointer;">Got it</button>' +
    '</div>' +
    '</div>'
  ).setWidth(600).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "Setup tracking");
}

// ---------- Reply tracking ----------

function _scanReplies(sheet) {
  _ensureTrackingHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, COL.replied_at).getValues();
  const emailToRow = {};
  for (let i = 0; i < data.length; i++) {
    const email = String(data[i][COL.email - 1] || "").trim().toLowerCase();
    const sentAt = data[i][COL.sent_at - 1];
    const alreadyReplied = data[i][COL.replied_at - 1];
    if (!email || !sentAt || alreadyReplied) continue;
    emailToRow[email] = i + 2;
  }
  if (!Object.keys(emailToRow).length) return 0;

  const threads = GmailApp.search("in:inbox newer_than:14d", 0, 200);
  let marked = 0;
  for (const t of threads) {
    const msgs = t.getMessages();
    for (const m of msgs) {
      const from = (m.getFrom() || "").toLowerCase();
      const match = from.match(/<([^>]+)>/);
      const addr = (match ? match[1] : from).trim();
      if (!addr || !emailToRow[addr]) continue;
      const row = emailToRow[addr];
      const ts = Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      sheet.getRange(row, COL.replied_at).setValue(ts);
      _updateStatusIcon(sheet, row);
      delete emailToRow[addr];
      marked++;
      break;
    }
  }
  return marked;
}

function checkReplies() {
  const sheet = _getLeadsSheet();
  if (!sheet) { SpreadsheetApp.getUi().alert("No leads sheet picked."); return; }
  const marked = _scanReplies(sheet);
  SpreadsheetApp.getUi().alert("Replies detected: " + marked + "\nRows updated in column " + COL.replied_at + " (replied_at).");
}


// ---------- Email verification ----------
//
// verifyEmails() iterates over rows in the leads sheet and writes a verifier
// status to a column called "email_verified" (auto-created if missing).
//
// Two modes:
//   1. Default (zero-config, free): DNS MX check via dns.google. Marks rows as
//      "no_mx" (domain has no mail server, definitely a fake email) or "mx_ok"
//      (domain accepts mail; existence of the specific address NOT verified).
//      Useful to drop obvious junk before sending — but most domains will pass.
//
//   2. SMTP probe (Settings → Verifier URL): if you run smtp_verifier.py
//      (shipped in this repo) as an HTTPS server, paste its URL into Settings.
//      verifyEmails() then calls that endpoint per row and gets the real
//      RCPT TO result: "verified", "not_found", "catch_all", "temp_fail".
//      Setup: see README "Email verification" section.
//
// Catch-all detection: providers like Gmail/Outlook/Yahoo accept any RCPT TO,
// so per-address verification is impossible. Marked as "catch_all" → still
// safe to send (just means we can't pre-confirm).
//
// Rate-limited to 1 request/second to stay polite to MX servers + dns.google.

const VERIFIER_HEADER = "email_verified";

function verifyEmails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = _getLeadsSheet();
  if (!sheet) { ui.alert("No leads sheet picked. Use Setup first."); return; }

  const verifierUrl = _getSetting(PROP_VERIFIER_URL, "");
  const mode = verifierUrl ? "SMTP probe (via " + verifierUrl + ")" : "DNS MX check (free, basic)";
  const ans = ui.alert(
    "Verify emails",
    "About to verify all leads with no email_verified value yet.\n\n" +
    "Mode: " + mode + "\n\n" +
    "Tip: for real SMTP RCPT TO checks, deploy smtp_verifier.py from this repo and paste its URL in Settings.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (ans !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert("No leads."); return; }

  // Find or create the email_verified column.
  const headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let verifyCol = headersRow.indexOf(VERIFIER_HEADER) + 1;
  if (!verifyCol) {
    verifyCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, verifyCol).setValue(VERIFIER_HEADER);
  }

  const emails = sheet.getRange(2, COL.email, lastRow - 1, 1).getValues();
  const existing = sheet.getRange(2, verifyCol, lastRow - 1, 1).getValues();
  let processed = 0;
  let stats = { verified: 0, not_found: 0, catch_all: 0, no_mx: 0, mx_ok: 0, error: 0, skipped: 0 };

  for (let i = 0; i < emails.length; i++) {
    const email = String(emails[i][0] || "").trim().toLowerCase();
    const already = String(existing[i][0] || "").trim();
    if (!email || already) { stats.skipped++; continue; }
    let status, mx = "";
    try {
      if (verifierUrl) {
        const r = _verifyViaEndpoint(verifierUrl, email);
        status = r.status || "error";
        mx = r.mx || "";
      } else {
        const r = _verifyViaDns(email);
        status = r.status;
        mx = r.mx || "";
      }
    } catch (e) {
      status = "error";
      mx = String(e).slice(0, 80);
    }
    sheet.getRange(i + 2, verifyCol).setValue(status + (mx ? " (" + mx + ")" : ""));
    stats[status] = (stats[status] || 0) + 1;
    processed++;
    Utilities.sleep(1000);  // 1 req/s — polite to dns.google and MX servers
  }

  const summary = Object.entries(stats).filter(([, v]) => v > 0).map(([k, v]) => k + ": " + v).join(" · ");
  ui.alert("Verified " + processed + " emails.\n\n" + summary);
}


function _verifyViaDns(email) {
  // dns.google JSON API: https://dns.google/resolve?name=DOMAIN&type=MX
  const at = email.indexOf("@");
  if (at < 0) return { status: "invalid_format" };
  const domain = email.slice(at + 1);
  const url = "https://dns.google/resolve?name=" + encodeURIComponent(domain) + "&type=MX";
  let resp;
  try {
    resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  } catch (e) {
    return { status: "error", mx: String(e).slice(0, 60) };
  }
  if (resp.getResponseCode() !== 200) return { status: "error", mx: "dns http " + resp.getResponseCode() };
  const data = JSON.parse(resp.getContentText());
  const answers = (data.Answer || []).filter(a => a.type === 15);
  if (!answers.length) return { status: "no_mx" };
  // Sort by priority and take the first
  answers.sort(function(a, b) {
    const pa = parseInt(String(a.data).split(" ")[0], 10) || 0;
    const pb = parseInt(String(b.data).split(" ")[0], 10) || 0;
    return pa - pb;
  });
  const mx = String(answers[0].data).split(" ").pop().replace(/\.$/, "");
  return { status: "mx_ok", mx: mx };
}


function _verifyViaEndpoint(url, email) {
  // Calls smtp_verifier.py /verify?email=...
  const u = url.replace(/\/+$/, "") + "/verify?email=" + encodeURIComponent(email);
  const resp = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) return { status: "error", mx: "http " + resp.getResponseCode() };
  const data = JSON.parse(resp.getContentText());
  return { status: data.status || "error", mx: data.mx || "" };
}
