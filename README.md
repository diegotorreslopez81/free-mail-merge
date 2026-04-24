<p align="center">
  <img src="logo.svg" width="128" height="128" alt="Free Mail Merge logo">
</p>

<h1 align="center">Free Mail Merge</h1>

<p align="center">
  A free, unlimited mail merge for Google Sheets + Gmail.<br>
  Drop-in replacement for YAMM. <strong>Built in 15 minutes with Claude, zero lines of code hand-written.</strong>
</p>

![Built with Claude](https://img.shields.io/badge/Built%20with-Claude-DA291C)
![License](https://img.shields.io/badge/license-MIT-blue)
![No dependencies](https://img.shields.io/badge/dependencies-0-green)
![Cost](https://img.shields.io/badge/cost-0%E2%82%AC-brightgreen)

## What this is

A Google Apps Script you paste into your spreadsheet. It turns any Google Sheet into a mail merge tool:

- Pick any Gmail draft as your template (dropdown, like YAMM).
- **Dynamic placeholders**: `{{any_column_header}}` is auto-replaced from the matching cell. No code changes needed.
- HTML signature, PDF attachments, all preserved from your original draft.
- **Settings UI** (no code editing): From name, alias and default test email are entered from a modal.
- Send yourself a real test email before firing the batch.
- Send in batches. Random 5-15 s delay between emails.
- **Gmail quota awareness**: batches stop automatically before hitting the daily limit.
- **Schedule batches** to fire at a specific date/time or every day at a fixed hour.
- **Runs server-side on Google's infra.** Scheduled batches fire with your laptop closed, Chrome killed, you offline. Zero local dependency.
- **Visible schedule tab** (`_Schedule`, read-only): see what's queued, when it fires next, batch size. Auto-refreshed.
- **Track everything per lead** (opt-in, one-time Web App deploy): `opened_at`, `clicked_at` (with URL), `unsubscribed_at`, `bounced_at`, plus `replied_at`.
- **Automatic unsubscribe link** in every email (GDPR-friendly). Unsubscribed addresses land in a `_Suppression` tab and are skipped on future sends.
- Status tracked in the sheet itself: `sent_at`, `sent_status`, `error`, `replied_at`, `opened_at`, `clicked_at`, `unsubscribed_at`, `bounced_at`.
- Resume where it left off if execution is interrupted.
- Use a Gmail alias as the `From:` address.

## Why it exists

YAMM charges €60/year for anything beyond 20 emails/day. I needed to send 450 personalized emails once. I asked Claude to build the equivalent. It took 15 minutes. Sharing it so nobody else pays for something this simple.

## Demo

![Demo GIF](screenshots/demo.gif)

## Install (2 minutes)

1. Open your Google Sheet.
2. **Extensions → Apps Script**.
3. Select all the default code (Cmd+A / Ctrl+A) and paste the contents of [`Code.gs`](./Code.gs).
4. Save (Cmd+S). Give the project any name (we use "Free Mail Merge").
5. Reload the Sheet (Cmd+R). A new menu **"✉️ Free Mail Merge"** appears.
6. First run — do this once, from the menu:
   - **⚙️ Settings** → enter your From name, alias, and default test email.
   - **🗂 Pick leads sheet** → select the tab with your leads.
   - **🎯 Pick template draft** → select your Gmail draft.

No code editing needed. Everything is configured from the menu and stored in the spreadsheet.

## Sheet layout

Your leads tab needs at minimum these columns (in order, starting at A1):

| A | B | C | D | ... | L | M | N | O | P |
|---|---|---|---|---|---|---|---|---|---|
| email | first_name | last_name | company | _any merge cols_ | lk_contacted | sent_at | sent_status | error | replied_at |

- **Columns A-L**: your lead data. Any columns between `company` and `lk_contacted` become merge variables if you reference them as `{{column_name}}` in your draft.
- **Columns M, N, O, P**: the script writes status here. Leave blank (column P is created automatically the first time you run `Check replies`).
- **`lk_contacted`** is optional (you can skip it or repurpose it). If present with value `YES`, the option "Send batch (skip LK-contacted)" skips those rows.

If your columns are different, edit the `COL` object at the top of `Code.gs`.

## Usage

### 1. Create a template draft in Gmail

- Compose a new email in Gmail (don't send it).
- Subject: anything you want. Use placeholders like `{{first_name}}, your opportunity at {{company}}`.
- Body: write normally. Use `{{first_name}}`, `{{company}}`, or any other column name from the sheet wrapped in `{{ }}`.
- Attach your PDF/images if you want.
- Let Gmail add your HTML signature automatically.
- Leave the `To:` field empty.
- If you want to send from a Send-As alias, set the `From:` dropdown in the compose window to the alias.
- Close the compose window → it auto-saves as draft.

![Draft in Gmail](screenshots/draft-gmail.png)

### 2. Pick the template from the Sheet

- In the Sheet, menu **Free Mail Merge → 🎯 Pick template draft**.
- Modal opens with a dropdown of your recent Gmail drafts.
- Pick the one you just created. Click Save.

![Pick template](screenshots/pick-template.png)

### 3. Send a test to yourself

- **Free Mail Merge → 📨 Send test email**.
- Prompt asks for a destination email (defaults to `TEST_EMAIL_DEFAULT`).
- Confirm. A real email is sent to you using the first pending lead's data as the merge sample.

Check:
- `From:` header shows your alias.
- HTML signature is present.
- PDF attachment is included.
- Placeholders are all replaced (no lingering `{{...}}`).
- Links work.

### 4. Fire the batch

- **Free Mail Merge → 📧 Send batch (skip LK-contacted)** (or "Send batch (include LK-contacted)" if you don't use the LK column).
- Prompt asks how many to send. Recommended: 30-40 per run (Apps Script has a 6-min execution limit).
- Hit OK. The script sends in order, 5-15 s between emails, writing `sent_at` and `sent_status` to the sheet row by row.
- When done, it shows a summary alert.

If it hits the 6-minute timeout, just run it again. Already-sent rows (with `sent_at` populated) are skipped automatically.

![Batch progress](screenshots/batch-progress.png)

### 5. Monitor

- **Free Mail Merge → ℹ️ Campaign status**: quick summary (total, sent, errors, pending).
- Open the `sent_at`, `sent_status`, `error`, `replied_at` columns directly to see per-lead status.
- Apps Script **Executions** tab (from the script editor) has per-run logs.

### 6. Schedule batches (optional)

Menu **Free Mail Merge → ⏰ Schedule**:

- **📅 Schedule one-time batch**: prompts for a date/time (`YYYY-MM-DD HH:MM`) and a batch size. Creates an Apps Script trigger that fires once and removes itself.
- **🔁 Schedule daily batch**: prompts for an hour (0-23) and a daily batch size. Fires every day at that hour until you cancel it. Perfect for a warm-up schedule (50/day → 100/day → 200/day).
- **📋 List scheduled jobs**: see what's queued.
- **🗑️ Cancel all scheduled jobs**: remove them all.

**Fully async — laptop can be closed.** Scheduled triggers run on Google's servers, not in your browser. Set daily 9:00 / 80 emails, close the laptop, go to bed. At 9:00 UTC the trigger fires, logs into Gmail on Google's side, sends the batch, writes status back to the Sheet. You can be offline the whole time. This is the killer feature vs any local script or Python cron on your machine.

Caveats:
- Each trigger execution has a 6-minute cap. At 5-15 s random delay, that's ~40-60 emails per run max. For daily sends over 100, schedule two triggers (e.g. 9:00 and 15:00).
- Gmail daily send quotas still apply (Workspace 1,500/day, personal Gmail 100/day).
- The template draft and Sheet tab must exist at trigger time.

Scheduled batches always skip LK-contacted rows and use the currently selected template.

### 7. Track replies and bounces

- Menu **Free Mail Merge → 💬 Check replies**: scans your Gmail inbox from the last 14 days, matches sender addresses against the `email` column, and fills a `replied_at` timestamp in column P for each match.
- Menu **Free Mail Merge → 📬 Check bounces**: scans for `mailer-daemon` / `postmaster` bounce notifications and marks `bounced_at`. Bounced addresses are auto-added to the `_Suppression` tab so you won't re-email them.
- Both are idempotent. Run them manually or add a time-based trigger in the Apps Script editor (e.g., every 30 minutes) for near-real-time updates.

### 8. Track opens, clicks and unsubscribes

This requires deploying the script as an Apps Script Web App, once. Takes 30 seconds.

1. Menu **Free Mail Merge → 🔗 Tracking → 🌐 Setup web app (once)** for the step-by-step modal.
2. In the Apps Script editor: **Deploy → New deployment → Web app**.
3. Description: "Free Mail Merge tracker". Execute as: **Me**. Access: **Anyone**.
4. Click **Deploy**, authorize, copy the URL.
5. Reload the Sheet. Tracking is active from the next send.

Once deployed:
- Every sent email carries an invisible `1x1` pixel that logs `opened_at` on first view.
- Every `<a href>` in the body is wrapped so clicks are logged (`clicked_at`, with the URL) before the recipient is redirected to the real destination.
- Every email ends with a small "Unsubscribe here" link. One click lands the address in `_Suppression` and stamps `unsubscribed_at`. Future batches skip suppressed addresses automatically.
- Menu **🔗 Tracking → 📊 Tracking status** shows whether the web app is deployed, the URL, and current suppression-list size.

Caveats:
- Gmail caches pixel loads via the Google Image Proxy. You'll see the first open; subsequent opens from the same client may be deduplicated.
- Web App URL needs `Access: Anyone` so recipients' email clients can hit it without being logged in to your Google account. Google's console will show you this as a security warning, approve it — the URL is an obscure UUID and only logs hits.

## Configuration reference

Constants at the top of `Code.gs`:

| Constant | Default | What it does |
|---|---|---|
| `SHEET_NAME` | `"YAMM · ..."` | Name of the tab with your leads. |
| `FROM_NAME` | display name | Shows next to `From:` in recipient's inbox. |
| `REPLY_TO` | alias email | Both the `from:` override and `reply-to:`. Must be a Send-As alias in Gmail if different from login address. |
| `BATCH_SIZE_DEFAULT` | 30 | Default batch size offered in the prompt. |
| `DELAY_MIN_MS` / `DELAY_MAX_MS` | 5000 / 15000 | Random delay between sends in ms. |
| `TEST_EMAIL_DEFAULT` | your email | Default destination for the test function. |

## Gmail quotas (not set by this script)

Google enforces daily send limits on your account:

- Workspace account: 1,500 recipients/day via Apps Script.
- Personal `@gmail.com`: 100 recipients/day.

Check your remaining quota in a test run:

```javascript
Logger.log(MailApp.getRemainingDailyQuota());
```

## Warmup warning (read this)

If your sending domain is fresh and has no history, sending 400+ emails on day one will tank your reputation and land future emails in spam. Ramp up: 30-50 day one, 80-100 day two, double from there. Your dominio will thank you.

## FAQ

**Does this work with personal Gmail accounts?**
Yes, but the daily quota is 100 recipients/day.

**Can I customize the delay between sends?**
Edit `DELAY_MIN_MS` and `DELAY_MAX_MS` at the top of the script.

**How do I add more merge variables?**
Any column header in your sheet becomes a merge variable automatically. Just reference it as `{{column_name}}` in the draft. No code changes needed.

**Why does the "From" show my login address instead of the alias?**
The alias must be configured as a "Send mail as" in Gmail settings (Settings → Accounts → Send mail as). Once added, Apps Script can override `from:` with it.

**The menu doesn't appear after reload.**
Wait 10-15 seconds, reload again. Sometimes Apps Script takes a moment to register the `onOpen` trigger on the first load.

## Credits

- Built in 15 minutes with [Claude](https://claude.com) (Opus model) as pair programmer.
- No frameworks, no libraries, no external services. Just Apps Script.
- Inspired by [YAMM](https://yamm.com/) (excellent product, paid).

## License

MIT. See [LICENSE](./LICENSE). Do whatever you want with it.
