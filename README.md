# Free Mail Merge

> A free, unlimited mail merge for Google Sheets + Gmail. Drop-in replacement for YAMM. **Built in 15 minutes with Claude, zero lines of code hand-written.**

![Built with Claude](https://img.shields.io/badge/Built%20with-Claude-DA291C)
![License](https://img.shields.io/badge/license-MIT-blue)
![No dependencies](https://img.shields.io/badge/dependencies-0-green)
![Cost](https://img.shields.io/badge/cost-0%E2%82%AC-brightgreen)

## What this is

A Google Apps Script you paste into your spreadsheet. It turns any Google Sheet into a mail merge tool:

- Pick any Gmail draft as your template (dropdown, like YAMM).
- Placeholders `{{first_name}}`, `{{company}}`, `{{anything}}` replaced per row.
- HTML signature, PDF attachments, all preserved from your original draft.
- Send yourself a real test email before firing the batch.
- Send in batches. Random 5-15 s delay between emails.
- **Schedule batches** to fire at a specific date/time or every day at a fixed hour.
- **Runs server-side on Google's infra.** Scheduled batches fire with your laptop closed, Chrome killed, you offline. Zero local dependency.
- **Track replies** automatically: one click scans your inbox and fills a `replied_at` column.
- Status tracked in the sheet itself: `sent_at`, `sent_status`, `error`, `replied_at`.
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
4. Edit these four constants at the top to match your setup:

```javascript
const FROM_NAME = "Your Name · Your Company";                // display name
const REPLY_TO = "you@yourdomain.com";                       // reply-to + from alias
const TEST_EMAIL_DEFAULT = "you@yourdomain.com";             // where test mails go
```

(The leads tab itself is picked from the menu, not hardcoded, so you can rename the tab anytime.)

5. Save (Cmd+S). Give the project any name (we use "Free Mail Merge").
6. Reload the Sheet (Cmd+R). A new menu **"Free Mail Merge"** appears.
7. First run: menu **Free Mail Merge → 🗂 Pick leads sheet** and select the tab with your leads.

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

### 7. Track replies

Menu **Free Mail Merge → 💬 Check replies**: scans your Gmail inbox from the last 14 days, matches sender addresses against the `email` column, and fills a `replied_at` timestamp in column P for each match. Skips rows where `replied_at` is already set. Run it manually whenever you want, or add a time-based trigger in the Apps Script editor (e.g., every 30 minutes) for near-real-time reply tracking.

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
Any column header in your sheet becomes a merge variable if you reference it as `{{column_name}}` in the draft. Note: the current code hard-codes `first_name`, `company`, `sector_huma`. Generalise by replacing the `_mergeTemplate` function with a loop over column headers.

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
