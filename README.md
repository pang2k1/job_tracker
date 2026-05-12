# job_tracker

**Summarized using Claude, so blame the AI if this README does not make sense**

A lightweight Google Apps Script that automatically tracks your job applications using Gmail and Google Sheets — no third-party tools, no paid subscriptions.

## How It Works

- **Scans your Gmail daily** for application confirmation emails and adds them as rows in a Google Sheet
- **Auto-detects rejections** from reply keywords (e.g. "unfortunately", "unable to move forward") or silence after 30 days
- **Sends you a daily email digest** at 18:00 Bangkok time (configurable) listing everything still under review
- **Never re-adds closed applications** — thread IDs are stored permanently so deleted or closed rows don't come back

---

## Files

```
JobTracker/
├── JobTracker.gs   # Main logic — Gmail sync, rejection detection, email digest
├── config.gs       # One-time setup — fill in your values and run once
└── README.md
```

---

## Google Sheet Structure

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| Company Name | Position | Date Applied | Days Pending *(auto)* | Status *(you fill)* | Thread ID |

- **Column D** — updated automatically every day
- **Column E** — leave blank to keep active. Type anything (`rejected`, `interview`, `withdrew`) to close it
- **Column F** — Stores Gmail thread IDs to prevent duplicate entries

---

## Setup

### 1. Create a Google Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet. Note the sheet tab name (default is `Sheet1` — you can rename it to anything, e.g. `Applications`).

### 2. Open Apps Script

In your spreadsheet, go to **Extensions → Apps Script**.

### 3. Add the files

In the Apps Script editor:
1. Rename the default `Code.gs` file to `JobTracker.gs` and paste the contents of `JobTracker.gs`
2. Click **+** next to Files → **Script** → name it `config` → paste the contents of `config.gs`

### 4. Set your config values

In `config.gs`, fill in your real values:

```javascript
function setScriptProperties() {
  PropertiesService.getScriptProperties().setProperties({
    'EMAIL':     'you@gmail.com',
    'GSHEET':    'Applications',   // must match your sheet tab name exactly
    'BEGINDATE': '2026/01/01'      // earliest date to scan Gmail from
  });
}
```

Run `setScriptProperties` by selecting it in the function dropdown and clicking **Run**. Approve the permissions prompt.

After running, **clear the values back to placeholders** before committing to GitHub.

### 5. Run setup

Select `setup` in the function dropdown and click **Run**. This will:
- Create the sheet with the correct columns and formatting
- Set up a daily trigger at 11:00 UTC (18:00 Bangkok / GMT+7)

### 6. Test it

Select `runDailyJobTracker` and click **Run**. Check your inbox — you should receive a digest email, and your sheet should be populated with any matching emails found since your `BEGINDATE`.

---

## Daily Behaviour

Each day at 18:00 (GMT+7), the script runs in this order:

1. **Gmail sync** — searches for new application confirmation emails, adds untracked ones as yellow-highlighted rows
2. **Auto-reject detection** — checks each active row for rejection signals (see rules below)
3. **Update days pending** — refreshes column D for all active rows
4. **Send digest email** — lists all rows where Status is blank

---

## Auto-Rejection Rules

| Rule | Trigger | Status set to |
|---|---|---|
| Keyword detection | A reply in the same Gmail thread contains rejection language | `rejected` |
| Silent rejection | No update after 30 days | `rejected (silent — 30d)` |

Rejected rows are highlighted **light red** in the sheet.

**Rejection keywords detected:**
`unfortunately` · `unable to move forward` · `not moving forward` · `not selected` · `move on with other candidates` · `position has been filled` · `selected another candidate` · `regret to inform` · `not a match` · and more

> **Note:** Keyword detection only works if the rejection email lands in the **same Gmail thread** as the original confirmation. If a company sends a rejection from a separate thread, it falls back to the 30-day silent rule.

---

## Managing Applications

| Action | What to do |
|---|---|
| New application | Nothing — Gmail sync adds it automatically |
| Auto-added row looks wrong | Edit company/position directly in the sheet (yellow = auto-detected, verify it) |
| Got rejected / accepted / withdrew | Type anything in column E — script will skip it from then on |
| Want to delete a row | Don't — type a status instead. Deleting loses the Thread ID and the application may be re-added from Gmail |

---

## Changing the Trigger Time

The trigger is set in `createDailyTrigger()` in `JobTracker.gs`:

```javascript
.atHour(11) // 11:00 UTC = 18:00 GMT+7
```

Change the hour to match your timezone. Apps Script triggers use UTC.

| Your timezone | UTC offset | atHour value for 18:00 local |
|---|---|---|
| GMT+7 (Bangkok) | +7 | 11 |
| GMT+8 (Singapore) | +8 | 10 |
| GMT+1 (Berlin) | +1 | 17 |
| GMT+0 (London) | +0 | 18 |

Re-run `createDailyTrigger()` after changing the value.

---

## Permissions Required

When you first run the script, Google will ask you to approve:
- **Gmail** — to read and search your emails
- **Google Sheets** — to read and write your spreadsheet
- **External services** — to send emails via `GmailApp.sendEmail`

All data stays within your own Google account. No external servers involved.

---

## Limitations

- Confirmation email parsing (company name, position) is best-effort from subject line and sender. Yellow-highlighted rows should be verified manually.
- Rejection detection via keywords only works for replies within the same Gmail thread.
- Gmail search is capped at 50 threads per sync — sufficient for most job searches.
