# Work Time Management Tracker — Google Apps Script

This Apps Script syncs three sources into the `Inbox` tab of your tracker:

1. Google Calendar meetings
2. Gmail inbox messages
3. Mattermost channel messages

It runs automatically every 15 minutes by default.

## Inbox structure expected

The script writes these columns:

| Column | Value |
|---|---|
| A | Added |
| B | Raw task / commitment |
| C | Source |
| D | Area |
| E | Notes / context |
| F | Processed? |

It also creates a hidden `_SyncLog` sheet to prevent duplicate imports.

## Installation

1. Open `Work_time_management_tracker` in Google Sheets.
2. Go to **Extensions → Apps Script**.
3. Delete the starter code in `Code.gs`.
4. Paste the contents of `Code.gs` from this package.
5. Save.
6. In Apps Script, run `setup()` manually once.
7. Approve permissions for:
   - Google Sheets
   - Gmail
   - Google Calendar
   - external requests (for Mattermost)

## Mattermost setup

You need:

- your Mattermost base URL, e.g. `https://chat.example.org`
- a Personal Access Token
- one or more Channel IDs

### Recommended way

Open **Project Settings → Script Properties** and add:

- `MATTERMOST_BASE_URL`
- `MATTERMOST_TOKEN`
- `MATTERMOST_CHANNEL_IDS`

For multiple channel IDs, use commas:

`channelA,channelB,channelC`

The script deliberately does **not** store the token in the spreadsheet cells.

## Scheduler

`setup()` creates an installable trigger for `syncAll()` every 15 minutes.

To stop automatic syncing, run:

`stopScheduler()`

To change frequency, edit:

`SYNC_EVERY_MINUTES: 15`

Google Apps Script supports standard installable trigger intervals such as 5, 10, 15, and 30 minutes.

## What gets imported

### Calendar
Upcoming events for the next 2 days. New meetings become Inbox items with:
- meeting title
- start/end time
- guests
- location
- description

### Gmail
New messages received in the inbox since the previous sync. The script:
- does not mark messages read
- does not archive/delete anything
- stores subject/from/body preview

### Mattermost
New posts from the configured channels since the previous sync.
- your own posts are ignored
- deleted posts are ignored
- channel message text is added to Inbox

## Suggested workflow

`Calendar / Gmail / Mattermost → Inbox → Tasks → Planned Date → Today`

The script only captures incoming work. It intentionally does not automatically decide your final priority or schedule.