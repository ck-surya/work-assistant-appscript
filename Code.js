/**
 * Work Time Management Tracker Sync
 *
 * Sources:
 * 1) Google Calendar meetings
 * 2) Gmail inbox messages
 * 3) Mattermost channel messages
 *
 * Destination:
 * Sheet tab "Inbox"
 * Columns:
 * A Added
 * B Raw task / commitment
 * C Source
 * D Area
 * E Notes / context
 * F Processed?
 *
 * A hidden "_SyncLog" tab is used for de-duplication.
 */

const CONFIG = {
  INBOX_SHEET: 'Inbox',
  SYNC_LOG_SHEET: '_SyncLog',

  // Google Calendar
  CALENDAR_ID: 'primary',
  CALENDAR_LOOKAHEAD_DAYS: 2,

  // Gmail
  GMAIL_MAX_THREADS: 50,

  // Trigger interval
  SYNC_EVERY_MINUTES: 15,

  // Mattermost Script Properties:
  // MATTERMOST_BASE_URL   e.g. https://chat.example.org
  // MATTERMOST_TOKEN      Personal Access Token
  // MATTERMOST_CHANNEL_IDS comma-separated channel IDs
};

/**
 * Run once manually after pasting this script.
 * Creates the Inbox sheet if missing, the hidden sync log,
 * and an installable time trigger.
 */
function setup() {
  ensureSheets_();
  createSyncTrigger_();

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LAST_GMAIL_SYNC')) {
    props.setProperty('LAST_GMAIL_SYNC', String(Date.now() - 60 * 60 * 1000));
  }
  if (!props.getProperty('LAST_MATTERMOST_SYNC')) {
    props.setProperty('LAST_MATTERMOST_SYNC', String(Date.now() - 60 * 60 * 1000));
  }

  SpreadsheetApp.getActive().toast(
    'Setup complete. Configure Mattermost Script Properties, then run syncAll().',
    'Work Tracker Sync',
    8
  );
}

/**
 * Main scheduler entry point.
 */
function syncAll() {
  ensureSheets_();

  const results = {
    calendar: safeRun_('Calendar', syncCalendar_),
    gmail: safeRun_('Gmail', syncGmail_),
    mattermost: safeRun_('Mattermost', syncMattermost_)
  };

  Logger.log(JSON.stringify(results));
}

/**
 * Fetch upcoming Calendar events and add new meetings to Inbox.
 */
function syncCalendar_() {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) throw new Error('Calendar not found: ' + CONFIG.CALENDAR_ID);

  const start = new Date();
  const end = new Date(start.getTime() + CONFIG.CALENDAR_LOOKAHEAD_DAYS * 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(start, end);

  let added = 0;

  events.forEach(event => {
    // Skip declined/cancelled style entries when possible.
    const title = event.getTitle() || '(Untitled meeting)';
    const eventStart = event.getStartTime();
    const eventEnd = event.getEndTime();

    const key = [
      'calendar',
      event.getId(),
      eventStart.getTime(),
      eventEnd.getTime(),
      title
    ].join('|');

    if (isSeen_(key)) return;

    const guests = event.getGuestList()
      .map(g => g.getEmail())
      .filter(Boolean)
      .join(', ');

    const notes = [
      'Calendar: ' + CONFIG.CALENDAR_ID,
      'Start: ' + formatDateTime_(eventStart),
      'End: ' + formatDateTime_(eventEnd),
      guests ? 'Guests: ' + guests : '',
      event.getLocation() ? 'Location: ' + event.getLocation() : '',
      event.getDescription() ? 'Description: ' + trim_(event.getDescription(), 600) : ''
    ].filter(Boolean).join('\n');

    appendInbox_(
      eventStart,
      'Meeting: ' + title,
      'Meeting',
      'Stewardship',
      notes
    );

    markSeen_(key, 'Calendar');
    added++;
  });

  return { added, scanned: events.length };
}

/**
 * Fetch new Gmail inbox messages since last sync.
 * This captures messages as Inbox items; it does not mark emails read.
 */
function syncGmail_() {
  const props = PropertiesService.getScriptProperties();
  const lastSync = Number(props.getProperty('LAST_GMAIL_SYNC') || (Date.now() - 60 * 60 * 1000));
  const now = Date.now();

  // Search a slightly wider window, then filter by message date.
  const threads = GmailApp.search('in:inbox newer_than:2d', 0, CONFIG.GMAIL_MAX_THREADS);

  let scanned = 0;
  let added = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      scanned++;
      const messageTime = message.getDate().getTime();
      if (messageTime <= lastSync) return;

      const key = 'gmail|' + message.getId();
      if (isSeen_(key)) return;

      const subject = message.getSubject() || '(No subject)';
      const from = message.getFrom() || '';
      const body = trim_(message.getPlainBody() || '', 700);

      const notes = [
        from ? 'From: ' + from : '',
        'Received: ' + formatDateTime_(message.getDate()),
        body ? 'Preview: ' + body : ''
      ].filter(Boolean).join('\n');

      appendInbox_(
        message.getDate(),
        'Email: ' + subject,
        'Email',
        inferArea_(subject + '\n' + body),
        notes
      );

      markSeen_(key, 'Gmail');
      added++;
    });
  });

  props.setProperty('LAST_GMAIL_SYNC', String(now));
  return { added, scanned };
}

/**
 * Fetch Mattermost posts from configured channels since last sync.
 *
 * Required Script Properties:
 * MATTERMOST_BASE_URL
 * MATTERMOST_TOKEN
 * MATTERMOST_CHANNEL_IDS
 */
function syncMattermost_() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = (props.getProperty('MATTERMOST_BASE_URL') || '').replace(/\/+$/, '');
  const token = props.getProperty('MATTERMOST_TOKEN') || '';
  const channelIds = (props.getProperty('MATTERMOST_CHANNEL_IDS') || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  if (!baseUrl || !token || channelIds.length === 0) {
    return {
      skipped: true,
      reason: 'Set MATTERMOST_BASE_URL, MATTERMOST_TOKEN and MATTERMOST_CHANNEL_IDS in Script Properties.'
    };
  }

  const lastSync = Number(props.getProperty('LAST_MATTERMOST_SYNC') || (Date.now() - 60 * 60 * 1000));
  const now = Date.now();

  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/json'
  };

  // Resolve current user once, so we can optionally ignore our own posts.
  const me = fetchJson_(baseUrl + '/api/v4/users/me', headers);
  const myUserId = me.id;

  let scanned = 0;
  let added = 0;

  channelIds.forEach(channelId => {
    const url = baseUrl + '/api/v4/channels/' + encodeURIComponent(channelId) +
      '/posts?since=' + encodeURIComponent(lastSync);

    const data = fetchJson_(url, headers);
    const order = data.order || [];
    const posts = data.posts || {};

    order.slice().reverse().forEach(postId => {
      const post = posts[postId];
      if (!post) return;

      scanned++;

      // Ignore deleted posts and our own posts.
      if (post.delete_at && post.delete_at > 0) return;
      if (post.user_id === myUserId) return;

      const key = 'mattermost|' + post.id;
      if (isSeen_(key)) return;

      const message = trim_(post.message || '', 900);
      if (!message) return;

      const createdAt = new Date(post.create_at);
      const notes = [
        'Channel ID: ' + channelId,
        'Post ID: ' + post.id,
        'Created: ' + formatDateTime_(createdAt),
        post.root_id ? 'Thread root: ' + post.root_id : ''
      ].filter(Boolean).join('\n');

      appendInbox_(
        createdAt,
        'Mattermost: ' + message,
        'Mattermost',
        inferArea_(message),
        notes
      );

      markSeen_(key, 'Mattermost');
      added++;
    });
  });

  props.setProperty('LAST_MATTERMOST_SYNC', String(now));
  return { added, scanned };
}

/**
 * Create / replace the scheduler trigger.
 */
function createSyncTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncAll')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .everyMinutes(CONFIG.SYNC_EVERY_MINUTES)
    .create();
}

/**
 * Delete scheduler trigger if you want to stop automatic syncing.
 */
function stopScheduler() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncAll')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * Helper to update Mattermost connection settings without opening Project Settings.
 * Run once, edit values below, then delete the token from this function afterwards.
 */
function setMattermostConfig() {
  PropertiesService.getScriptProperties().setProperties({
    MATTERMOST_BASE_URL: 'https://YOUR-MATTERMOST-SERVER',
    MATTERMOST_TOKEN: 'PASTE_PERSONAL_ACCESS_TOKEN_HERE',
    MATTERMOST_CHANNEL_IDS: 'CHANNEL_ID_1,CHANNEL_ID_2'
  });

  SpreadsheetApp.getActive().toast('Mattermost Script Properties saved.', 'Work Tracker Sync', 5);
}

/* ------------------------- Internal helpers ------------------------- */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inbox = ss.getSheetByName(CONFIG.INBOX_SHEET);
  if (!inbox) {
    inbox = ss.insertSheet(CONFIG.INBOX_SHEET);
    inbox.getRange(1, 1, 1, 6).setValues([[
      'Added',
      'Raw task / commitment',
      'Source',
      'Area',
      'Notes / context',
      'Processed?'
    ]]);
    inbox.setFrozenRows(1);
  }

  let log = ss.getSheetByName(CONFIG.SYNC_LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(CONFIG.SYNC_LOG_SHEET);
    log.getRange(1, 1, 1, 3).setValues([['Key', 'Source', 'Seen At']]);
    log.setFrozenRows(1);
    log.hideSheet();
  }
}

function appendInbox_(added, task, source, area, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.INBOX_SHEET);
  sheet.appendRow([
    added || new Date(),
    task || '',
    source || '',
    area || 'Admin',
    notes || '',
    'No'
  ]);
}

function isSeen_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYNC_LOG_SHEET);
  if (sheet.getLastRow() < 2) return false;

  const finder = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(key)
    .matchEntireCell(true)
    .findNext();

  return !!finder;
}

function markSeen_(key, source) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYNC_LOG_SHEET);
  sheet.appendRow([key, source, new Date()]);
}

function safeRun_(name, fn) {
  try {
    return fn();
  } catch (err) {
    console.error(name + ' sync failed:', err);
    return { error: String(err && err.message ? err.message : err) };
  }
}

function fetchJson_(url, headers) {
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' for ' + url + ': ' + trim_(text, 500));
  }

  return JSON.parse(text);
}

function inferArea_(text) {
  const s = String(text || '').toLowerCase();

  if (/(policy|steward|council|decision|insurance|member|governance|meeting follow.?up)/.test(s)) {
    return 'Stewardship';
  }
  if (/(bug|api|code|deploy|test|frontend|backend|css|database|github|pull request|pr\b)/.test(s)) {
    return 'Development';
  }
  if (/(learn|study|reading|dsa|course|practice|contest|algorithm|tutorial)/.test(s)) {
    return 'Learning';
  }
  return 'Admin';
}

function formatDateTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd-MMM-yyyy HH:mm'
  );
}

function trim_(value, maxLen) {
  const s = String(value || '').replace(/\s+/g, ' ').trim();
  return s.length <= maxLen ? s : s.slice(0, maxLen - 1) + '…';
}