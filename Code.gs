/**
 * Zoom Webinar → Google Calendar Automation
 *
 * Watches Gmail for Zoom webinar registration confirmations and
 * automatically creates Google Calendar events with reminders.
 *
 * Setup:
 *   1. Go to script.google.com → New project
 *   2. Paste this file, replacing all content
 *   3. Update CONFIG.keywords to match your webinar programme
 *   4. Run createProcessedLabel() → Run installTrigger()
 *   5. Optionally run backfillHistoricalWebinars() once
 *
 * License: MIT
 */

// ============================================================
// CONFIG — edit this section to match your webinar programme
// ============================================================
var CONFIG = {
  // Gmail search keywords — emails must match at least one
  keywords: [
    'ihpba',
    'eahpba',
    'hpb grand rounds',
    'hpb cases',
    'hepatobiliary',
    'ecg webinar',
    'ecg roadshow'
  ],

  // Gmail label applied to processed emails (prevents duplicates)
  processedLabel: 'Webinar-Calendar-Added',

  // Calendar event colour (Pale Blue = 1, Banana = 5, Sage = 2, etc.)
  eventColor: CalendarApp.EventColor.PALE_BLUE,

  // Reminders
  emailReminderMinutes: 1440,  // 1 day before
  popupReminderMinutes: 60     // 1 hour before
};

// ============================================================
// MAIN — runs on an hourly trigger
// ============================================================
function checkForNewWebinars() {
  var query = buildQuery_();
  var threads = GmailApp.search(query, 0, 20);
  var label = getOrCreateLabel_(CONFIG.processedLabel);

  Logger.log('Query: ' + query);
  Logger.log('Found ' + threads.length + ' unprocessed thread(s).');

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      processMessage_(messages[j]);
    }
    threads[i].addLabel(label);
  }
}

// ============================================================
// BACKFILL — run once to catch historical emails
// ============================================================
function backfillHistoricalWebinars() {
  var query = buildBackfillQuery_();
  var threads = GmailApp.search(query, 0, 100);
  var label = getOrCreateLabel_(CONFIG.processedLabel);

  Logger.log('Backfill query: ' + query);
  Logger.log('Found ' + threads.length + ' historical thread(s).');

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      processMessage_(messages[j]);
    }
    threads[i].addLabel(label);
  }
}

// ============================================================
// SETUP HELPERS — run these manually once
// ============================================================

/** Creates the Gmail label used for deduplication. Run this first. */
function createProcessedLabel() {
  getOrCreateLabel_(CONFIG.processedLabel);
  Logger.log('Label created: ' + CONFIG.processedLabel);
}

/** Installs the hourly trigger. Run this after createProcessedLabel. */
function installTrigger() {
  // Remove any existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkForNewWebinars') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('checkForNewWebinars')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly trigger installed for checkForNewWebinars().');
}

// ============================================================
// INTERNAL — parsing and calendar logic
// ============================================================

function processMessage_(message) {
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // Only process Zoom confirmation emails
  if (subject.indexOf('Confirmation') === -1) return;

  var title = subject.replace(/\s*Confirmation\s*$/i, '').trim();
  Logger.log('Processing: "' + subject + '"');

  // Parse date/time from Zoom's format:
  //   Date & Time  Oct 23, 2025 06:00 PM Amsterdam, Berlin, Rome, Stockholm, Vienna
  var dateTimeInfo = parseDateTimeFromBody_(body);

  if (!dateTimeInfo) {
    Logger.log('Could not parse date/time from: "' + subject + '" — skipping.');
    return;
  }

  // Check for duplicate on calendar
  var cal = CalendarApp.getDefaultCalendar();
  var searchStart = new Date(dateTimeInfo.start.getTime() - 3600000); // 1hr before
  var searchEnd = new Date(dateTimeInfo.start.getTime() + 7200000);   // 2hr after
  var existing = cal.getEvents(searchStart, searchEnd);

  for (var k = 0; k < existing.length; k++) {
    if (existing[k].getTitle().indexOf(title) !== -1) {
      Logger.log('Already on calendar: "' + title + '"');
      return;
    }
  }

  // Create the event
  var event = cal.createEvent(title, dateTimeInfo.start, dateTimeInfo.end);
  event.setColor(CONFIG.eventColor);
  event.addEmailReminder(CONFIG.emailReminderMinutes);
  event.addPopupReminder(CONFIG.popupReminderMinutes);
  event.setDescription('Auto-added from Zoom registration confirmation.\nOriginal email: ' + message.getDate());

  Logger.log('Created calendar event: "' + title + '" on ' + dateTimeInfo.start);
}

function parseDateTimeFromBody_(body) {
  // Match Zoom's "Date & Time" line
  // Example: "Date & Time  Oct 23, 2025 06:00 PM Amsterdam, Berlin, Rome, Stockholm, Vienna"
  var pattern = /Date\s*&\s*Time\s+([A-Z][a-z]{2}\s+\d{1,2},?\s+\d{4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s+(.+)/i;
  var match = body.match(pattern);

  if (!match) {
    // Try alternative format: "Date  Time  Oct 23, 2025 06:00 PM ..."
    pattern = /Date\s+Time\s+([A-Z][a-z]{2}\s+\d{1,2},?\s+\d{4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s+(.+)/i;
    match = body.match(pattern);
  }

  if (!match) return null;

  var dateStr = match[1].trim();
  var timeStr = match[2].trim();
  var tzCities = match[3].trim();

  // Map city list to IANA timezone
  var tz = mapCitiesToTimezone_(tzCities);

  // Parse the date
  var dateTimeStr = dateStr + ' ' + timeStr;
  var parsed = new Date(dateTimeStr);

  if (isNaN(parsed.getTime())) return null;

  // Adjust for timezone if we can detect it
  if (tz) {
    // Use Utilities.formatDate to convert properly
    var formatted = Utilities.formatDate(parsed, tz, "yyyy-MM-dd'T'HH:mm:ss");
    parsed = new Date(formatted);
  }

  var endTime = new Date(parsed.getTime() + 3600000); // default 1hr duration

  return { start: parsed, end: endTime };
}

function mapCitiesToTimezone_(cityStr) {
  var lower = cityStr.toLowerCase();

  var mappings = [
    { cities: ['amsterdam', 'berlin', 'rome', 'stockholm', 'vienna', 'paris', 'brussels', 'madrid'], tz: 'Europe/Berlin' },
    { cities: ['london', 'lisbon', 'dublin', 'edinburgh'], tz: 'Europe/London' },
    { cities: ['new york', 'toronto', 'eastern'], tz: 'America/New_York' },
    { cities: ['chicago', 'central', 'mexico city'], tz: 'America/Chicago' },
    { cities: ['denver', 'mountain'], tz: 'America/Denver' },
    { cities: ['los angeles', 'pacific', 'san francisco', 'vancouver'], tz: 'America/Los_Angeles' },
    { cities: ['tokyo', 'seoul', 'osaka'], tz: 'Asia/Tokyo' },
    { cities: ['sydney', 'melbourne', 'canberra'], tz: 'Australia/Sydney' },
    { cities: ['hong kong'], tz: 'Asia/Hong_Kong' },
    { cities: ['singapore', 'kuala lumpur'], tz: 'Asia/Singapore' },
    { cities: ['mumbai', 'kolkata', 'new delhi', 'chennai'], tz: 'Asia/Kolkata' },
    { cities: ['dubai', 'abu dhabi', 'muscat'], tz: 'Asia/Dubai' },
    { cities: ['moscow', 'st. petersburg'], tz: 'Europe/Moscow' },
    { cities: ['athens', 'bucharest', 'helsinki', 'istanbul', 'cairo'], tz: 'Europe/Athens' },
    { cities: ['beijing', 'shanghai'], tz: 'Asia/Shanghai' },
    { cities: ['sao paulo', 'brasilia'], tz: 'America/Sao_Paulo' }
  ];

  for (var i = 0; i < mappings.length; i++) {
    for (var j = 0; j < mappings[i].cities.length; j++) {
      if (lower.indexOf(mappings[i].cities[j]) !== -1) {
        return mappings[i].tz;
      }
    }
  }

  // Default to CET if unknown (most medical webinars are European-timed)
  return 'Europe/Berlin';
}

// ============================================================
// QUERY BUILDERS
// ============================================================

function buildQuery_() {
  var keywordParts = CONFIG.keywords.map(function(kw) {
    return 'subject:"' + kw + '"';
  });
  return 'from:no-reply@zoom.us subject:confirmation (' + keywordParts.join(' OR ') + ') -label:' + CONFIG.processedLabel;
}

function buildBackfillQuery_() {
  var keywordParts = CONFIG.keywords.map(function(kw) {
    return 'subject:"' + kw + '"';
  });
  return 'from:no-reply@zoom.us subject:confirmation (' + keywordParts.join(' OR ') + ')';
}

// ============================================================
// UTILITY
// ============================================================

function getOrCreateLabel_(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}
