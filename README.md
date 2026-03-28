# Zoom Webinar → Google Calendar Automation

A Google Apps Script that watches your Gmail for Zoom webinar registration confirmation emails and automatically creates Google Calendar events with reminders.

## Features

- **Hourly scanning** — runs via a time-based trigger, no manual intervention
- **Reliable parsing** — handles Zoom's `Date & Time` format with city-based timezones
- **Deduplication** — checks calendar for existing events + labels processed emails
- **Backfill** — one-time function to process all historical confirmation emails
- **Easy to customize** — single `CONFIG` block at the top of the script

## Use Case

Many professional societies (medical, legal, engineering) require a minimum attendance rate at webinar series to earn certification or CPD credits. This script ensures every Zoom registration automatically appears on your calendar with reminders, so you never miss a session.

## Setup

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Delete the default code and paste the contents of `Code.gs`
3. Edit the `CONFIG` block at the top:
   - Update `keywords` to match your webinar programme (e.g., subject line keywords)
   - Adjust `processedLabel`, event colour, and reminder times if needed
4. In the function dropdown, select `createProcessedLabel` → click **Run**
   - Grant the permissions Google asks for (Gmail read + Calendar write)
5. Select `installTrigger` → click **Run**
6. *(Optional)* Select `backfillHistoricalWebinars` → click **Run** to catch past emails

The script now runs every hour in the background.

## Configuration

```javascript
var CONFIG = {
  keywords: [         // Emails must match at least one keyword in subject
    'ihpba',
    'eahpba',
    'hpb grand rounds',
    // Add your own keywords here
  ],
  processedLabel: 'Webinar-Calendar-Added',  // Gmail label for processed emails
  eventColor: CalendarApp.EventColor.PALE_BLUE,
  emailReminderMinutes: 1440,  // 1 day before
  popupReminderMinutes: 60     // 1 hour before
};
```

## How It Works

1. Searches Gmail for unprocessed Zoom confirmation emails matching your keywords
2. Parses the `Date & Time` field from the email body
3. Maps Zoom's city-based timezone strings (e.g., "Amsterdam, Berlin, Rome") to IANA timezones
4. Checks the calendar for duplicates before creating the event
5. Labels the email so it's never processed again

## Supported Timezones

The script maps common Zoom city strings to IANA timezones:

| Cities | Timezone |
|--------|----------|
| Amsterdam, Berlin, Rome, Stockholm, Vienna | Europe/Berlin |
| London, Dublin, Edinburgh | Europe/London |
| New York, Toronto | America/New_York |
| Tokyo, Seoul, Osaka | Asia/Tokyo |
| Sydney, Melbourne | Australia/Sydney |
| And 10+ more... | See `mapCitiesToTimezone_()` |

## Limitations

- Only processes **Zoom** confirmation emails (from `no-reply@zoom.us`)
- Does not register you for webinars — you still need to register manually
- Does not track actual attendance — only registration
- The "unsafe app" warning during setup is normal for personal Apps Scripts

## Blog Post

For a detailed walkthrough, see the [blog post](https://chrislongros.com/?p=10373).

## License

MIT
