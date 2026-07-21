const LEAD_HEADERS = [
  'Timestamp',
  'Lead ID',
  'Name',
  'Email',
  'Company',
  'Focus',
  'Message',
  'Source Page',
  'Referrer',
  'UTM Source',
  'UTM Medium',
  'UTM Campaign',
  'UTM Content',
  'UTM Term',
  'Status',
  'Appointment Date',
  'Calendar Event ID',
  'Last Updated'
];

function doGet() {
  return jsonResponse_({ ok: true, service: 'AKRD lead intake' });
}

function doPost(event) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = parsePayload_(event);

    // Honeypot submissions receive a neutral response and are not stored.
    if (String(payload.website || '').trim()) {
      return jsonResponse_({ ok: true });
    }

    const name = cleanText_(payload.name, 120);
    const email = cleanText_(payload.email, 180).toLowerCase();
    const company = cleanText_(payload.company, 180);
    const focus = cleanText_(payload.focus, 180);
    const message = cleanText_(payload.message, 4000);

    if (!name || !isEmail_(email) || !message) {
      return jsonResponse_({ ok: false, error: 'invalid_submission' });
    }

    if (isRateLimited_(email)) {
      return jsonResponse_({ ok: false, error: 'rate_limited' });
    }

    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME') || 'Website Leads';
    const notificationEmail = properties.getProperty('NOTIFICATION_EMAIL');
    if (!spreadsheetId) throw new Error('Missing SPREADSHEET_ID script property.');

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    ensureHeaders_(sheet);

    const now = new Date();
    const leadId = createLeadId_(now);
    const row = [
      now,
      leadId,
      safeCell_(name),
      safeCell_(email),
      safeCell_(company),
      safeCell_(focus),
      safeCell_(message),
      safeCell_(cleanText_(payload.sourcePage, 500)),
      safeCell_(cleanText_(payload.referrer, 500)),
      safeCell_(cleanText_(payload.utmSource, 180)),
      safeCell_(cleanText_(payload.utmMedium, 180)),
      safeCell_(cleanText_(payload.utmCampaign, 180)),
      safeCell_(cleanText_(payload.utmContent, 180)),
      safeCell_(cleanText_(payload.utmTerm, 180)),
      'New',
      '',
      '',
      now
    ];
    sheet.appendRow(row);

    if (notificationEmail) {
      sendOwnerNotification_({
        to: notificationEmail,
        leadId,
        name,
        email,
        company,
        focus,
        message,
        sourcePage: cleanText_(payload.sourcePage, 500)
      });
    }

    return jsonResponse_({ ok: true, leadId: leadId });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ ok: false, error: 'server_error' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function syncCalendarBookings() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const sheetName = properties.getProperty('SHEET_NAME') || 'Website Leads';
  const calendarId = properties.getProperty('CALENDAR_ID');
  if (!spreadsheetId) throw new Error('Missing SPREADSHEET_ID script property.');

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  ensureHeaders_(sheet);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), LEAD_HEADERS.length).getValues();
  const header = values[0];
  const emailColumn = header.indexOf('Email');
  const statusColumn = header.indexOf('Status');
  const appointmentColumn = header.indexOf('Appointment Date');
  const eventIdColumn = header.indexOf('Calendar Event ID');
  const updatedColumn = header.indexOf('Last Updated');
  if ([emailColumn, statusColumn, appointmentColumn, eventIdColumn, updatedColumn].some(index => index < 0)) return;

  const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  if (!calendar) throw new Error('Calendar could not be accessed.');

  const start = new Date();
  start.setDate(start.getDate() - 30);
  const end = new Date();
  end.setDate(end.getDate() + 120);
  const events = calendar.getEvents(start, end);
  const bookingByEmail = {};

  events.forEach(event => {
    event.getGuestList().forEach(guest => {
      const guestEmail = String(guest.getEmail() || '').toLowerCase();
      if (!guestEmail) return;
      const existing = bookingByEmail[guestEmail];
      if (!existing || event.getStartTime() > existing.getStartTime()) bookingByEmail[guestEmail] = event;
    });
  });

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const email = String(values[rowIndex][emailColumn] || '').toLowerCase();
    const event = bookingByEmail[email];
    if (!event) continue;
    values[rowIndex][statusColumn] = 'Appointment booked';
    values[rowIndex][appointmentColumn] = event.getStartTime();
    values[rowIndex][eventIdColumn] = event.getId();
    values[rowIndex][updatedColumn] = new Date();
  }

  sheet.getRange(2, 1, values.length - 1, LEAD_HEADERS.length).setValues(values.slice(1));
}

function installBookingSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'syncCalendarBookings')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('syncCalendarBookings').timeBased().everyHours(1).create();
}

function parsePayload_(event) {
  if (!event) return {};
  const type = String(event.postData && event.postData.type || '').toLowerCase();
  if (type.indexOf('application/json') >= 0) {
    try {
      return JSON.parse(event.postData.contents || '{}');
    } catch (error) {
      return {};
    }
  }
  return event.parameter || {};
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LEAD_HEADERS);
    sheet.setFrozenRows(1);
    return;
  }
  const current = sheet.getRange(1, 1, 1, LEAD_HEADERS.length).getValues()[0];
  if (current.join('|') !== LEAD_HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, LEAD_HEADERS.length).setValues([LEAD_HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function sendOwnerNotification_(lead) {
  const subject = 'New AKRD website lead: ' + (lead.focus || 'Strategy inquiry');
  const body = [
    'Lead ID: ' + lead.leadId,
    'Name: ' + lead.name,
    'Email: ' + lead.email,
    'Company: ' + (lead.company || 'Not provided'),
    'Focus: ' + (lead.focus || 'Not selected'),
    'Source: ' + (lead.sourcePage || 'Unknown'),
    '',
    lead.message
  ].join('\n');
  MailApp.sendEmail({
    to: lead.to,
    subject: subject,
    body: body,
    replyTo: lead.email,
    name: 'AK Randall Digital website'
  });
}

function isRateLimited_(email) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email)
    .map(byte => ('0' + (byte & 255).toString(16)).slice(-2))
    .join('')
    .slice(0, 24);
  const cache = CacheService.getScriptCache();
  const key = 'lead-' + digest;
  if (cache.get(key)) return true;
  cache.put(key, '1', 60);
  return false;
}

function cleanText_(value, maxLength) {
  return String(value || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, maxLength);
}

function safeCell_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function isEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ''));
}

function createLeadId_(date) {
  return 'AKRD-' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
