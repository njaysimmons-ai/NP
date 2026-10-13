/**
 * NP Appointment Requests — Google Sheets + Apps Script
 *
 * Patients open the web app link, pick an open time, and submit their name,
 * contact info and what they need. Requests land in the "Requests" tab of
 * this spreadsheet. The NP manages everything from the spreadsheet using
 * their existing Google sign-in; there are no separate usernames/passwords.
 *
 * See README.md for setup and the HIPAA checklist.
 */

// ─── EDIT THESE ──────────────────────────────────────────────────────────────
var CONFIG = {
  PRACTICE_NAME: 'Your Practice Name',
  PROVIDER_NAME: 'Your Name, NP',
  PRACTICE_PHONE: '(555) 555-0100',
  // Optional. Gets a "new request" alert that contains NO patient details,
  // just a link to the spreadsheet. Leave blank ('') to turn alerts off.
  NOTIFY_EMAIL: '',
  // How far ahead patients can see open times.
  MAX_DAYS_AHEAD: 60,
  // Flood protection: maximum submissions accepted per hour (all patients).
  MAX_REQUESTS_PER_HOUR: 30,
};
// ─────────────────────────────────────────────────────────────────────────────

var SHEET_AVAILABILITY = 'Availability';
var SHEET_REQUESTS = 'Requests';

var AVAIL_HEADERS = ['Slot ID', 'Date', 'Start', 'End', 'Booked by (Request ID)'];
var AV = { ID: 1, DATE: 2, START: 3, END: 4, BOOKED: 5 };

var REQ_HEADERS = ['Request ID', 'Submitted', 'Appointment Date', 'Appointment Time', 'Status',
  'Full Name', 'Email', 'Phone', 'Preferred Contact', 'Reason / Needs', 'NP Notes'];
var RQ = { ID: 1, SUBMITTED: 2, DATE: 3, TIME: 4, STATUS: 5 };

var STATUSES = ['Pending', 'Confirmed', 'Declined', 'Cancelled', 'Completed'];
// Setting a request to one of these puts its time back on the open list.
var RELEASING_STATUSES = ['Declined', 'Cancelled'];
var CONTACT_METHODS = ['Email', 'Phone call', 'Text message'];

// ─── Web app (patient-facing) ────────────────────────────────────────────────

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.config = CONFIG;
  return template.evaluate()
    .setTitle('Request an appointment · ' + CONFIG.PRACTICE_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Open, future, unbooked times. Contains no patient information. */
function getAvailableSlots() {
  var ss = getSpreadsheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = ensureAvailabilitySheet_(ss);
    ensureSlotIds_(sheet);
    var now = nowKey_(tz);
    var limit = Utilities.formatDate(new Date(Date.now() + CONFIG.MAX_DAYS_AHEAD * 86400000), tz, 'yyyy-MM-dd');
    return readSlots_(sheet, tz)
      .filter(function (s) { return !s.bookedBy && slotKey_(s) > now && s.date <= limit; })
      .sort(function (a, b) { return slotKey_(a) < slotKey_(b) ? -1 : 1; })
      .map(function (s) {
        return { id: s.id, date: s.date, dateLabel: dateLabel_(s.date, tz), timeLabel: timeRangeLabel_(s) };
      });
  } finally {
    lock.releaseLock();
  }
}

/** Called from the patient form. Returns {ok, ref, when} or {ok:false, error}. */
function submitRequest(form) {
  form = form || {};
  // Hidden "website" field: humans never fill it in, bots usually do.
  if (form.website) return { ok: true, ref: makeId_('REQ'), when: '' };

  var v = validateRequest_(form);
  if (v.error) return { ok: false, error: v.error };

  var ss = getSpreadsheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { ok: false, error: 'The system is busy. Please try again in a moment.' };
  }
  var ref, when;
  try {
    if (!rateLimitOk_()) {
      return { ok: false, error: 'We are receiving a lot of requests right now. Please try again later or call the office.' };
    }
    var avail = ensureAvailabilitySheet_(ss);
    var slot = readSlots_(avail, tz).filter(function (s) { return s.id === v.slotId; })[0];
    if (!slot || slot.bookedBy || slotKey_(slot) <= nowKey_(tz)) {
      return { ok: false, slotTaken: true, error: 'Sorry, that time is no longer available. Please choose another time.' };
    }

    ref = makeId_('REQ');
    when = dateLabel_(slot.date, tz) + ' at ' + timeRangeLabel_(slot);
    var requests = ensureRequestsSheet_(ss);
    var row = requests.getLastRow() + 1;
    var range = requests.getRange(row, 1, 1, REQ_HEADERS.length);
    range.setNumberFormats([['@', 'yyyy-mm-dd h:mm am/pm', 'ddd, mmm d, yyyy', '@', '@', '@', '@', '@', '@', '@', '@']]);
    range.setValues([[
      ref,
      new Date(),
      parseDateKey_(slot.date, tz),
      timeRangeLabel_(slot),
      'Pending',
      cellSafe_(v.name),
      cellSafe_(v.email),
      cellSafe_(v.phone),
      v.preferred,
      cellSafe_(v.reason),
      '',
    ]]);
    avail.getRange(slot.row, AV.BOOKED).setValue(ref);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  notifyNewRequest_(ss);
  return { ok: true, ref: ref, when: when };
}

// ─── Spreadsheet menu & triggers (NP-facing) ────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Appointments')
    .addItem('Add open time slots…', 'showAddSlotsSidebar')
    .addItem('Remove past unbooked slots', 'removePastOpenSlots')
    .addSeparator()
    .addItem('Set up / repair sheets', 'setup')
    .addToUi();
}

/**
 * Runs automatically when the NP edits the sheet:
 *  - Request marked Declined/Cancelled → its time slot opens up again.
 *  - Rows typed into Availability get a Slot ID.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  var firstRow = Math.max(2, e.range.getRow());
  var lastRow = e.range.getLastRow();
  if (lastRow < firstRow) return;

  if (name === SHEET_REQUESTS &&
      e.range.getColumn() <= RQ.STATUS && e.range.getLastColumn() >= RQ.STATUS) {
    var rows = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, RQ.STATUS).getValues();
    var released = rows
      .filter(function (r) { return RELEASING_STATUSES.indexOf(String(r[RQ.STATUS - 1])) !== -1; })
      .map(function (r) { return String(r[RQ.ID - 1]); });
    if (released.length) releaseSlots_(e.source, released);
  } else if (name === SHEET_AVAILABILITY) {
    ensureSlotIds_(sheet);
  }
}

function showAddSlotsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Add open time slots');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Called from the sidebar. Creates consecutive slots, optionally weekly. */
function addSlots(form) {
  var date = String(form.date || '');
  var start = String(form.start || '');
  var end = String(form.end || '');
  var minutes = parseInt(form.minutes, 10);
  var weeks = parseInt(form.repeatWeeks || '0', 10);

  if (!isValidDateKey_(date)) throw new Error('Please choose a valid date.');
  if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) throw new Error('Please choose start and end times.');
  if (!(minutes >= 5 && minutes <= 240)) throw new Error('Appointment length must be 5–240 minutes.');
  if (!(weeks >= 0 && weeks <= 26)) throw new Error('Repeat must be 0–26 weeks.');
  if (toMinutes_(end) - toMinutes_(start) < minutes) throw new Error('End time must be at least one appointment length after the start time.');

  var ss = getSpreadsheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = ensureAvailabilitySheet_(ss);
    var existing = {};
    readSlots_(sheet, tz).forEach(function (s) { existing[s.date + ' ' + s.start] = true; });
    var now = nowKey_(tz);
    var rows = buildSlotTimes_(date, start, end, minutes, weeks)
      .filter(function (s) { return !existing[s.date + ' ' + s.start] && slotKey_(s) > now; })
      .map(function (s) {
        return [makeId_('S'), parseDateKey_(s.date, tz), formatTime12_(s.start), formatTime12_(s.end), ''];
      });
    if (rows.length) {
      var range = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, AVAIL_HEADERS.length);
      range.setNumberFormats(rows.map(function () { return ['@', 'ddd, mmm d, yyyy', '@', '@', '@']; }));
      range.setValues(rows);
    }
    return rows.length;
  } finally {
    lock.releaseLock();
  }
}

function removePastOpenSlots() {
  var ss = getSpreadsheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var removed = 0;
  try {
    var sheet = ensureAvailabilitySheet_(ss);
    var now = nowKey_(tz);
    var past = readSlots_(sheet, tz).filter(function (s) { return !s.bookedBy && slotKey_(s) <= now; });
    // Delete bottom-up so row numbers stay valid.
    past.sort(function (a, b) { return b.row - a.row; }).forEach(function (s) { sheet.deleteRow(s.row); });
    removed = past.length;
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.getUi().alert('Removed ' + removed + ' past unbooked slot(s).');
}

/** Creates/repairs both tabs. Safe to run more than once. */
function setup() {
  var ss = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ensureAvailabilitySheet_(ss);
  ensureRequestsSheet_(ss);
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  try {
    SpreadsheetApp.getUi().alert('Sheets are ready. Next: Appointments → Add open time slots…');
  } catch (err) {
    // Not running from the spreadsheet UI (e.g. the script editor); nothing to show.
  }
}

// ─── Sheet helpers ───────────────────────────────────────────────────────────

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run Appointments → Set up / repair sheets first.');
  return SpreadsheetApp.openById(id);
}

function ensureAvailabilitySheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_AVAILABILITY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AVAILABILITY);
    sheet.getRange(1, 1, 1, AVAIL_HEADERS.length).setValues([AVAIL_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A2:A').setNumberFormat('@');
    sheet.getRange('B2:B').setNumberFormat('ddd, mmm d, yyyy');
    // Times are kept as plain text ("9:00 AM") so Sheets can't shift them.
    sheet.getRange('C2:E').setNumberFormat('@');
    sheet.setColumnWidths(1, 5, 150);
  }
  return sheet;
}

function ensureRequestsSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_REQUESTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REQUESTS);
    sheet.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(10, 360);
    sheet.getRange('J2:J').setWrap(true);
  }
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sheet.getRange(2, RQ.STATUS, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
  return sheet;
}

/** Reads Availability rows into {row, id, date:'yyyy-MM-dd', start:'HH:mm', end:'HH:mm', bookedBy}. */
function readSlots_(sheet, tz) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var range = sheet.getRange(2, 1, last - 1, AVAIL_HEADERS.length);
  var values = range.getValues();
  // Display values for times avoid Sheets' 1899-epoch time-zone quirks.
  var display = range.getDisplayValues();
  var slots = [];
  for (var i = 0; i < values.length; i++) {
    var date = toDateKey_(values[i][AV.DATE - 1], tz);
    var start = parseTimeText_(display[i][AV.START - 1]);
    var end = parseTimeText_(display[i][AV.END - 1]);
    var id = String(values[i][AV.ID - 1]).trim();
    if (!id || !date || !start || !end) continue;
    slots.push({ row: i + 2, id: id, date: date, start: start, end: end, bookedBy: String(values[i][AV.BOOKED - 1]).trim() });
  }
  return slots;
}

function ensureSlotIds_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var range = sheet.getRange(2, 1, last - 1, 2);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (r) {
    if (!String(r[0]).trim() && r[1] !== '') {
      r[0] = makeId_('S');
      changed = true;
    }
  });
  if (changed) sheet.getRange(2, 1, values.length, 1).setValues(values.map(function (r) { return [r[0]]; }));
}

function releaseSlots_(ss, requestIds) {
  var sheet = ss.getSheetByName(SHEET_AVAILABILITY);
  if (!sheet || sheet.getLastRow() < 2) return;
  var range = sheet.getRange(2, AV.BOOKED, sheet.getLastRow() - 1, 1);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (r) {
    if (r[0] && requestIds.indexOf(String(r[0])) !== -1) {
      r[0] = '';
      changed = true;
    }
  });
  if (changed) range.setValues(values);
}

function notifyNewRequest_(ss) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: CONFIG.NOTIFY_EMAIL,
      subject: 'New appointment request',
      body: 'A new appointment request was submitted.\n\n' +
        'For privacy, no patient details are included in this email. ' +
        'Open the Requests tab to review it:\n' + ss.getUrl(),
    });
  } catch (err) {
    // An email failure must not lose the patient's request.
  }
}

function rateLimitOk_() {
  var cache = CacheService.getScriptCache();
  var key = 'submissions_' + Math.floor(Date.now() / 3600000);
  var count = parseInt(cache.get(key) || '0', 10);
  if (count >= CONFIG.MAX_REQUESTS_PER_HOUR) return false;
  cache.put(key, String(count + 1), 3600);
  return true;
}

// ─── Pure helpers (unit-tested in test/code.test.js) ─────────────────────────

function validateRequest_(form) {
  var clean = function (s, max) {
    return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, max);
  };
  var v = {
    slotId: clean(form.slotId, 40),
    name: clean(form.name, 100).replace(/\s+/g, ' '),
    email: clean(form.email, 254),
    phone: clean(form.phone, 30),
    preferred: clean(form.preferred, 20),
    reason: clean(form.reason, 1000),
  };
  var digits = v.phone.replace(/\D/g, '');

  if (!/^[A-Za-z0-9-]{3,40}$/.test(v.slotId)) return { error: 'Please choose an appointment time.' };
  if (v.name.length < 2) return { error: 'Please enter your full name.' };
  if (!v.email && !v.phone) return { error: 'Please enter an email address or phone number so we can reach you.' };
  if (v.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.email)) return { error: 'Please enter a valid email address.' };
  if (v.phone && (!/^[0-9+().\-\s]+$/.test(v.phone) || digits.length < 7 || digits.length > 15)) {
    return { error: 'Please enter a valid phone number.' };
  }
  if (CONTACT_METHODS.indexOf(v.preferred) === -1) return { error: 'Please choose how you would like us to contact you.' };
  if (v.preferred === 'Email' && !v.email) return { error: 'You chose email. Please enter your email address.' };
  if (v.preferred !== 'Email' && !v.phone) return { error: 'You chose phone/text. Please enter your phone number.' };
  if (v.reason.length < 3) return { error: 'Please tell us briefly what you would like to discuss.' };
  if (form.consent !== true && form.consent !== 'yes') return { error: 'Please check the box to acknowledge the privacy notice.' };
  return v;
}

/** Stops a patient's text from being treated as a spreadsheet formula. */
function cellSafe_(s) {
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/** "9:00 AM", "9am", "13:30", "1:30 p.m." → "HH:mm" (24h), or '' if invalid. */
function parseTimeText_(text) {
  var t = String(text || '').trim().toLowerCase().replace(/\./g, '');
  var m = t.match(/^(\d{1,2})(?::(\d{2}))?(?::\d{2})?\s*(am|pm|a|p)?$/);
  if (!m) return '';
  var h = parseInt(m[1], 10);
  var min = m[2] ? parseInt(m[2], 10) : 0;
  var ap = m[3];
  if (min > 59) return '';
  if (ap) {
    if (h < 1 || h > 12) return '';
    if (ap.charAt(0) === 'p' && h !== 12) h += 12;
    if (ap.charAt(0) === 'a' && h === 12) h = 0;
  } else if (h > 23 || !m[2]) {
    return '';
  }
  return pad2_(h) + ':' + pad2_(min);
}

function formatTime12_(hhmm) {
  var h = parseInt(hhmm.slice(0, 2), 10);
  return (h % 12 || 12) + ':' + hhmm.slice(3) + ' ' + (h >= 12 ? 'PM' : 'AM');
}

function timeRangeLabel_(slot) {
  return formatTime12_(slot.start) + ' – ' + formatTime12_(slot.end);
}

function toMinutes_(hhmm) {
  return parseInt(hhmm.slice(0, 2), 10) * 60 + parseInt(hhmm.slice(3), 10);
}

function fromMinutes_(mins) {
  return pad2_(Math.floor(mins / 60)) + ':' + pad2_(mins % 60);
}

function addDaysKey_(dateKey, days) {
  var p = dateKey.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + days)).toISOString().slice(0, 10);
}

function isValidDateKey_(dateKey) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return false;
  return addDaysKey_(dateKey, 0) === dateKey;
}

function buildSlotTimes_(dateKey, start, end, minutes, repeatWeeks) {
  var out = [];
  for (var w = 0; w <= repeatWeeks; w++) {
    var day = addDaysKey_(dateKey, w * 7);
    for (var t = toMinutes_(start); t + minutes <= toMinutes_(end); t += minutes) {
      out.push({ date: day, start: fromMinutes_(t), end: fromMinutes_(t + minutes) });
    }
  }
  return out;
}

function slotKey_(slot) {
  return slot.date + 'T' + slot.start;
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

// ─── Apps Script–dependent helpers ───────────────────────────────────────────

function nowKey_(tz) {
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm");
}

function toDateKey_(value, tz) {
  if (value instanceof Date && !isNaN(value)) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  var s = String(value || '').trim();
  if (isValidDateKey_(s)) return s;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var key = m[3] + '-' + pad2_(+m[1]) + '-' + pad2_(+m[2]);
    if (isValidDateKey_(key)) return key;
  }
  return '';
}

function parseDateKey_(dateKey, tz) {
  return Utilities.parseDate(dateKey, tz, 'yyyy-MM-dd');
}

function dateLabel_(dateKey, tz) {
  // Noon avoids any chance of a daylight-saving boundary shifting the day.
  return Utilities.formatDate(Utilities.parseDate(dateKey + ' 12:00', tz, 'yyyy-MM-dd HH:mm'), tz, 'EEEE, MMMM d, yyyy');
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}
