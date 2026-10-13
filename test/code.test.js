// Runs Code.gs in Node against small in-memory fakes of the Google services it
// uses, so the booking logic can be tested without deploying.  `node --test`
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import crypto from 'node:crypto';

const SOURCE = fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8');
const TZ = 'UTC';

// Columns whose values are plain text strings (number format '@'), so the
// display value equals the stored value.
class FakeRange {
  constructor(sheet, row, col, rows, cols) { Object.assign(this, { sheet, row, col, rows, cols }); }
  #map(fn) {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const line = [];
      for (let c = 0; c < this.cols; c++) line.push(fn(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  getValues() { return this.#map((r, c) => this.sheet.get(r, c)); }
  getDisplayValues() {
    return this.#map((r, c) => {
      const v = this.sheet.get(r, c);
      return Object.prototype.toString.call(v) === '[object Date]' ? v.toISOString().slice(0, 10) : String(v);
    });
  }
  setValues(values) {
    values.forEach((line, r) => line.forEach((v, c) => this.sheet.set(this.row + r, this.col + c, v)));
    return this;
  }
  setValue(v) { this.sheet.set(this.row, this.col, v); return this; }
  getRow() { return this.row; }
  getLastRow() { return this.row + this.rows - 1; }
  getColumn() { return this.col; }
  getLastColumn() { return this.col + this.cols - 1; }
  getSheet() { return this.sheet; }
  setNumberFormats() { return this; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setDataValidation() { return this; }
  setWrap() { return this; }
}

class FakeSheet {
  constructor(name) { this.name = name; this.cells = []; }
  getName() { return this.name; }
  get(r, c) { return this.cells[r - 1]?.[c - 1] ?? ''; }
  set(r, c, v) { (this.cells[r - 1] ??= [])[c - 1] = v; }
  getLastRow() { return this.cells.length; }
  getMaxRows() { return 1000; }
  getRange(row, col, rows = 1, cols = 1) {
    if (typeof row === 'string') return new FakeRange(this, 1, 1, 0, 0); // A1 notation: formatting only
    return new FakeRange(this, row, col, rows, cols);
  }
  deleteRow(r) { this.cells.splice(r - 1, 1); }
  setFrozenRows() {}
  setColumnWidth() {}
  setColumnWidths() {}
}

function makeEnv() {
  const sheets = new Map();
  const ss = {
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new FakeSheet(n); sheets.set(n, s); return s; },
    getSheets: () => [...sheets.values()],
    deleteSheet: (s) => sheets.delete(s.getName()),
    getSpreadsheetTimeZone: () => TZ,
    getId: () => 'sheet-id',
    getUrl: () => 'https://docs.google.com/spreadsheets/d/sheet-id',
  };
  const mail = [];
  const cache = new Map();
  const pad = (n) => String(n).padStart(2, '0');
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush() {},
      getUi: () => ({ alert() {}, createMenu() { return this; }, addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
      newDataValidation: () => ({ requireValueInList() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } }),
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache.get(k) ?? null, put: (k, v) => cache.set(k, v) }) },
    PropertiesService: { getScriptProperties: () => ({ setProperty() {}, getProperty: () => null }) },
    MailApp: { sendEmail: (m) => mail.push(m) },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      // Only the patterns Code.gs uses; the tests run in UTC.
      formatDate(d, tz, fmt) {
        const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
        if (fmt === 'yyyy-MM-dd') return date;
        if (fmt === "yyyy-MM-dd'T'HH:mm") return `${date}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
        if (fmt === 'EEEE, MMMM d, yyyy') {
          return new Intl.DateTimeFormat('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', timeZone: 'UTC' }).format(d);
        }
        throw new Error(`unexpected format ${fmt}`);
      },
      parseDate(s) { return new ctx.Date(s.replace(' ', 'T') + (s.length > 10 ? ':00Z' : 'T00:00:00Z')); },
    },
    HtmlService: {},
  };
  vm.createContext(ctx);
  vm.runInContext(SOURCE, ctx);
  // Dates must come from the sandbox's own Date so `instanceof Date` holds.
  ctx.Date = vm.runInContext('Date', ctx);
  const date = (ymd) => new ctx.Date(`${ymd}T00:00:00Z`);
  return { ctx, ss, sheets, mail, date };
}

const futureDate = (days) => new Date(Date.now() + days * 86400000).toISOString().slice(0, 10);

function validForm(slotId, extra = {}) {
  return {
    slotId, name: 'Pat Example', email: 'pat@example.com', phone: '', preferred: 'Email',
    reason: 'Follow-up on blood pressure', consent: true, ...extra,
  };
}

test('parseTimeText_ handles common formats', () => {
  const { ctx } = makeEnv();
  assert.equal(ctx.parseTimeText_('9:00 AM'), '09:00');
  assert.equal(ctx.parseTimeText_('12:30 pm'), '12:30');
  assert.equal(ctx.parseTimeText_('12:00 AM'), '00:00');
  assert.equal(ctx.parseTimeText_('1:30 p.m.'), '13:30');
  assert.equal(ctx.parseTimeText_('9am'), '09:00');
  assert.equal(ctx.parseTimeText_('13:45'), '13:45');
  assert.equal(ctx.parseTimeText_('13:45:00'), '13:45');
  assert.equal(ctx.parseTimeText_('25:00'), '');
  assert.equal(ctx.parseTimeText_('13 pm'), '');
  assert.equal(ctx.parseTimeText_('banana'), '');
});

test('buildSlotTimes_ makes consecutive and weekly slots', () => {
  const { ctx } = makeEnv();
  const slots = ctx.buildSlotTimes_('2026-12-28', '09:00', '10:15', 30, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(slots)), [
    { date: '2026-12-28', start: '09:00', end: '09:30' },
    { date: '2026-12-28', start: '09:30', end: '10:00' },
    { date: '2027-01-04', start: '09:00', end: '09:30' },
    { date: '2027-01-04', start: '09:30', end: '10:00' },
  ]);
});

test('validateRequest_ enforces required fields and consistency', () => {
  const { ctx } = makeEnv();
  assert.ok(!ctx.validateRequest_(validForm('S-1234ABCD')).error);
  assert.match(ctx.validateRequest_(validForm('')).error, /choose an appointment time/);
  assert.match(ctx.validateRequest_(validForm('S-1', { name: ' ' })).error, /full name/);
  assert.match(ctx.validateRequest_(validForm('S-1', { email: '', phone: '' })).error, /email address or phone/);
  assert.match(ctx.validateRequest_(validForm('S-1', { email: 'nope' })).error, /valid email/);
  assert.match(ctx.validateRequest_(validForm('S-1', { preferred: 'Text message' })).error, /phone number/);
  assert.match(ctx.validateRequest_(validForm('S-1', { phone: '12' , preferred: 'Phone call' })).error, /valid phone/);
  assert.match(ctx.validateRequest_(validForm('S-1', { consent: false })).error, /privacy notice/);
  assert.match(ctx.validateRequest_(validForm('S-1', { reason: '' })).error, /what you would like/);
});

test('cellSafe_ neutralizes spreadsheet formulas', () => {
  const { ctx } = makeEnv();
  assert.equal(ctx.cellSafe_('=IMPORTXML("http://evil")'), '\'=IMPORTXML("http://evil")');
  assert.equal(ctx.cellSafe_('+15555550100'), "'+15555550100");
  assert.equal(ctx.cellSafe_('Knee pain'), 'Knee pain');
});

test('full flow: add slots, book, prevent double-booking, release on decline', () => {
  const { ctx, sheets, mail } = makeEnv();
  ctx.CONFIG.NOTIFY_EMAIL = 'np@example.com';
  ctx.setup();
  const day = futureDate(3);
  assert.equal(ctx.addSlots({ date: day, start: '09:00', end: '10:00', minutes: '30', repeatWeeks: '0' }), 2);
  // Adding the same block again is a no-op.
  assert.equal(ctx.addSlots({ date: day, start: '09:00', end: '10:00', minutes: '30', repeatWeeks: '0' }), 0);

  let open = ctx.getAvailableSlots();
  assert.equal(open.length, 2);
  assert.equal(open[0].timeLabel, '9:00 AM – 9:30 AM');
  // Public listing never exposes the row or booking info.
  assert.deepEqual(Object.keys(open[0]).sort(), ['date', 'dateLabel', 'id', 'timeLabel']);

  const res = ctx.submitRequest(validForm(open[0].id, { reason: '=HYPERLINK("x")' }));
  assert.equal(res.ok, true);
  assert.match(res.ref, /^REQ-[0-9A-F]{8}$/);

  const requests = sheets.get('Requests');
  assert.equal(requests.getLastRow(), 2);
  assert.equal(requests.get(2, 1), res.ref);
  assert.equal(requests.get(2, 5), 'Pending');
  assert.equal(requests.get(2, 6), 'Pat Example');
  assert.equal(requests.get(2, 10), '\'=HYPERLINK("x")');

  // Notification email contains no PHI.
  assert.equal(mail.length, 1);
  assert.ok(!mail[0].body.includes('Pat') && !mail[0].body.includes('HYPERLINK'));

  // Slot is gone from the public list and can't be double-booked.
  open = ctx.getAvailableSlots();
  assert.equal(open.length, 1);
  const again = ctx.submitRequest(validForm(sheets.get('Availability').get(2, 1)));
  assert.equal(again.ok, false);
  assert.equal(again.slotTaken, true);

  // NP declines the request → slot opens again.
  requests.set(2, 5, 'Declined');
  ctx.onEdit({ range: requests.getRange(2, 5), source: { getSheetByName: (n) => sheets.get(n) } });
  assert.equal(ctx.getAvailableSlots().length, 2);
});

test('manually typed availability rows get IDs and are offered', () => {
  const { ctx, sheets, date } = makeEnv();
  ctx.setup();
  const av = sheets.get('Availability');
  av.set(2, 2, date(futureDate(5)));
  av.set(2, 3, '2:00 PM');
  av.set(2, 4, '2:30 PM');
  const open = ctx.getAvailableSlots();
  assert.equal(open.length, 1);
  assert.match(av.get(2, 1), /^S-/);
  assert.equal(open[0].timeLabel, '2:00 PM – 2:30 PM');
});

test('past slots are hidden and honeypot submissions are dropped', () => {
  const { ctx, sheets, date } = makeEnv();
  ctx.setup();
  const av = sheets.get('Availability');
  av.set(2, 1, 'S-PAST0001');
  av.set(2, 2, date(futureDate(-2)));
  av.set(2, 3, '9:00 AM');
  av.set(2, 4, '9:30 AM');
  assert.equal(ctx.getAvailableSlots().length, 0);
  assert.equal(ctx.submitRequest(validForm('S-PAST0001')).ok, false);

  const bot = ctx.submitRequest(validForm('S-PAST0001', { website: 'http://spam' }));
  assert.equal(bot.ok, true);
  assert.equal(sheets.get('Requests').getLastRow(), 1);
});

test('rate limit caps submissions per hour', () => {
  const { ctx } = makeEnv();
  ctx.CONFIG.MAX_REQUESTS_PER_HOUR = 1;
  ctx.setup();
  ctx.addSlots({ date: futureDate(2), start: '09:00', end: '10:00', minutes: '30', repeatWeeks: '0' });
  const [a, b] = ctx.getAvailableSlots();
  assert.equal(ctx.submitRequest(validForm(a.id)).ok, true);
  const second = ctx.submitRequest(validForm(b.id));
  assert.equal(second.ok, false);
  assert.match(second.error, /lot of requests/);
});
