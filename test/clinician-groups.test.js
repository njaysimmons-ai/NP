// Runs clinician-groups/AssignClinicianGroups.gs in Node against in-memory
// fakes of the Sheets service.  `node --test`
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const SOURCE = fs.readFileSync(new URL('../clinician-groups/AssignClinicianGroups.gs', import.meta.url), 'utf8');

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
  setValues(values) {
    values.forEach((line, r) => line.forEach((v, c) => {
      const rule = this.sheet.validations.get(`${this.row + r},${this.col + c}`);
      if (rule && !rule.includes(v) && v !== '') throw new Error(`validation rejected ${v}`);
      this.sheet.set(this.row + r, this.col + c, v);
    }));
    return this;
  }
  getDataValidations() { return this.#map((r, c) => this.sheet.validations.get(`${r},${c}`) || null); }
  clearDataValidations() { this.#map((r, c) => this.sheet.validations.delete(`${r},${c}`)); return this; }
  setDataValidations(rules) {
    rules.forEach((line, r) => line.forEach((rule, c) => {
      if (rule) this.sheet.validations.set(`${this.row + r},${this.col + c}`, rule);
    }));
    return this;
  }
}

class FakeSheet {
  constructor(name, rows) { this.name = name; this.rows = rows.map((r) => r.slice()); this.validations = new Map(); }
  get(r, c) { const v = (this.rows[r - 1] || [])[c - 1]; return v === undefined || v === null ? '' : v; }
  set(r, c, v) { while (this.rows.length < r) this.rows.push([]); this.rows[r - 1][c - 1] = v; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map((r) => r.length)); }
  getRange(row, col, rows = 1, cols = 1) { return new FakeRange(this, row, col, rows, cols); }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
}

function run(sheets, activeName) {
  const alerts = [];
  const toasts = [];
  const ss = {
    getActiveSheet: () => sheets.find((s) => s.getName() === activeName),
    getSheets: () => sheets,
    getSheetByName: (n) => sheets.find((s) => s.getName() === n) || null,
    toast: (m) => toasts.push(m),
  };
  sheets.forEach((s, i) => { s.getIndex = () => i + 1; });
  const ctx = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ({ alert: (m) => alerts.push(m) }) },
    Math, String, parseInt,
  });
  vm.runInContext(SOURCE, ctx);
  ctx.assignClinicianGroups();
  return { alerts, toasts };
}

// Clinician Data as laid out in the real workbook: names in L, availability
// M–V, Unallowed W, Permanent X, Supervision Y, and group flags lower in L/M.
function clinicianData(people, flags = []) {
  const rows = [];
  const header = new Array(25).fill('');
  header[0] = 'CLINICIANS:';
  header[11] = 'Available';
  ['Mon\nAM', 'Mon\nPM', 'Tues\nAM', 'Tues\nPM', 'Wed\nAM', 'Wed\nPM', 'Thur\nAM', 'Thur\nPM', 'Fri\nAM', 'Fri\nPM']
    .forEach((h, i) => { header[12 + i] = h; });
  header[22] = 'Unallowed Groups';
  header[23] = 'Permanent Groups';
  header[24] = 'Supervision';
  rows.push(header);
  people.forEach((p) => {
    const row = new Array(25).fill('');
    row[11] = p.name;
    (p.avail || new Array(10).fill(true)).forEach((v, i) => { row[12 + i] = v; });
    row[22] = p.unallowed || '';
    row[23] = p.permanent || '';
    rows.push(row);
  });
  // a stray checkbox row with no name, like the real tab
  const stray = new Array(25).fill('');
  stray[10] = false;
  stray[12] = false;
  rows.push(stray);
  flags.forEach(([label, status]) => {
    const row = new Array(25).fill('');
    row[11] = label;
    row[12] = status;
    rows.push(row);
  });
  return new FakeSheet('Clinician Data', rows);
}

// Weekly tab: AM groups from amStart, PM groups after them, then Tasks, Out of
// Office, Unplaced.
function weekly(name, amLabels, pmLabels, tasks = {}) {
  const amStart = 3;
  const pmStart = amStart + amLabels.length;
  const tasksCol = pmStart + pmLabels.length;
  const rows = Array.from({ length: 36 }, () => []);
  rows[0][0] = name;
  rows[0][amStart - 1] = 'AM';
  rows[0][pmStart - 1] = 'PM';
  rows[0][tasksCol - 1] = 'Tasks ';
  rows[0][tasksCol] = 'Out of Office';
  rows[0][tasksCol + 1] = 'Unplaced';
  [2, 9, 16, 23, 30].forEach((lead) => {
    amLabels.concat(pmLabels).forEach((l, i) => { rows[lead + 2][amStart - 1 + i] = l; });
    amLabels.concat(pmLabels).forEach((_, i) => { rows[lead + 4][amStart - 1 + i] = false; });
    (tasks[lead] || []).forEach((t, i) => { rows[lead - 1 + i][tasksCol - 1] = t; });
  });
  const sheet = new FakeSheet(name, rows);
  sheet.layout = { amStart, pmStart, tasksCol, unplacedCol: tasksCol + 2 };
  return sheet;
}

const AM6 = ['Group 1 (PHP)', 'Group 2 (PHP)', 'Group 3 (PHP)', 'Group 4 (PHP)', 'Group 5 (IOP)', 'Group 6 (IOP)'];
const PM5 = ['Group 1 (PHP)', 'Group 2 (PHP)', 'Group 3 (PHP)', 'Group 4 (PHP)', 'Group 5 (IOP)'];
const NAMES = ['Ana', 'Ben', 'Cal', 'Dee', 'Eve', 'Fay', 'Gus', 'Hal', 'Ivy', 'Jo', 'Kim', 'Lou', 'Max', 'Ned'];

function leaderRow(sheet, row) {
  const { amStart, tasksCol } = sheet.layout;
  return Array.from({ length: tasksCol - amStart }, (_, i) => sheet.get(row, amStart + i));
}
function unplaced(sheet, row) {
  return Array.from({ length: 6 }, (_, i) => sheet.get(row + i, sheet.layout.unplacedCol)).filter(Boolean);
}

test('reads names from the column left of Mon AM and fills every active group', () => {
  const data = clinicianData(NAMES.map((name) => ({ name })));
  const week = weekly('Week', AM6, PM5);
  const { alerts, toasts } = run([week, data], 'Week');
  assert.deepEqual(alerts, []);
  assert.equal(toasts.length, 1);
  for (const row of [2, 9, 16, 23, 30]) {
    const placed = leaderRow(week, row);
    assert.ok(placed.every((n) => NAMES.includes(n)), `row ${row}: ${placed}`);
    assert.equal(new Set(placed).size, placed.length, 'nobody leads two groups in one day');
    assert.equal(unplaced(week, row).length, NAMES.length - placed.length);
  }
});

test('Unused groups stay empty and availability follows the right day', () => {
  // Ana is only available Tuesday PM (index 3)
  const avail = new Array(10).fill(false);
  avail[3] = true;
  const people = [{ name: 'Ana', avail }].concat(NAMES.slice(1).map((name) => ({ name })));
  const data = clinicianData(people, [['AM Group 1', 'Unused'], ['PM Group 4', 'Unused']]);
  const week = weekly('Week', AM6, PM5);
  run([week, data], 'Week');
  for (const row of [2, 9, 16, 23, 30]) {
    const placed = leaderRow(week, row);
    assert.equal(placed[0], '', 'AM Group 1 unused');
    assert.equal(placed[6 + 3], '', 'PM Group 4 unused');
    const anaAM = placed.slice(0, 6).includes('Ana');
    assert.equal(anaAM, false, 'Ana never placed in AM');
    if (row !== 9) {
      assert.ok(!placed.includes('Ana') && !unplaced(week, row).includes('Ana'), `Ana is out on row ${row}`);
    }
  }
});

test('works on a tab with fewer groups without touching Tasks', () => {
  const data = clinicianData(NAMES.map((name) => ({ name })));
  const week = weekly('Week', AM6.slice(0, 5), PM5.slice(0, 4), { 2: ['Ben', 'cal '] });
  run([week, data], 'Week');
  assert.equal(week.get(2, week.layout.tasksCol), 'Ben');
  assert.equal(week.get(3, week.layout.tasksCol), 'cal ');
  const placed = leaderRow(week, 2);
  assert.equal(placed.length, 9);
  assert.ok(!placed.includes('Ben') && !placed.includes('Cal'), 'on-task clinicians are not placed');
  assert.ok(!unplaced(week, 2).includes('Ben'), 'on-task clinicians are not unplaced');
});

test('permanent and unallowed groups are respected; validation is restored', () => {
  const people = NAMES.map((name) => ({ name }));
  people[0].permanent = 'AM Group 2';
  people[1].unallowed = 'Group 1, Group 3';
  const data = clinicianData(people);
  const week = weekly('Week', AM6, PM5);
  week.validations.set(`2,${week.layout.amStart}`, ['only-this']);
  run([week, data], 'Week');
  assert.deepEqual(week.validations.get(`2,${week.layout.amStart}`), ['only-this']);
  for (const row of [2, 9, 16, 23, 30]) {
    const placed = leaderRow(week, row);
    assert.equal(placed[1], 'Ana');
    [0, 2, 6, 8].forEach((i) => assert.notEqual(placed[i], 'Ben'));
  }
});

test('clinicians listed under Out of Office are skipped that day', () => {
  const data = clinicianData(NAMES.map((name) => ({ name })));
  const week = weekly('Week', AM6, PM5);
  week.set(30, week.layout.tasksCol + 1, 'Ana');
  run([week, data], 'Week');
  assert.ok(!leaderRow(week, 30).includes('Ana'));
  assert.ok(!unplaced(week, 30).includes('Ana'));
});

test('refuses to run on a tab without the expected headers', () => {
  const data = clinicianData(NAMES.map((name) => ({ name })));
  const other = new FakeSheet('Notes', [['hello']]);
  const { alerts } = run([other, data], 'Notes');
  assert.equal(alerts.length, 1);
  assert.match(alerts[0], /missing header/);
  const { alerts: a2 } = run([other, data], 'Clinician Data');
  assert.match(a2[0], /cannot be run/);
});
