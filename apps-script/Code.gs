/**
 * Kalarigram Attendance API — Google Apps Script, bound to the Kalarigram Sheet
 *
 * SETUP (one time)
 * 1. Open the Sheet (Kalarigram Google account) → Extensions → Apps Script → paste this file.
 * 2. Project Settings → Script properties → add  PIN = <kiosk PIN>
 * 3. Select setup() → Run (authorise). Creates tabs + headers, seeds Classes, builds Dashboard.
 *    Then review/complete the Classes tab (locations, dates).
 * 4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → copy URL.
 * 5. Later changes: Deploy → Manage deployments → ✎ → Version: New version (same URL).
 *
 * API
 *   GET  ?action=all&pin=…              → { classes, students:{classId:[{id,name}]} }
 *   GET  ?action=classes&pin=…
 *   GET  ?action=students&class=ID&pin=…
 *   POST (text/plain JSON) {action:'attendance', pin, records:[{entryId,classId,studentId,ts}]}
 *        → one row per student + class + day; repeats are skipped but still reported as saved
 *   POST (text/plain JSON) {action:'register', pin, name, phone, classIds:[…]}  (Email / Emergency Contact columns kept, left blank)
 *        → {ok, id, enrolled} or, if the phone is already registered, {ok, id, name, existing:true}
 *   POST (text/plain JSON) {action:'dropin', pin, records:[{entryId,classId,name,phone,ts}]}
 *        → unregistered visitors, written to the Drop-ins tab; one row per phone + class + day
 */

const TZ = 'Asia/Kolkata';
const CACHE_KEY = 'bootstrap_v1';
const CACHE_TTL = 600; // seconds; also cleared on edits / registrations

const TABS = {
  Students:    ['Student ID', 'Name', 'Phone', 'Email', 'Emergency Contact', 'Status', 'Registered On'],
  Classes:     ['Class ID', 'Name', 'Location', 'Schedule', 'Type', 'Start', 'End', 'Active'],
  Enrollments: ['Student ID', 'Class ID', 'Enrolled On'],
  Attendance:  ['Timestamp', 'Date', 'Class ID', 'Student ID', 'Entry ID'],
  Payments:    ['Student ID', 'Amount', 'Date', 'Mode', 'Receipt No', 'Notes'],
  'Drop-ins':  ['Timestamp', 'Date', 'Class ID', 'Name', 'Phone', 'Entry ID'],
};

// Starter rows for Classes — edit locations/schedules/dates in the sheet afterwards.
const SEED_CLASSES = [
  ['KG-6AM',   'Kalarigram 6 AM',              'Kalarigram',    '6:00 AM daily',  'Regular',  '', '', true],
  ['KG-5PM-B', 'Kalarigram 5 PM Beginners',    'Kalarigram',    '5:00 PM daily',  'Regular',  '', '', true],
  ['KG-5PM-A', 'Kalarigram 5 PM Advanced',     'Kalarigram',    '5:00 PM daily',  'Regular',  '', '', true],
  ['BK',       'Bharat Kalari',                'Bharat Kalari', '',               'Regular',  '', '', true],
  ['DIP-AM',   'Diploma Morning',              'Kalarigram',    '7:00–9:00 AM',   'Diploma',  '', '', true],
  ['DIP-PM',   'Diploma Evening',              'Kalarigram',    '4:00–5:00 PM',   'Diploma',  '', '', true],
  ['WS-STICK', 'Workshop: Long & Short Stick', 'Kalarigram',    '',               'Workshop', new Date(2026, 11, 1), new Date(2027, 1, 28), true],
];

/* ---------------- setup ---------------- */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(TABS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(TABS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, TABS[name].length).setFontWeight('bold');
    }
  });
  const classes = ss.getSheetByName('Classes');
  if (classes.getLastRow() === 1) {
    classes.getRange(2, 1, SEED_CLASSES.length, SEED_CLASSES[0].length).setValues(SEED_CLASSES);
    classes.getRange(2, 8, SEED_CLASSES.length, 1).insertCheckboxes();
  }
  ss.getSheetByName('Students').getRange('C:C').setNumberFormat('@'); // phone as text
  ss.getSheetByName('Attendance').getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ss.getSheetByName('Attendance').getRange('B:B').setNumberFormat('yyyy-mm-dd');
  formatDropIns_(ss.getSheetByName('Drop-ins'));
  buildDashboard_(ss);
  clearCache();
}

function buildDashboard_(ss) {
  const sh = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  sh.clear();
  sh.getRange('A1:A4').setValues([['Today'], ['Unique students today'], ['Class check-ins today'], ['Drop-ins today']]);
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('dd mmm yyyy');
  sh.getRange('B2').setFormula('=COUNTUNIQUEIFS(Attendance!D2:D,Attendance!B2:B,B1)');
  sh.getRange('B3').setFormula('=IFERROR(ROWS(UNIQUE(FILTER(Attendance!C2:C&"|"&Attendance!D2:D,Attendance!B2:B=B1))),0)');
  sh.getRange('B4').setFormula("=COUNTIFS('Drop-ins'!B2:B,B1)");

  sh.getRange('A5:C5').setValues([['Class ID', 'Class', 'Present today']]);
  sh.getRange('A6').setFormula('=FILTER(Classes!A2:B,Classes!H2:H=TRUE)');
  sh.getRange('C6').setFormula('=MAP(A6:A,LAMBDA(c,IF(c="","",COUNTUNIQUEIFS(Attendance!D2:D,Attendance!C2:C,c,Attendance!B2:B,$B$1))))');

  sh.getRange('E5:F5').setValues([['Date', 'Unique students']]);
  sh.getRange('E6').setFormula('=IFERROR(SORT(UNIQUE(FILTER(Attendance!B2:B,Attendance!B2:B<>"")),1,FALSE),"")');
  sh.getRange('F6').setFormula('=MAP(E6:E,LAMBDA(d,IF(d="","",COUNTUNIQUEIFS(Attendance!D2:D,Attendance!B2:B,d))))');
  sh.getRange('E6:E').setNumberFormat('dd mmm yyyy');

  sh.getRange('A1:A4').setFontWeight('bold');
  sh.getRange('A5:F5').setFontWeight('bold');
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 230);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

/* ---------------- HTTP ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    checkPin_(p.pin);
    const data = getBootstrap_();
    switch (p.action) {
      case 'all':      return json_({ ok: true, classes: data.classes, students: data.students, updated: data.updated });
      case 'classes':  return json_({ ok: true, classes: data.classes });
      case 'students': return json_({ ok: true, students: data.students[p.class] || [] });
      default:         return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    checkPin_(body.pin);
    if (body.action === 'attendance') return json_(markAttendance_(body.records || []));
    if (body.action === 'register')   return json_(register_(body));
    if (body.action === 'dropin')     return json_(markDropIns_(body.records || []));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------------- actions ---------------- */

// Append-only. Skips entryIds already written (safe for client retries) and a second
// check-in for the same student + class + day. Skipped records are still reported as
// saved so the phone clears them from its queue.
function markAttendance_(records) {
  const valid = records.filter(r => r && r.entryId && r.classId && r.studentId);
  if (!valid.length) return { ok: true, saved: [] };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Attendance');
    const last = sh.getLastRow();
    const start = Math.max(2, last - 999);
    const recent = last >= 2 ? sh.getRange(start, 2, last - start + 1, 4).getValues() : []; // Date, Class, Student, Entry
    const seen = new Set(recent.map(r => String(r[3])));
    const seenDay = new Set(recent.map(r => dayKey_(r[0], r[1], r[2])));

    const rows = [];
    valid.forEach(r => {
      if (seen.has(String(r.entryId))) return;
      seen.add(String(r.entryId));
      const ts = r.ts ? new Date(r.ts) : new Date();
      const day = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd');
      const key = dayKey_(day, r.classId, r.studentId);
      if (seenDay.has(key)) return; // already present today
      seenDay.add(key);
      rows.push([ts, new Date(day + 'T00:00:00+05:30'), String(r.classId), String(r.studentId), String(r.entryId)]);
    });
    if (rows.length) sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, saved: valid.map(r => r.entryId) };
  } finally {
    lock.releaseLock();
  }
}

// Unregistered visitors. Same rules as attendance, keyed on phone instead of Student ID.
// The tab is created on first use, so setup() doesn't need re-running.
function markDropIns_(records) {
  const valid = records.filter(r => r && r.entryId && r.classId && String(r.name || '').trim() && normPhone_(r.phone).length === 10);
  if (!valid.length) return { ok: true, saved: records.filter(r => r && r.entryId).map(r => r.entryId) };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Drop-ins');
    if (!sh) {
      sh = ss.insertSheet('Drop-ins');
      sh.appendRow(TABS['Drop-ins']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, TABS['Drop-ins'].length).setFontWeight('bold');
      formatDropIns_(sh);
    }
    const last = sh.getLastRow();
    const start = Math.max(2, last - 999);
    const recent = last >= 2 ? sh.getRange(start, 2, last - start + 1, 5).getValues() : []; // Date, Class, Name, Phone, Entry
    const seen = new Set(recent.map(r => String(r[4])));
    const seenDay = new Set(recent.map(r => dayKey_(r[0], r[1], normPhone_(r[3]))));

    const rows = [];
    valid.forEach(r => {
      if (seen.has(String(r.entryId))) return;
      seen.add(String(r.entryId));
      const ts = r.ts ? new Date(r.ts) : new Date();
      const day = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd');
      const phone = normPhone_(r.phone);
      const key = dayKey_(day, r.classId, phone);
      if (seenDay.has(key)) return; // already checked in today
      seenDay.add(key);
      rows.push([ts, new Date(day + 'T00:00:00+05:30'), String(r.classId), String(r.name).trim(), phone, String(r.entryId)]);
    });
    if (rows.length) sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
    // invalid records are acknowledged too, so a bad entry can't block the phone's queue
    return { ok: true, saved: records.filter(r => r && r.entryId).map(r => r.entryId) };
  } finally {
    lock.releaseLock();
  }
}

function formatDropIns_(sh) {
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:E').setNumberFormat('@'); // phone as text
}

// One student per phone number: a repeat registration returns the existing ID instead.
function register_(b) {
  const name = String(b.name || '').trim();
  if (!name) throw new Error('name required');
  const phone = normPhone_(b.phone);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const st = ss.getSheetByName('Students');
    const n = st.getLastRow() - 1;
    const existing = n > 0 ? st.getRange(2, 1, n, 3).getValues() : []; // ID, Name, Phone

    if (phone.length === 10) {
      const dup = existing.find(r => r[0] && normPhone_(r[2]) === phone);
      if (dup) return { ok: true, id: String(dup[0]).trim(), name: String(dup[1]).trim(), existing: true };
    }

    const max = existing.reduce((m, r) => {
      const num = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      return isNaN(num) ? m : Math.max(m, num);
    }, 0);
    const id = 'KG-' + String(max + 1).padStart(3, '0');
    const now = new Date();

    st.appendRow([id, name, phone || String(b.phone || ''), String(b.email || ''), String(b.emergency || ''), 'Active', now]);

    const validClasses = new Set(getBootstrap_().classes.map(c => c.id));
    const enrol = (b.classIds || []).map(String).filter(c => validClasses.has(c)).map(c => [id, c, now]);
    if (enrol.length) {
      const en = ss.getSheetByName('Enrollments');
      en.getRange(en.getLastRow() + 1, 1, enrol.length, 3).setValues(enrol);
    }
    clearCache();
    return { ok: true, id, enrolled: enrol.map(r => r[1]) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- data + cache ---------------- */

function getBootstrap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const ss = SpreadsheetApp.getActive();
  const classes = rows_(ss, 'Classes')
    .filter(r => r[0] && isTrue_(r[7]))
    .map(r => ({ id: String(r[0]).trim(), name: String(r[1]), location: String(r[2]), schedule: String(r[3]), type: String(r[4]) }));

  const active = {};
  rows_(ss, 'Students').forEach(r => {
    const status = String(r[5] || 'Active').trim();
    if (r[0] && status === 'Active') active[String(r[0]).trim()] = String(r[1]).trim();
  });

  const students = {};
  classes.forEach(c => (students[c.id] = []));
  rows_(ss, 'Enrollments').forEach(r => {
    const sid = String(r[0]).trim(), cid = String(r[1]).trim();
    if (students[cid] && active[sid] && !students[cid].some(s => s.id === sid)) {
      students[cid].push({ id: sid, name: active[sid] }); // ID + name only, no personal data
    }
  });
  Object.values(students).forEach(list => list.sort((a, b) => a.name.localeCompare(b.name)));

  const data = { classes, students, updated: new Date().toISOString() };
  try { cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL); } catch (e) { /* >100KB: skip cache */ }
  return data;
}

function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// Simple trigger: manual edits to these tabs invalidate the cache.
function onEdit(e) {
  const name = e && e.range.getSheet().getName();
  if (['Students', 'Classes', 'Enrollments'].indexOf(name) !== -1) clearCache();
}

/* ---------------- helpers ---------------- */

function rows_(ss, name) {
  const sh = ss.getSheetByName(name);
  return sh && sh.getLastRow() > 1 ? sh.getDataRange().getValues().slice(1) : [];
}

function isTrue_(v) {
  return v === true || /^(true|yes|y|1)$/i.test(String(v).trim());
}

// Last 10 digits: '+91 98765-43210' → '9876543210'.
function normPhone_(v) {
  const d = String(v || '').replace(/\D/g, '');
  return d.length > 10 ? d.slice(-10) : d;
}

// 'yyyy-MM-dd|classId|studentId'; day is a Date (Attendance column B) or a 'yyyy-MM-dd' string.
function dayKey_(day, classId, studentId) {
  const d = day instanceof Date ? Utilities.formatDate(day, TZ, 'yyyy-MM-dd') : String(day).slice(0, 10);
  return d + '|' + String(classId).trim() + '|' + String(studentId).trim();
}

function checkPin_(given) {
  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return; // PIN is optional: add Script property PIN to switch it on
  if (String(given || '') !== pin) throw new Error('bad pin');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
