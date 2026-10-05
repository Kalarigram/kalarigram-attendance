/**
 * Kalarigram Attendance API — Google Apps Script, bound to the Kalarigram Sheet
 * v2 (2026-10-05): drop-ins, same-day duplicate guard, phone checks, Bharat Kalari classes, dashboard split.
 *
 * SETUP / UPGRADE
 * 1. Select setup() → Run. Safe to run any time: creates missing tabs/headers, runs the one-off v2
 *    migration (only once), rebuilds the Dashboard, clears the cache. Never deletes data.
 * 2. Deploy → Manage deployments → ✎ → Version: New version (keeps the same URL).
 * 3. Optional: Script property FRONTEND_URL = GitHub Pages check-in page, then run getClassLinks()
 *    to write one ?class= link per active class into the "QR Links" tab.
 * 4. Optional: Script property PIN switches on the PIN check (off when not set).
 *
 * API (all responses JSON)
 *   GET  ?action=all                    → { ok, classes:[{id,name,location,schedule,type}], students:{classId:[{id,name}]}, updated }
 *   GET  ?action=classes                → { ok, classes }
 *   GET  ?action=students&class=ID      → { ok, students:[{id,name}] }
 *   POST (text/plain JSON) {action:'attendance', records:[ record, ... ]}
 *        record = { entryId, classId, ts, type:'Regular'|'Drop-in', studentId?, name?, phone?, method? }
 *        (type missing → 'Regular', so old clients keep working)
 *        → { ok, saved:[entryId...], results:[{ entryId, status, message, personId?, name?, visit? }] }
 *          status: 'saved' | 'duplicate' | 'rejected'
 *          'saved' also covers an entryId that was already written (safe retry).
 *          Every processed entryId is listed in "saved", so the client can always clear it from its queue.
 *   POST {action:'register', name, phone, email, emergency, classIds:[...]}
 *        → { ok, id, enrolled:[classId...], convertedFrom? }  or { ok:false, error }
 *          Rejects an invalid phone, or a phone already used by an Active student.
 */

const TZ = 'Asia/Kolkata';
const CACHE_KEY = 'bootstrap_v1';
const CACHE_TTL = 600; // seconds; also cleared on edits / registrations
const SCAN_ROWS = 2000; // recent Attendance rows scanned for duplicates

const TABS = {
  Students:    ['Student ID', 'Name', 'Phone', 'Email', 'Emergency Contact', 'Status', 'Registered On'],
  Classes:     ['Class ID', 'Name', 'Location', 'Schedule', 'Type', 'Start', 'End', 'Active'],
  Enrollments: ['Student ID', 'Class ID', 'Enrolled On'],
  Attendance:  ['Timestamp', 'Date', 'Class ID', 'Student ID', 'Entry ID', 'Person Type', 'Name', 'Method', 'Notes'],
  Dropins:     ['Drop-in ID', 'Name', 'Phone', 'First Seen', 'Last Seen', 'Visits', 'Converted To', 'Notes'],
  Payments:    ['Student ID', 'Amount', 'Date', 'Mode', 'Receipt No', 'Notes', 'Person Type', 'Class ID', 'Purpose'],
};

const BK_CLASSES = [
  ['BK-KIDS',  'Bharat Kalari Kids',            'Bharat Kalari', '4:30 PM daily', 'Regular', '', '', true],
  ['BK-7AM',   'Bharat Kalari Morning 7 AM',    'Bharat Kalari', '7:00 AM daily', 'Regular', '', '', true],
  ['BK-530PM', 'Bharat Kalari Evening 5:30 PM', 'Bharat Kalari', '5:30 PM daily', 'Regular', '', '', true],
];

// Starter rows for an empty Classes tab (fresh install only).
const SEED_CLASSES = [
  ['KG-6AM',   'Kalarigram 6 AM',              'Kalarigram',    '6:00 AM daily',  'Regular',  '', '', true],
  ['KG-5PM-B', 'Kalarigram 5 PM Beginners',    'Kalarigram',    '5:00 PM daily',  'Regular',  '', '', true],
  ['KG-5PM-A', 'Kalarigram 5 PM Advanced',     'Kalarigram',    '5:00 PM daily',  'Regular',  '', '', true],
  ['DIP-AM',   'Diploma Morning',              'Kalarigram',    '7:00–9:00 AM',   'Diploma',  '', '', true],
  ['DIP-PM',   'Diploma Evening',              'Kalarigram',    '4:00–5:00 PM',   'Diploma',  '', '', true],
  ['WS-STICK', 'Workshop: Long & Short Stick', 'Kalarigram',    '',               'Workshop', new Date(2026, 11, 1), new Date(2027, 1, 28), true],
].concat(BK_CLASSES);

/* ---------------- setup + migration ---------------- */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(TABS).forEach(function (name) { ensureHeaders_(ss, name, TABS[name]); });

  const classes = sheet_(ss, 'Classes');
  if (classes.getLastRow() === 1) {
    classes.getRange(2, 1, SEED_CLASSES.length, SEED_CLASSES[0].length).setValues(SEED_CLASSES);
    classes.getRange(2, 8, SEED_CLASSES.length, 1).insertCheckboxes().setValue(true);
  }

  migrateV2_(ss);

  sheet_(ss, 'Students').getRange('C:C').setNumberFormat('@'); // phone as text
  sheet_(ss, 'Dropins').getRange('C:C').setNumberFormat('@');
  sheet_(ss, 'Dropins').getRange('D:E').setNumberFormat('yyyy-mm-dd');
  sheet_(ss, 'Attendance').getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet_(ss, 'Attendance').getRange('B:B').setNumberFormat('yyyy-mm-dd');
  buildDashboard_(ss);
  clearCache();
}

// Writes a header only where the cell is empty — never renames or moves existing columns.
function ensureHeaders_(ss, name, headers) {
  const sh = sheet_(ss, name) || ss.insertSheet(name, ss.getSheets().length);
  const width = Math.max(sh.getLastColumn(), headers.length);
  const existing = sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, width).getValues()[0] : [];
  headers.forEach(function (h, i) {
    if (String(existing[i] || '').trim() === '') sh.getRange(1, i + 1).setValue(h);
  });
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return sh;
}

// One-off upgrade to schema v2. Guarded by Script property SCHEMA_VERSION.
function migrateV2_(ss) {
  const props = PropertiesService.getScriptProperties();
  if (Number(props.getProperty('SCHEMA_VERSION') || 1) >= 2) return;

  // Classes: add the three Bharat Kalari classes, switch off the old single "BK".
  const cl = sheet_(ss, 'Classes');
  const ids = rows_(ss, 'Classes').map(function (r) { return String(r[0]).trim(); });
  const add = BK_CLASSES.filter(function (r) { return ids.indexOf(r[0]) === -1; });
  if (add.length) {
    const start = cl.getLastRow() + 1;
    cl.getRange(start, 1, add.length, 8).setValues(add);
    cl.getRange(start, 8, add.length, 1).insertCheckboxes().setValue(true);
  }
  const bk = ids.indexOf('BK');
  if (bk > -1) cl.getRange(bk + 2, 8).setValue(false);

  // Attendance: backfill Person Type / Name / Method on existing rows.
  const at = sheet_(ss, 'Attendance');
  const n = at.getLastRow() - 1;
  if (n > 0) {
    const names = {};
    rows_(ss, 'Students').forEach(function (r) { names[String(r[0]).trim()] = String(r[1]).trim(); });
    const d = at.getRange(2, 4, n, 5).getValues(); // D..H
    const out = d.map(function (r) {
      return [r[2] || 'Regular', r[3] || names[String(r[0]).trim()] || '', r[4] || 'QR'];
    });
    at.getRange(2, 6, n, 3).setValues(out);
  }

  props.setProperty('SCHEMA_VERSION', '2');
}

function buildDashboard_(ss) {
  const sh = sheet_(ss, 'Dashboard') || ss.insertSheet('Dashboard', ss.getSheets().length);
  sh.clear();
  const A = 'Attendance!';
  sh.getRange('A1:A5').setValues([['Today'], ['Unique people today'], ['Regular students today'], ['Drop-ins today'], ['Class check-ins today']]);
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('dd mmm yyyy');
  sh.getRange('B2').setFormula('=COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'B2:B,B1)');
  sh.getRange('B3').setFormula('=COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'B2:B,B1,' + A + 'F2:F,"<>Drop-in")');
  sh.getRange('B4').setFormula('=COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'B2:B,B1,' + A + 'F2:F,"Drop-in")');
  sh.getRange('B5').setFormula('=IFERROR(ROWS(UNIQUE(FILTER(' + A + 'C2:C&"|"&' + A + 'D2:D,' + A + 'B2:B=B1))),0)');

  // Per class, today (active classes, grouped by location)
  sh.getRange('A7:F7').setValues([['Class ID', 'Class', 'Location', 'Regular', 'Drop-in', 'Total']]);
  sh.getRange('A8').setFormula('=IFERROR(SORT(FILTER(Classes!A2:C,Classes!H2:H=TRUE),3,TRUE,1,TRUE),"")');
  sh.getRange('D8').setFormula('=MAP(A8:A,LAMBDA(c,IF(c="","",COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'C2:C,c,' + A + 'B2:B,$B$1,' + A + 'F2:F,"<>Drop-in"))))');
  sh.getRange('E8').setFormula('=MAP(A8:A,LAMBDA(c,IF(c="","",COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'C2:C,c,' + A + 'B2:B,$B$1,' + A + 'F2:F,"Drop-in"))))');
  sh.getRange('F8').setFormula('=MAP(A8:A,D8:D,E8:E,LAMBDA(c,r,d,IF(c="","",r+d)))');

  // Per day
  sh.getRange('H7:J7').setValues([['Date', 'Unique people', 'Drop-ins']]);
  sh.getRange('H8').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + A + 'B2:B,' + A + 'B2:B<>"")),1,FALSE),"")');
  sh.getRange('I8').setFormula('=MAP(H8:H,LAMBDA(d,IF(d="","",COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'B2:B,d))))');
  sh.getRange('J8').setFormula('=MAP(H8:H,LAMBDA(d,IF(d="","",COUNTUNIQUEIFS(' + A + 'D2:D,' + A + 'B2:B,d,' + A + 'F2:F,"Drop-in"))))');
  sh.getRange('H8:H').setNumberFormat('dd mmm yyyy');
  sh.getRange('D8:F').setNumberFormat('0');
  sh.getRange('I8:J').setNumberFormat('0');

  sh.getRange('A1:A5').setFontWeight('bold');
  sh.getRange('A7:J7').setFontWeight('bold');
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 120);
  const blank = sheet_(ss, 'Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

// Writes one check-in link per active class to the "QR Links" tab (needs Script property FRONTEND_URL).
function getClassLinks() {
  const base = PropertiesService.getScriptProperties().getProperty('FRONTEND_URL');
  if (!base) throw new Error('Add Script property FRONTEND_URL (the GitHub Pages check-in page URL) first.');
  const ss = SpreadsheetApp.getActive();
  const sh = sheet_(ss, 'QR Links') || ss.insertSheet('QR Links', ss.getSheets().length);
  sh.clear();
  const sep = base.indexOf('?') > -1 ? '&' : '?';
  const rows = rows_(ss, 'Classes').filter(function (r) { return r[0] && isTrue_(r[7]); }).map(function (r) {
    return [String(r[0]).trim(), String(r[1]), String(r[2]), base + sep + 'class=' + encodeURIComponent(String(r[0]).trim())];
  });
  sh.getRange(1, 1, 1, 4).setValues([['Class ID', 'Class', 'Location', 'Check-in link']]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
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
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------------- attendance ---------------- */

// Rules: one check-in per person per class per date (the record's own date, so offline
// records are checked against the day they happened). Repeat entryIds are skipped quietly.
function markAttendance_(records) {
  const list = (records || []).filter(function (r) { return r && r.entryId && r.classId; });
  if (!list.length) return { ok: true, saved: [], results: [] };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = sheet_(ss, 'Attendance');
    const last = sh.getLastRow();
    const start = Math.max(2, last - SCAN_ROWS + 1);
    const recent = last >= 2 ? sh.getRange(start, 1, last - start + 1, 5).getValues() : [];
    const seenEntry = new Set();
    const seenKey = new Set();
    recent.forEach(function (r) {
      seenEntry.add(String(r[4]));
      seenKey.add(fmtDate_(r[1]) + '|' + String(r[2]).trim() + '|' + String(r[3]).trim());
    });

    const ctx = context_(ss);
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    let drops = null;
    const rows = [], results = [], saved = [];

    list.forEach(function (r) {
      const entryId = String(r.entryId);
      const done = function (status, message, extra) {
        const res = { entryId: entryId, status: status, message: message };
        if (extra) Object.keys(extra).forEach(function (k) { res[k] = extra[k]; });
        results.push(res);
        saved.push(entryId);
      };

      if (seenEntry.has(entryId)) return done('saved', 'Already recorded.');

      const classId = String(r.classId).trim();
      const cls = ctx.classes[classId];
      if (!cls) return done('rejected', 'Unknown class: ' + classId);

      let ts = r.ts ? new Date(r.ts) : new Date();
      if (isNaN(ts.getTime())) ts = new Date();
      const day = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd');
      const when = day === today ? 'today' : 'on ' + day;
      const method = String(r.method || 'QR');
      const wantsDrop = /^drop/i.test(String(r.type || ''));

      let personId = '', name = '', personType = 'Regular', notes = '', drop = null, isNewDrop = false, phone = '';

      if (!wantsDrop) {
        personId = String(r.studentId || '').trim();
        if (!personId) return done('rejected', 'Please select your name.');
        const st = ctx.students[personId];
        name = st ? st.name : String(r.studentName || r.name || '').trim();
        if (!st) notes = 'unknown student ID';
        else if (!ctx.enrolled.has(personId + '|' + classId)) notes = 'not enrolled in this class';
      } else {
        name = String(r.name || '').trim();
        phone = normalizePhone_(r.phone);
        if (name.length < 2) return done('rejected', 'Please enter your full name.');
        if (!phone) return done('rejected', 'Please enter a valid phone number.');
        const sid = ctx.phoneToStudent[phone];
        if (sid) {
          // Registered student used the drop-in form → count them as a regular check-in.
          personId = sid;
          name = ctx.students[sid].name;
          notes = 'checked in via drop-in form' + (ctx.enrolled.has(sid + '|' + classId) ? '' : '; not enrolled in this class');
        } else {
          personType = 'Drop-in';
          drops = drops || loadDropins_(ss);
          drop = drops.byPhone[phone] || null;
          if (drop) { personId = drop.id; if (drop.name) name = drop.name; }
          else isNewDrop = true;
        }
      }

      if (!isNewDrop) {
        const key = day + '|' + classId + '|' + personId;
        if (seenKey.has(key)) {
          return done('duplicate', (name || 'This person') + ' is already checked in for ' + cls.name + ' ' + when + '.',
                      { personId: personId, name: name });
        }
      }

      if (personType === 'Drop-in') {
        if (isNewDrop) {
          drops.max += 1;
          drop = { id: 'DROP-' + String(drops.max).padStart(3, '0'), name: name, phone: phone,
                   firstSeen: day, lastSeen: day, visits: 0, row: 0, isNew: true, dirty: true };
          drops.byPhone[phone] = drop;
          drops.created.push(drop);
          personId = drop.id;
        }
        drop.visits += 1;
        if (!drop.lastSeen || day > drop.lastSeen) drop.lastSeen = day;
        if (!drop.firstSeen || day < drop.firstSeen) drop.firstSeen = day;
        drop.dirty = true;
      }

      rows.push([ts, dateVal_(day), classId, personId, entryId, personType, name, method, notes]);
      seenEntry.add(entryId);
      seenKey.add(day + '|' + classId + '|' + personId);

      let msg;
      if (personType === 'Drop-in') msg = drop.visits === 1 ? 'Welcome, ' + name + '! Drop-in recorded.' : 'Welcome back, ' + name + '! Visit #' + drop.visits + '.';
      else if (wantsDrop) msg = 'Welcome back, ' + name + '! You are a registered student — next time choose Regular.';
      else msg = 'Welcome back, ' + name + '!';
      done('saved', msg, { personId: personId, name: name, personType: personType, visit: drop ? drop.visits : undefined });
    });

    if (rows.length) sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
    if (drops) saveDropins_(drops);
    return { ok: true, saved: saved, results: results };
  } finally {
    lock.releaseLock();
  }
}

function context_(ss) {
  const classes = {};
  rows_(ss, 'Classes').forEach(function (r) {
    const id = String(r[0]).trim();
    if (id) classes[id] = { name: String(r[1]).trim() || id, active: isTrue_(r[7]) };
  });
  const students = {}, phoneToStudent = {};
  rows_(ss, 'Students').forEach(function (r) {
    const id = String(r[0]).trim();
    if (!id) return;
    const status = String(r[5] || 'Active').trim();
    students[id] = { name: String(r[1]).trim(), active: status === 'Active' };
    const pk = phoneKey_(r[2]);
    if (pk && status === 'Active' && !phoneToStudent[pk]) phoneToStudent[pk] = id;
  });
  const enrolled = new Set(rows_(ss, 'Enrollments').map(function (r) {
    return String(r[0]).trim() + '|' + String(r[1]).trim();
  }));
  return { classes: classes, students: students, phoneToStudent: phoneToStudent, enrolled: enrolled };
}

function loadDropins_(ss) {
  const sh = sheet_(ss, 'Dropins') || ensureHeaders_(ss, 'Dropins', TABS.Dropins);
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, 8).getValues() : [];
  const byPhone = {};
  let max = 0;
  vals.forEach(function (r, i) {
    const id = String(r[0]).trim();
    const num = parseInt(id.replace(/\D/g, ''), 10);
    if (!isNaN(num)) max = Math.max(max, num);
    const pk = phoneKey_(r[2]);
    if (id && pk && !byPhone[pk]) {
      byPhone[pk] = { id: id, name: String(r[1]).trim(), phone: pk, firstSeen: fmtDate_(r[3]), lastSeen: fmtDate_(r[4]),
                      visits: Number(r[5]) || 0, row: i + 2, isNew: false, dirty: false };
    }
  });
  return { sh: sh, byPhone: byPhone, max: max, created: [] };
}

function saveDropins_(d) {
  Object.keys(d.byPhone).forEach(function (k) {
    const x = d.byPhone[k];
    if (x.dirty && !x.isNew) d.sh.getRange(x.row, 4, 1, 3).setValues([[dateVal_(x.firstSeen), dateVal_(x.lastSeen), x.visits]]);
  });
  if (d.created.length) {
    const rows = d.created.map(function (x) {
      return [x.id, x.name, x.phone, dateVal_(x.firstSeen), dateVal_(x.lastSeen), x.visits, '', ''];
    });
    d.sh.getRange(d.sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  }
}

/* ---------------- registration ---------------- */

function register_(b) {
  const name = String(b.name || '').trim();
  if (!name) throw new Error('name required');
  const rawPhone = String(b.phone || '').trim();
  const phone = rawPhone ? normalizePhone_(rawPhone) : '';
  if (rawPhone && !phone) throw new Error('Please enter a valid phone number.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const st = sheet_(ss, 'Students');
    const existing = rows_(ss, 'Students');

    if (phone) {
      const dup = existing.filter(function (r) {
        return String(r[5] || 'Active').trim() === 'Active' && phoneKey_(r[2]) === phone;
      })[0];
      if (dup) throw new Error('This phone number is already registered as ' + String(dup[1]).trim() + ' (' + String(dup[0]).trim() + ').');
    }

    const max = existing.reduce(function (m, r) {
      const num = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      return isNaN(num) ? m : Math.max(m, num);
    }, 0);
    const id = 'KG-' + String(max + 1).padStart(3, '0');
    const now = new Date();

    st.appendRow([id, name, phone, String(b.email || '').trim(), String(b.emergency || '').trim(), 'Active', now]);

    const validClasses = new Set(getBootstrap_().classes.map(function (c) { return c.id; }));
    const enrol = (b.classIds || []).map(String).filter(function (c) { return validClasses.has(c); })
      .map(function (c) { return [id, c, now]; });
    if (enrol.length) {
      const en = sheet_(ss, 'Enrollments');
      en.getRange(en.getLastRow() + 1, 1, enrol.length, 3).setValues(enrol);
    }

    // A known drop-in has now registered: link the two records.
    let convertedFrom;
    if (phone) {
      const d = loadDropins_(ss);
      const x = d.byPhone[phone];
      if (x) { d.sh.getRange(x.row, 7).setValue(id); convertedFrom = x.id; }
    }

    clearCache();
    return { ok: true, id: id, enrolled: enrol.map(function (r) { return r[1]; }), convertedFrom: convertedFrom };
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
    .filter(function (r) { return r[0] && isTrue_(r[7]); })
    .map(function (r) { return { id: String(r[0]).trim(), name: String(r[1]), location: String(r[2]), schedule: String(r[3]), type: String(r[4]) }; });

  const active = {};
  rows_(ss, 'Students').forEach(function (r) {
    const status = String(r[5] || 'Active').trim();
    if (r[0] && status === 'Active') active[String(r[0]).trim()] = String(r[1]).trim();
  });

  const students = {};
  classes.forEach(function (c) { students[c.id] = []; });
  rows_(ss, 'Enrollments').forEach(function (r) {
    const sid = String(r[0]).trim(), cid = String(r[1]).trim();
    if (students[cid] && active[sid] && !students[cid].some(function (s) { return s.id === sid; })) {
      students[cid].push({ id: sid, name: active[sid] }); // ID + name only, no personal data
    }
  });
  Object.keys(students).forEach(function (k) { students[k].sort(function (a, b) { return a.name.localeCompare(b.name); }); });

  const data = { classes: classes, students: students, updated: new Date().toISOString() };
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

// Valid Indian mobile → 10 digits. International (+ not 91) → '+digits' (8–15). Invalid → ''.
function normalizePhone_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  let d = s.replace(/\D/g, '');
  if (s.charAt(0) === '+' && d.indexOf('91') !== 0) return d.length >= 8 && d.length <= 15 ? '+' + d : '';
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2);
  else if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return /^[6-9]\d{9}$/.test(d) ? d : '';
}

// Matching key for phones already in the sheet (lenient, so old/test data still matches).
function phoneKey_(raw) {
  const n = normalizePhone_(raw);
  if (n) return n;
  const d = String(raw == null ? '' : raw).replace(/\D/g, '');
  return d.length >= 10 ? d.slice(-10) : d;
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}

function dateVal_(day) {
  return day ? new Date(day + 'T00:00:00+05:30') : '';
}

// Looks a tab up by name. (getSheetByName fails in this spreadsheet with "Sheet 1759759947 not found", so we scan getSheets instead.)
function sheet_(ss, name) {
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) if (all[i].getName() === name) return all[i];
  return null;
}

function rows_(ss, name) {
  const sh = sheet_(ss, name);
  return sh && sh.getLastRow() > 1 ? sh.getDataRange().getValues().slice(1) : [];
}

function isTrue_(v) {
  return v === true || /^(true|yes|y|1)$/i.test(String(v).trim());
}

function checkPin_(given) {
  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return; // PIN is optional: add Script property PIN to switch it on
  if (String(given || '') !== pin) throw new Error('bad pin');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
