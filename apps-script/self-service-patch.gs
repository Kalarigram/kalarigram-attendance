/**
 * Self-service patch for Code.gs (2026-10-03)
 * ---------------------------------------------------------------
 * Two server-side guards for the student-facing page:
 *   1. register: one student per phone number (returns the existing ID).
 *   2. attendance: one row per student + class + day.
 *
 * HOW TO APPLY (Apps Script editor → Code.gs):
 *   a) Paste the three helper functions below at the end of Code.gs.
 *   b) In register_(), BEFORE the script lock / new-ID code, add:
 *
 *        var existing = findStudentByPhone_(body.phone);
 *        if (existing) return { ok: true, id: existing.id, name: existing.name, existing: true };
 *
 *      (use whatever variable register_ receives the POST body as.)
 *   c) In the attendance handler, build the day keys once before the loop:
 *
 *        var seenDay = recentDayKeys_(1000);
 *
 *      and inside the loop, next to the existing entryId check:
 *
 *        var key = dayKey_(rec.ts, rec.classId, rec.studentId);
 *        if (seenDay[key]) { saved.push(rec.entryId); continue; }   // already present today
 *        seenDay[key] = true;
 *
 *      (skipped records still go in `saved` so the phone clears its queue.)
 *   d) Save → Deploy → Manage deployments → ✎ → Version: New version → Deploy.
 *      Same URL; never create a new deployment.
 *
 * Once the real Code.gs is committed to this folder, merge this in and delete this file.
 */

/** Last 10 digits of a phone number ('+91 98765-43210' → '9876543210'). */
function normPhone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  return d.length > 10 ? d.slice(-10) : d;
}

/** Students row whose phone matches, as {id, name}, or null. Students: A=ID, B=Name, C=Phone. */
function findStudentByPhone_(phone) {
  var p = normPhone_(phone);
  if (p.length < 10) return null;
  var sh = SpreadsheetApp.getActive().getSheetByName('Students');
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] && normPhone_(rows[i][2]) === p) return { id: String(rows[i][0]), name: String(rows[i][1]) };
  }
  return null;
}

/** 'yyyy-MM-dd|classId|studentId' in Asia/Kolkata time. */
function dayKey_(ts, classId, studentId) {
  var d = ts instanceof Date ? ts : new Date(ts);
  if (isNaN(d)) d = new Date();
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd') + '|' + classId + '|' + studentId;
}

/** Day keys already in the last n Attendance rows. Attendance: A=Timestamp, B=Date, C=Class ID, D=Student ID. */
function recentDayKeys_(n) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Attendance');
  var last = sh.getLastRow();
  var seen = {};
  if (last < 2) return seen;
  var start = Math.max(2, last - n + 1);
  var rows = sh.getRange(start, 1, last - start + 1, 4).getValues();
  rows.forEach(function (r) {
    var when = r[1] instanceof Date ? r[1] : r[0];   // prefer the Date column, fall back to Timestamp
    if (r[2] && r[3]) seen[dayKey_(when, r[2], r[3])] = true;
  });
  return seen;
}
