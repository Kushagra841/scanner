// Google Apps Script: receives scans from the packer page and adds each one as a row.
// Paste this into Extensions > Apps Script inside the Google Sheet that should hold the scans.

const SECRET = "shantam@143";   // must be the same as SECRET in the page (index.html)
const ADMIN_KEY = "set-your-own-admin-password";   // ONLY for downloads. Never put this in the page.
const SHEET_NAME = "Scans";
const HEADERS = ["Scan ID", "Date", "Time", "Packer", "Barcode", "Repeat scan", "Received at", "Status", "Camera / Table",
                 "Times scanned", "First scan"];
// Repeat scans:
//  - same barcode again on the SAME day -> that day's row is updated (latest time, packer, camera;
//    Times scanned +1; First scan keeps the first one). No new row.
//  - same barcode on a LATER day        -> the old row is never touched. A new row is added with
//    Repeat scan = "OLD (first date)" so a wrong scan can be found later.
const STATUS_TEXT = "Forward";            // goes in the Status column of every row
const MAX_BATCH = 50;                     // most scans accepted in one request
const ARCHIVE_PREFIX = "Archive ";        // old months move to tabs like "Archive 2026-09"
const KEEP_MONTHS = 2;                    // months kept on the main tab: this month + last month

// Adds the Status column to an older sheet and fills it for rows already there
function ensureStatus(sh) {
  if (sh.getLastRow() < 1) return;
  [9, 10, 11].forEach(function (c) {          // newer columns get their header on an older sheet
    if (sh.getRange(1, c).getValue() !== HEADERS[c - 1]) sh.getRange(1, c).setValue(HEADERS[c - 1]).setFontWeight("bold");
  });
  if (sh.getRange(1, 8).getValue() === "Status") return;
  sh.getRange(1, 8).setValue("Status").setFontWeight("bold");
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 8, n, 1).setValue(STATUS_TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);                      // handles several packers scanning at the same moment
  try {
    const d = JSON.parse(e.postData.contents);

    // Download request from the page: needs the admin password
    if (d.action === "export") {
      if (ADMIN_KEY.indexOf("set-your-own") === 0) return reply({ ok: false, error: "admin password not set" });
      if (d.admin !== ADMIN_KEY) return reply({ ok: false, error: "wrong admin password" });
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const main = ss.getSheetByName(SHEET_NAME);
      if (main) ensureStatus(main);
      const tz = Session.getScriptTimeZone();
      const day = d.day ? String(d.day) : "";            // "2026-11-24" = only that day, "" = every day
      // where to look: that month's archive tab (if any) + the main tab; for all days, every tab
      const tabs = day ? [ss.getSheetByName(ARCHIVE_PREFIX + day.slice(0, 7)), main]
                       : archiveTabs(ss).concat([main]);
      let rows = [];
      tabs.forEach(function (s) { if (s) rows = rows.concat(readRows(s, day)); });
      rows = rows.map(function (r) {
        r[6] = r[6] instanceof Date ? Utilities.formatDate(r[6], tz, "yyyy-MM-dd HH:mm:ss") : r[6];
        return r.map(String);
      });
      return reply({ ok: true, headers: HEADERS, day: day, rows: rows });
    }

    if (d.key !== SECRET) return reply({ ok: false, error: "bad key" });

    // The page sends several waiting scans at once as "items"; an older page sends one scan.
    const batch = Array.isArray(d.items);
    const items = batch ? d.items.slice(0, MAX_BATCH) : [d];
    if (!items.length || items.some(function (it) { return !it || !it.id || !it.code; })) {
      return reply({ ok: false, error: "missing data" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
      sh.setFrozenRows(1);
      sh.getRange("A:E").setNumberFormat("@");        // keep barcodes and times exactly as typed
      sh.getRange("G:G").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }

    ensureStatus(sh);
    const last = sh.getLastRow();
    const now = new Date();
    const newRows = [], results = [];
    const idsInBatch = {}, codesInBatch = {};      // scans in this same request, not yet in the sheet
    const sameDay = {};                            // "barcode|date" -> { row: sheet row or 0 if new, v: values }
    const hasArchive = archiveTabs(ss).length > 0;

    items.forEach(function (it) {
      const id = String(it.id), code = String(it.code), date = String(it.date);

      // same scan sent twice (for example after a bad connection): ignore the second copy.
      // A row updated by repeat scans holds all their IDs, so the search is not whole-cell.
      if (idsInBatch[id] || (last > 1 && findText(sh.getRange(2, 1, last - 1, 1), id))) {
        results.push({ id: id, ok: true, repeat: false, duplicateRequest: true });
        return;
      }
      idsInBatch[id] = true;

      // 1) same barcode already scanned on this same day: update that row, no new row
      const key = code + "|" + date;
      if (!(key in sameDay)) {
        const r = last > 1 ? findSameDayRow(sh, last, code, date) : 0;
        sameDay[key] = r ? { row: r, v: sh.getRange(r, 1, 1, HEADERS.length).getValues()[0] } : null;
      }
      const t = sameDay[key];
      if (t) {
        const v = t.v;
        const first = v[10] ? String(v[10]) : describe(v);
        v[0] = String(v[0]) + " " + id;            // keep every scan ID in the row
        v[2] = String(it.time); v[3] = String(it.packer); v[6] = now; v[8] = String(it.table || "");
        v[9] = (Number(v[9]) || 1) + 1;
        v[10] = first;
        if (String(v[5]).indexOf("OLD") !== 0) v[5] = "YES";
        t.dirty = true;
        results.push({ id: id, ok: true, repeat: true, sameDay: true, times: v[9], firstSeen: first });
        return;
      }

      // 2) scanned on an earlier day? Google's own search looks through the whole Barcode
      // column at once, so it stays fast however many rows there are.
      // Archive tabs hold the oldest scans, so they are checked first.
      let firstSeen = "", firstDate = "";
      let hit = hasArchive ? findInArchives(ss, code) : null;
      if (!hit && last > 1) hit = findExact(sh.getRange(2, 5, last - 1, 1), code);   // first (oldest) match
      if (hit) {
        const r = hit.getSheet().getRange(hit.getRow(), 1, 1, HEADERS.length).getValues()[0];
        firstSeen = describe(r); firstDate = String(r[1]);
      }
      if (!firstSeen && codesInBatch[code]) { firstSeen = codesInBatch[code].who; firstDate = codesInBatch[code].date; }

      // 3) new row (marked OLD when the barcode was already scanned on an earlier day)
      const row = [id, date, String(it.time), String(it.packer), code,
                   firstSeen ? "OLD (" + firstDate + ")" : "", now, STATUS_TEXT, String(it.table || ""),
                   1, firstSeen];
      newRows.push(row);
      sameDay[key] = { row: 0, v: row };           // a later scan in this request updates this new row
      if (!codesInBatch[code]) codesInBatch[code] = { who: describe(row), date: date };
      results.push({ id: id, ok: true, repeat: !!firstSeen, old: !!firstSeen, firstSeen: firstSeen });
    });

    // rows updated by same-day repeat scans, then all new rows in one go
    Object.keys(sameDay).forEach(function (k) {
      const t = sameDay[k];
      if (t && t.row && t.dirty) sh.getRange(t.row, 1, 1, HEADERS.length).setValues([t.v]);
    });
    if (newRows.length) sh.getRange(last + 1, 1, newRows.length, HEADERS.length).setValues(newRows);

    if (batch) return reply({ ok: true, results: results });
    const one = results[0];                        // reply in the old shape for a single scan
    return reply(one.duplicateRequest ? { ok: true, repeat: false, duplicateRequest: true }
                                      : { ok: true, repeat: one.repeat, firstSeen: one.firstSeen,
                                          sameDay: !!one.sameDay, old: !!one.old, times: one.times });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Exact, case-sensitive search inside one column. Returns the first matching cell or null.
function findExact(range, value) {
  return range.createTextFinder(String(value)).matchEntireCell(true).matchCase(true).findNext();
}

// Like findExact, but the value may be one word among several in the cell (Scan ID column).
function findText(range, value) {
  return range.createTextFinder(String(value)).matchEntireCell(false).matchCase(true).findNext();
}

// Row number on the main tab where this barcode was scanned on this date (latest one), or 0.
function findSameDayRow(sh, last, code, date) {
  const hits = sh.getRange(2, 5, last - 1, 1).createTextFinder(String(code))
    .matchEntireCell(true).matchCase(true).findAll();
  for (let i = hits.length - 1; i >= 0; i--) {
    const r = hits[i].getRow();
    if (String(sh.getRange(r, 2).getValue()) === date) return r;
  }
  return 0;
}

// "2026-11-10 10:00:05 by Ravi, Camera 1" from a row's values
function describe(v) {
  return String(v[1]) + " " + String(v[2]).slice(0, 8) + " by " + String(v[3]) + (v[8] ? ", " + v[8] : "");
}

// All "Archive YYYY-MM" tabs, oldest month first
function archiveTabs(ss) {
  return ss.getSheets()
    .filter(function (s) { return s.getName().indexOf(ARCHIVE_PREFIX) === 0; })
    .sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
}

// Oldest earlier scan of this barcode in the archive tabs, or null. One search covers every tab.
function findInArchives(ss, code) {
  const hits = ss.createTextFinder(String(code)).matchEntireCell(true).matchCase(true).findAll()
    .filter(function (h) {
      return h.getColumn() === 5 && h.getRow() > 1 && h.getSheet().getName().indexOf(ARCHIVE_PREFIX) === 0;
    });
  if (!hits.length) return null;
  hits.sort(function (a, b) {
    const x = a.getSheet().getName(), y = b.getSheet().getName();
    return x < y ? -1 : x > y ? 1 : a.getRow() - b.getRow();
  });
  return hits[0];
}

// Rows of one tab: only the given day ("2026-11-24"), or every row when day is "".
function readRows(s, day) {
  const last = s.getLastRow();
  if (last < 2) return [];
  let from = 2, to = last;
  if (day) {
    // find just the rows of that day, then read only that block (not the whole tab)
    const hits = s.getRange(2, 2, last - 1, 1).createTextFinder(day).matchEntireCell(true).findAll();
    if (!hits.length) return [];
    from = hits[0].getRow(); to = hits[hits.length - 1].getRow();
  }
  return s.getRange(from, 1, to - from + 1, HEADERS.length).getValues()
    .filter(function (r) { return !day || String(r[1]) === day; });
}

// ---------------- monthly archive ----------------
// Moves scans older than last month from "Scans" to "Archive YYYY-MM" tabs, so the main tab
// stays small and fast. Runs by itself on the 1st of every month after setupMonthlyArchive().
// Safe to run again: rows already copied to an archive tab are not copied twice.
function archiveOldScans() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);                      // scans wait a moment while the move happens
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) { console.log("Nothing to archive."); return "Nothing to archive."; }

    // first day that stays on the main tab, e.g. "2026-10-01" when run in November
    const ym = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM").split("-");
    const first = new Date(Number(ym[0]), Number(ym[1]) - KEEP_MONTHS, 1);
    const cutoff = first.getFullYear() + "-" + String(first.getMonth() + 1).padStart(2, "0") + "-01";

    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const keep = [], byMonth = {};
    data.forEach(function (r) {
      const dt = String(r[1]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(dt) && dt < cutoff) (byMonth[dt.slice(0, 7)] = byMonth[dt.slice(0, 7)] || []).push(r);
      else keep.push(r);
    });
    const months = Object.keys(byMonth).sort();
    if (!months.length) { console.log("Nothing older than " + cutoff + "."); return "Nothing older than " + cutoff + "."; }

    // 1) copy old rows into their month's archive tab
    months.forEach(function (m) {
      const a = archiveSheet(ss, m);
      const al = a.getLastRow();
      const have = {};
      if (al > 1) a.getRange(2, 1, al - 1, 1).getValues().forEach(function (x) { have[String(x[0])] = true; });
      const rows = byMonth[m].filter(function (r) { return !have[String(r[0])]; });
      if (rows.length) a.getRange(al + 1, 1, rows.length, HEADERS.length).setValues(rows);
    });
    SpreadsheetApp.flush();                  // archive copies are saved before anything is removed

    // 2) only then shrink the main tab to the rows that stay
    if (keep.length) sh.getRange(2, 1, keep.length, HEADERS.length).setValues(keep);
    const extra = (last - 1) - keep.length;
    if (extra > 0) {
      if (keep.length) sh.deleteRows(2 + keep.length, extra);
      else {                                 // Google needs at least one row under the header
        sh.getRange(2, 1, 1, HEADERS.length).clearContent();
        if (extra > 1) sh.deleteRows(3, extra - 1);
      }
    }
    const msg = "Moved " + (data.length - keep.length) + " scans (" + months.join(", ") + ") to archive tabs. " +
                keep.length + " scans stay on " + SHEET_NAME + ".";
    console.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

function archiveSheet(ss, month) {
  const name = ARCHIVE_PREFIX + month;
  let a = ss.getSheetByName(name);
  if (!a) {
    a = ss.insertSheet(name, ss.getSheets().length);   // new tab goes at the end
    a.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    a.setFrozenRows(1);
    a.getRange("A:E").setNumberFormat("@");
    a.getRange("G:G").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } else if (a.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
    a.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");   // older tab: add new headers
  }
  return a;
}

// Run this ONCE from the Apps Script editor: it makes archiveOldScans run on the 1st of every month.
function setupMonthlyArchive() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "archiveOldScans") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("archiveOldScans").timeBased().onMonthDay(1).atHour(3).create();
  console.log("Monthly archive is on: runs on the 1st of each month around 3 AM.");
}

function doGet() { return reply({ ok: true, message: "Scan receiver is running" }); }

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
