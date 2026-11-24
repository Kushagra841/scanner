// Google Apps Script: receives scans from the packer page and adds each one as a row.
// Paste this into Extensions > Apps Script inside the Google Sheet that should hold the scans.

const SECRET = "shantam@143";   // must be the same as SECRET in the page (index.html)
const ADMIN_KEY = "set-your-own-admin-password";   // ONLY for downloads. Never put this in the page.
const SHEET_NAME = "Scans";
const HEADERS = ["Scan ID", "Date", "Time", "Packer", "Barcode", "Repeat scan", "Received at", "Status", "Camera / Table"];
const STATUS_TEXT = "Forward";            // goes in the Status column of every row
const MAX_BATCH = 50;                     // most scans accepted in one request
const ARCHIVE_PREFIX = "Archive ";        // old months move to tabs like "Archive 2026-09"
const KEEP_MONTHS = 2;                    // months kept on the main tab: this month + last month

// Adds the Status column to an older sheet and fills it for rows already there
function ensureStatus(sh) {
  if (sh.getLastRow() < 1) return;
  if (sh.getRange(1, 9).getValue() !== "Camera / Table") sh.getRange(1, 9).setValue("Camera / Table").setFontWeight("bold");
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
    const hasArchive = archiveTabs(ss).length > 0;

    items.forEach(function (it) {
      const id = String(it.id), code = String(it.code);

      // same scan sent twice (for example after a bad connection): ignore the second copy
      if (idsInBatch[id] || (last > 1 && findExact(sh.getRange(2, 1, last - 1, 1), id))) {
        results.push({ id: id, ok: true, repeat: false, duplicateRequest: true });
        return;
      }

      // has this barcode been scanned before? Google's own search looks through the
      // whole Barcode column at once, so it stays fast however many rows there are.
      // Archive tabs hold the oldest scans, so they are checked first.
      let firstSeen = "";
      let hit = hasArchive ? findInArchives(ss, code) : null;
      if (!hit && last > 1) hit = findExact(sh.getRange(2, 5, last - 1, 1), code);   // first (oldest) match
      if (hit) {
        const r = hit.getSheet().getRange(hit.getRow(), 2, 1, 3).getValues()[0];   // Date, Time, Packer
        firstSeen = r[0] + " " + r[1] + " by " + r[2];
      }
      if (!firstSeen && codesInBatch[code]) firstSeen = codesInBatch[code];

      const row = [id, String(it.date), String(it.time), String(it.packer), code,
                   firstSeen ? "YES" : "", now, STATUS_TEXT, String(it.table || "")];
      newRows.push(row);
      idsInBatch[id] = true;
      if (!codesInBatch[code]) codesInBatch[code] = row[1] + " " + row[2] + " by " + row[3];
      results.push({ id: id, ok: true, repeat: !!firstSeen, firstSeen: firstSeen });
    });

    // all new rows are written in one go
    if (newRows.length) sh.getRange(last + 1, 1, newRows.length, HEADERS.length).setValues(newRows);

    if (batch) return reply({ ok: true, results: results });
    const one = results[0];                        // reply in the old shape for a single scan
    return reply(one.duplicateRequest ? { ok: true, repeat: false, duplicateRequest: true }
                                      : { ok: true, repeat: one.repeat, firstSeen: one.firstSeen });
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
