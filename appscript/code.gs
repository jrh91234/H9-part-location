/**
 * H9 Part Location — เก็บการเปลี่ยนคู่ Part ↔ Location (เมนูตั้งค่าในหน้าเว็บ)
 * Backend: Google Apps Script + Google Sheets
 *
 * ข้อมูลหลักมาจากไฟล์ Excel (js/data.js ในหน้าเว็บ) · ชีตนี้เก็บเฉพาะ "ช่องที่ถูกเปลี่ยน" ทับข้อมูล Excel
 * หน้าเว็บ (GitHub Pages) ส่ง POST { action, args: [...] } แบบ text/plain → ตอบ { ok, data } หรือ { ok:false, error }
 */

// ======================= ค่าคงที่ =======================
const ASSIGN_SHEET = 'Assignments';
const ASSIGN_HEADERS = ['location', 'part_code', 'updated_by', 'updated_at', 'note'];
const LOG_SHEET = 'Log';
const LOG_HEADERS = ['at', 'by', 'action', 'location', 'old_code', 'new_code', 'note'];
const NEW_SHEET = 'NewParts'; // Part ที่สแกนเจอแต่ไม่มีในระบบ (แจ้งจากหน้าเว็บ) — เก็บไว้ทำ Location
const NEW_HEADERS = ['reported_at', 'reported_by', 'part_code', 'name', 'found_at', 'note', 'raw_scan'];
const TZ = 'Asia/Bangkok';
const DEFAULT_PIN = '1234'; // เปลี่ยนได้ที่ Project Settings → Script Properties → ADMIN_PIN
const LOC_RE = /^H9-L\d+-\d{2}$/;
const MAX_CHANGES = 10;

// ======================= Web App =======================
const API_ACTIONS = { ping: ping, checkPin: checkPin, getAssignments: getAssignments, saveAssignments: saveAssignments,
  revertAssignment: revertAssignment, getLog: getLog,
  reportNewPart: reportNewPart, getNewParts: getNewParts, deleteNewPart: deleteNewPart };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_ACTIONS[req.action];
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง: ' + req.action);
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** เปิด /exec ในเบราว์เซอร์เพื่อเช็คว่า API ทำงาน */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'H9 Part Location API', time: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================= เมนูใน Google Sheet =======================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('H9 Location')
    .addItem('ตั้งค่าเริ่มต้น (สร้างชีต)', 'setupFromMenu')
    .addToUi();
}

function setupFromMenu() {
  assignSheet_();
  logSheet_();
  newSheet_();
  SpreadsheetApp.getUi().alert('พร้อมใช้งาน: ชีต "' + ASSIGN_SHEET + '", "' + LOG_SHEET + '" และ "' + NEW_SHEET + '"');
}

// ======================= ชีต =======================
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Apps Script ไม่ได้ผูกกับ Google Sheet — ใส่ Script Properties: SPREADSHEET_ID = ID ของชีต H9 Part Location');
  return ss;
}

function sheetWith_(name, headers) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@'); // เก็บเป็นข้อความ (กันเลข 0 นำหน้าหาย)
  }
  return sh;
}
function assignSheet_() { return sheetWith_(ASSIGN_SHEET, ASSIGN_HEADERS); }
function logSheet_() { return sheetWith_(LOG_SHEET, LOG_HEADERS); }
function newSheet_() { return sheetWith_(NEW_SHEET, NEW_HEADERS); }

function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
// Sheets อาจแปลงข้อความเวลาเป็น Date เอง → อ่านกลับเป็นรูปแบบเดียวกันเสมอ (เรียงลำดับได้)
function ts_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss") : String(v); }

function readAssign_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, ASSIGN_HEADERS.length).getValues().map(function (r, i) {
    return { row: i + 2, loc: String(r[0]).trim(), code: String(r[1]).trim(), by: String(r[2]), at: ts_(r[3]), note: String(r[4]) };
  }).filter(function (r) { return r.loc; });
}
function publicRows_(rows) {
  return rows.map(function (r) { return { loc: r.loc, code: r.code, by: r.by, at: r.at, note: r.note }; });
}

// ======================= ตรวจรหัส =======================
function pin_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_PIN; }
function requirePin_(pin) {
  if (String(pin || '') !== pin_()) throw new Error('รหัสตั้งค่าไม่ถูกต้อง');
}
function cleanBy_(by) {
  by = String(by || '').trim().slice(0, 40);
  if (by.length < 2) throw new Error('กรุณาใส่ชื่อผู้แก้ไข');
  return by;
}

// ======================= คำสั่ง API =======================
function ping() { return { time: now_() }; }

function checkPin(pin) { return String(pin || '') === pin_(); }

/** ทุกช่องที่ถูกเปลี่ยนทับข้อมูล Excel: [{ loc, code ('' = ช่องว่าง), by, at, note }] */
function getAssignments() { return publicRows_(readAssign_(assignSheet_())); }

/**
 * บันทึกการเปลี่ยน (ทำทีละชุด เช่น สลับ 2 ช่อง)
 * changes: [{ loc, code, old }]  code '' = ทำให้ช่องว่าง · old = Part เดิมของช่อง (เก็บใน Log)
 */
function saveAssignments(pin, by, changes, note) {
  requirePin_(pin);
  by = cleanBy_(by);
  note = String(note || '').slice(0, 200);
  if (!Array.isArray(changes) || !changes.length) throw new Error('ไม่มีรายการเปลี่ยน');
  if (changes.length > MAX_CHANGES) throw new Error('เปลี่ยนได้ครั้งละไม่เกิน ' + MAX_CHANGES + ' ช่อง');
  changes.forEach(function (c) {
    if (!LOC_RE.test(String(c.loc || ''))) throw new Error('รูปแบบ Location ไม่ถูกต้อง: ' + c.loc);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = assignSheet_(), log = logSheet_(), at = now_();
    const rows = readAssign_(sh), byLoc = {};
    rows.forEach(function (r) { byLoc[r.loc] = r; });
    changes.forEach(function (c) {
      const code = String(c.code || '').trim().toUpperCase().slice(0, 40);
      const vals = [[c.loc, code, by, at, note]];
      if (byLoc[c.loc]) sh.getRange(byLoc[c.loc].row, 1, 1, ASSIGN_HEADERS.length).setValues(vals);
      else { sh.appendRow(vals[0]); byLoc[c.loc] = { row: sh.getLastRow() }; }
      log.appendRow([at, by, code ? 'assign' : 'clear', c.loc, String(c.old || ''), code, note]);
    });
    return publicRows_(readAssign_(sh));
  } finally {
    lock.releaseLock();
  }
}

/** คืนค่าช่องกลับไปตามข้อมูล Excel (ลบแถวที่เปลี่ยนไว้) */
function revertAssignment(pin, by, loc, oldCode, newCode) {
  requirePin_(pin);
  by = cleanBy_(by);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = assignSheet_();
    const rows = readAssign_(sh).filter(function (r) { return r.loc === loc; });
    rows.sort(function (a, b) { return b.row - a.row; }).forEach(function (r) { sh.deleteRow(r.row); });
    logSheet_().appendRow([now_(), by, 'revert', loc, String(oldCode || ''), String(newCode || ''), 'คืนค่าตาม Excel']);
    return publicRows_(readAssign_(sh));
  } finally {
    lock.releaseLock();
  }
}

/** ประวัติล่าสุด (ใหม่ → เก่า) */
function getLog(limit) {
  const sh = logSheet_(), n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  limit = Math.min(Math.max(Number(limit) || 50, 1), 500);
  const start = Math.max(2, sh.getLastRow() - limit + 1);
  return sh.getRange(start, 1, sh.getLastRow() - start + 1, LOG_HEADERS.length).getValues().reverse().map(function (r) {
    return { at: ts_(r[0]), by: String(r[1]), action: String(r[2]), loc: String(r[3]), old: String(r[4]), code: String(r[5]), note: String(r[6]) };
  });
}

// ======================= Part ใหม่ที่ไม่มีในระบบ =======================
function readNew_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, NEW_HEADERS.length).getValues().map(function (r, i) {
    return { row: i + 2, at: ts_(r[0]), by: String(r[1]), code: String(r[2]).trim(), name: String(r[3]), found: String(r[4]), note: String(r[5]), raw: String(r[6]) };
  }).filter(function (r) { return r.code; });
}
function publicNew_(rows) {
  return rows.map(function (r) { return { at: r.at, by: r.by, code: r.code, name: r.name, found: r.found, note: r.note, raw: r.raw }; });
}

/** รายการ Part ใหม่ทั้งหมด (เก่า → ใหม่) */
function getNewParts() { return publicNew_(readNew_(newSheet_())); }

/**
 * แจ้ง Part ที่สแกนเจอแต่ไม่มีในระบบ (ทุกคนแจ้งได้ ไม่ต้องใช้รหัส)
 * item: { code, name, found, note, raw } · รหัสซ้ำ → อัปเดตแถวเดิม (ไม่เพิ่มซ้ำ)
 */
function reportNewPart(by, item) {
  by = cleanBy_(by);
  item = item || {};
  const code = String(item.code || '').trim().toUpperCase().slice(0, 40);
  if (code.length < 3) throw new Error('กรุณาใส่รหัส Part อย่างน้อย 3 ตัวอักษร');
  const found = LOC_RE.test(String(item.found || '')) ? String(item.found) : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = newSheet_(), at = now_();
    const vals = [at, by, code, String(item.name || '').slice(0, 120), found, String(item.note || '').slice(0, 200), String(item.raw || '').slice(0, 300)];
    const old = readNew_(sh).filter(function (r) { return r.code === code; })[0];
    if (old) {
      if (!vals[3]) vals[3] = old.name;
      if (!vals[4]) vals[4] = old.found;
      sh.getRange(old.row, 1, 1, NEW_HEADERS.length).setValues([vals]);
    } else sh.appendRow(vals);
    logSheet_().appendRow([at, by, 'new_part', found, '', code, String(item.name || '')]);
    return publicNew_(readNew_(sh));
  } finally {
    lock.releaseLock();
  }
}

/** ลบรายการ Part ใหม่ (ผู้ดูแล — ต้องใช้รหัส) เช่น เพิ่มลงไฟล์ Excel แล้ว / แจ้งผิด */
function deleteNewPart(pin, by, code) {
  requirePin_(pin);
  by = cleanBy_(by);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = newSheet_();
    readNew_(sh).filter(function (r) { return r.code === code; })
      .sort(function (a, b) { return b.row - a.row; }).forEach(function (r) { sh.deleteRow(r.row); });
    logSheet_().appendRow([now_(), by, 'delete_new_part', '', code, '', '']);
    return publicNew_(readNew_(sh));
  } finally {
    lock.releaseLock();
  }
}
