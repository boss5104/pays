/**
 * Smart Expense Tracker — ตัวเชื่อม Google Sheets (Google Apps Script)
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *  1. สร้าง Google Sheets ใหม่ (https://sheets.new) — ไฟล์จะอยู่ใน Google Drive ของคุณ
 *  2. เมนู "ส่วนขยาย" → "Apps Script" → ลบโค้ดเดิมทิ้ง แล้ววางโค้ดทั้งหมดนี้ → กดบันทึก
 *  3. กด "ทำให้ใช้งานได้" → "การทำให้ใช้งานได้รายการใหม่" → เลือกประเภท "เว็บแอป"
 *       - เรียกใช้ในฐานะ: ฉัน (Me)
 *       - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน (Anyone)
 *     แล้วกด "ทำให้ใช้งานได้" และอนุญาตสิทธิ์
 *     (ถ้าขึ้นว่า "Google ยังไม่ได้ยืนยันแอปนี้" ให้กด "ขั้นสูง" → "ไปที่ ... (ไม่ปลอดภัย)" — เป็นสคริปต์ของคุณเอง)
 *  4. คัดลอก "URL ของเว็บแอป" (ลงท้ายด้วย /exec) ไปวางในแอป Smart Expense
 *
 * ถ้าแก้โค้ดนี้ภายหลัง ต้อง "ทำให้ใช้งานได้" → "จัดการการทำให้ใช้งานได้" → แก้ไข → เวอร์ชันใหม่ ทุกครั้ง
 */

// (ไม่บังคับ) ตั้งรหัสลับ เช่น 'my-secret-123' แล้วกรอกรหัสเดียวกันในแอป
// เพื่อให้คนที่รู้แค่ URL อ่าน/เขียนข้อมูลไม่ได้
const SECRET = '';

const FIXED_SHEET = 'ค่าใช้จ่ายประจำ';
const INCOME_SHEET = 'รายได้';
const META_SHEET = '_meta';
const FIXED_HEADERS = ['เดือน (ปปปป-ดด)', 'รายการ', 'จำนวนเงิน (บาท)', 'จ่ายแล้ว', 'id'];
const INCOME_HEADERS = ['รายการ', 'จำนวนเงิน (บาท)', 'id'];

function doGet(e) {
  return respond_(handle_((e && e.parameter) || {}));
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return respond_({ ok: false, error: 'bad json' });
  }
  return respond_(handle_(Object.assign({}, (e && e.parameter) || {}, body)));
}

function handle_(p) {
  try {
    if (SECRET && String(p.key || '') !== SECRET) return { ok: false, error: 'unauthorized' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (p.action || 'load') {
      case 'load':
        return { ok: true, data: readAll_(ss), sheetUrl: ss.getUrl(), sheetName: ss.getName() };
      case 'save': {
        const lock = LockService.getScriptLock();
        lock.waitLock(20000);
        try {
          const at = writeAll_(ss, p.data || {});
          return { ok: true, updatedAt: at, sheetUrl: ss.getUrl(), sheetName: ss.getName() };
        } finally {
          lock.releaseLock();
        }
      }
      default:
        return { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ถ้าแก้ข้อมูลในชีตด้วยมือ ให้บันทึกเวลาแก้ไข เพื่อให้แอปดึงข้อมูลใหม่ไปใช้
function onEdit(e) {
  const name = e && e.range && e.range.getSheet().getName();
  if (name === FIXED_SHEET || name === INCOME_SHEET) setUpdatedAt_(e.source, Date.now());
}

// (ไม่บังคับ) กดรันฟังก์ชันนี้ในหน้า Apps Script เพื่อสร้างแผ่นงานและอนุญาตสิทธิ์ล่วงหน้า
function setup() {
  readAll_(SpreadsheetApp.getActiveSpreadsheet());
}

// ---------------- อ่าน / เขียนข้อมูล ----------------

function readAll_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const fixedSh = getSheet_(ss, FIXED_SHEET, FIXED_HEADERS);
  const incomeSh = getSheet_(ss, INCOME_SHEET, INCOME_HEADERS);

  const fixedExpenses = readRows_(fixedSh, FIXED_HEADERS.length, 4)
    .map(function (r) {
      return {
        id: r.id,
        month: toMonth_(r.v[0], tz),
        name: String(r.v[1]).trim(),
        amount: toNum_(r.v[2]),
        paid: r.v[3] === true || String(r.v[3]).toUpperCase() === 'TRUE'
      };
    })
    .filter(function (x) { return x.name && x.month; });

  const customIncomes = readRows_(incomeSh, INCOME_HEADERS.length, 2)
    .map(function (r) {
      return { id: r.id, name: String(r.v[0]).trim(), amount: toNum_(r.v[1]) };
    })
    .filter(function (x) { return x.name; });

  return { fixedExpenses: fixedExpenses, customIncomes: customIncomes, updatedAt: getUpdatedAt_(ss) };
}

function writeAll_(ss, data) {
  const fixed = (data.fixedExpenses || []).map(function (f) {
    return [String(f.month || ''), String(f.name || ''), toNum_(f.amount), !!f.paid, String(f.id || Utilities.getUuid())];
  });
  fixed.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  const incomes = (data.customIncomes || []).map(function (i) {
    return [String(i.name || ''), toNum_(i.amount), String(i.id || Utilities.getUuid())];
  });

  writeTable_(getSheet_(ss, FIXED_SHEET, FIXED_HEADERS), FIXED_HEADERS.length, fixed, { textCols: [1, 2, 5], moneyCol: 3, checkCol: 4 });
  writeTable_(getSheet_(ss, INCOME_SHEET, INCOME_HEADERS), INCOME_HEADERS.length, incomes, { textCols: [1, 3], moneyCol: 2 });

  const at = toNum_(data.updatedAt) || Date.now();
  setUpdatedAt_(ss, at);
  return at;
}

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#0B1020')
      .setFontColor('#E8C46A');
    sh.setFrozenRows(1);
  }
  return sh;
}

// อ่านแถวข้อมูล — แถวที่เพิ่มเองในชีตและยังไม่มี id จะถูกสร้าง id ให้อัตโนมัติ
function readRows_(sh, width, idCol) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, width).getValues();
  const out = [];
  let changed = false;
  values.forEach(function (v) {
    if (v.every(function (c) { return c === '' || c === false; })) return;
    if (!v[idCol]) { v[idCol] = Utilities.getUuid(); changed = true; }
    out.push({ v: v, id: String(v[idCol]) });
  });
  if (changed) {
    sh.getRange(2, idCol + 1, values.length, 1).setNumberFormat('@')
      .setValues(values.map(function (v) { return [v[idCol]]; }));
  }
  return out;
}

function writeTable_(sh, width, rows, opt) {
  const last = sh.getLastRow();
  if (last > 1) {
    const old = sh.getRange(2, 1, last - 1, width);
    old.clearContent();
    old.clearDataValidations();
  }
  if (!rows.length) return;
  const n = rows.length;
  (opt.textCols || []).forEach(function (c) { sh.getRange(2, c, n, 1).setNumberFormat('@'); });
  if (opt.moneyCol) sh.getRange(2, opt.moneyCol, n, 1).setNumberFormat('#,##0.00');
  if (opt.checkCol) sh.getRange(2, opt.checkCol, n, 1).insertCheckboxes(); // ต้องใส่ก่อน setValues
  sh.getRange(2, 1, n, width).setValues(rows);
  sh.autoResizeColumns(1, width);
}

function getUpdatedAt_(ss) {
  const sh = ss.getSheetByName(META_SHEET);
  return sh ? toNum_(sh.getRange('B1').getValue()) : 0;
}

function setUpdatedAt_(ss, at) {
  let sh = ss.getSheetByName(META_SHEET);
  if (!sh) {
    sh = ss.insertSheet(META_SHEET);
    sh.getRange('A1').setValue('updatedAt');
    sh.hideSheet();
  }
  sh.getRange('B1').setValue(at);
}

function toMonth_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM');
  const m = String(v).trim().match(/^(\d{4})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) : '';
}

function toNum_(v) {
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,฿\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
