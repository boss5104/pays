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
const LOAN_SHEET = 'ผ่อนชำระ';
const LOAN_PAY_SHEET = 'งวดผ่อน';
const SAV_ACC_SHEET = 'บัญชีเงินฝาก';
const SAV_TX_SHEET = 'รายการฝากเงิน';
const META_SHEET = '_meta';
const FIXED_HEADERS = ['เดือน (ปปปป-ดด)', 'รายการ', 'จำนวนเงิน (บาท)', 'จ่ายแล้ว', 'id'];
const INCOME_HEADERS = ['รายการ', 'จำนวนเงิน (บาท)', 'id', 'ใช้แล้ว (เดือน)'];
const LOAN_HEADERS = ['รายการผ่อน', 'ไอคอน', 'จำนวนงวด (เดือน)', 'ค่างวดปกติ (บาท)', 'งวดแรก (ปปปป-ดด)', 'ครบกำหนดทุกวันที่', 'หมายเหตุ', 'id', 'ค่าใช้จ่ายเพิ่มต่องวด (ค่าอะไร=บาท; …)', 'จ่ายจริงต่องวด (บาท)'];
const LOAN_PAY_HEADERS = ['รายการผ่อน', 'งวดที่', 'ครบกำหนด', 'จ่ายแล้ว', 'วันที่จ่าย', 'หมายเหตุ', 'id รายการผ่อน', 'จ่ายจริง (บาท)'];
const SAV_ACC_HEADERS = ['ฝากเงินอะไร (บัญชี / เป้าหมาย)', 'ไอคอน', 'เป้าหมาย (บาท)', 'หมายเหตุ', 'id', 'ประเภทเงินฝาก', 'ฝากต่อเดือน (บาท)'];
const SAV_TX_HEADERS = ['วันที่', 'ประเภท', 'บัญชี', 'จำนวนเงิน (บาท)', 'หมายเหตุ', 'id', 'id บัญชี'];
const DATA_SHEETS = [FIXED_SHEET, INCOME_SHEET, LOAN_SHEET, LOAN_PAY_SHEET, SAV_ACC_SHEET, SAV_TX_SHEET];
// แอปใช้ตรวจว่าเป็นโค้ดเวอร์ชันใหม่
//   v2: ช่องติ๊ก "ใช้แล้ว" ของรายรับ + อัปเดตอัตโนมัติ · v3: ผ่อนชำระ + เงินฝาก
//   v4: ผ่อน — ค่าใช้จ่ายเพิ่ม/จ่ายจริง · เงินฝาก — ประเภท/ฝากต่อเดือน
const VERSION = 4;

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
      case 'ping': // เบา ๆ: แอปเรียกทุก 15 วินาทีเพื่อตรวจว่ามีการแก้ไขจากเครื่องอื่นหรือไม่
        return { ok: true, updatedAt: getUpdatedAt_(ss), version: VERSION };
      case 'load':
        return { ok: true, data: readAll_(ss), sheetUrl: ss.getUrl(), sheetName: ss.getName(), version: VERSION };
      case 'save': {
        const lock = LockService.getScriptLock();
        lock.waitLock(20000);
        try {
          const at = writeAll_(ss, p.data || {});
          return { ok: true, updatedAt: at, sheetUrl: ss.getUrl(), sheetName: ss.getName(), version: VERSION };
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
  if (DATA_SHEETS.indexOf(name) >= 0) setUpdatedAt_(e.source, Date.now());
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
      return { id: r.id, name: String(r.v[0]).trim(), amount: toNum_(r.v[1]), used: toMonthList_(r.v[3], tz) };
    })
    .filter(function (x) { return x.name; });

  // ---- ผ่อนชำระ: ข้อมูลรายการอยู่แผ่นงาน "ผ่อนชำระ" และสถานะแต่ละงวดอยู่แผ่นงาน "งวดผ่อน" ----
  const loans = readRows_(getSheet_(ss, LOAN_SHEET, LOAN_HEADERS), LOAN_HEADERS.length, 7)
    .map(function (r) {
      return {
        id: r.id, name: String(r.v[0]).trim(), icon: String(r.v[1] || '').trim(), months: toNum_(r.v[2]),
        amount: toNum_(r.v[3]), start: toMonth_(r.v[4], tz), dueDay: toNum_(r.v[5]), note: String(r.v[6] || ''),
        extras: parseExtras_(r.v[8]), pays: []
      };
    })
    .filter(function (x) { return x.name; });
  const loanById = {}, loanByName = {};
  loans.forEach(function (l) { loanById[l.id] = l; if (!loanByName[l.name]) loanByName[l.name] = l; });
  readValues_(getSheet_(ss, LOAN_PAY_SHEET, LOAN_PAY_HEADERS), LOAN_PAY_HEADERS.length).forEach(function (v) {
    const l = loanById[String(v[6]).trim()] || loanByName[String(v[0]).trim()];
    const n = Math.round(toNum_(v[1]));
    if (!l || n < 1) return;
    const paid = v[3] === true || String(v[3]).toUpperCase() === 'TRUE';
    const date = toDate_(v[4], tz), note = String(v[5] || '').trim();
    // จ่ายจริง: ว่าง หรือเท่ากับยอดปกติของรายการ (ค่างวด + ค่าใช้จ่ายเพิ่ม) = ใช้ยอดปกติ
    let actual = null;
    if (String(v[7]).trim() !== '') {
      const a = toNum_(v[7]);
      if (Math.abs(a - loanActual_(l)) >= 0.005) actual = a;
    }
    if (paid || date || note || actual !== null) l.pays.push({ n: n, paid: paid, date: date, note: note, actual: actual });
  });

  // ---- เงินฝาก ----
  const savAccounts = readRows_(getSheet_(ss, SAV_ACC_SHEET, SAV_ACC_HEADERS), SAV_ACC_HEADERS.length, 4)
    .map(function (r) {
      return {
        id: r.id, name: String(r.v[0]).trim(), icon: String(r.v[1] || '').trim(), goal: toNum_(r.v[2]), note: String(r.v[3] || ''),
        kind: String(r.v[5] || '').trim(), monthly: toNum_(r.v[6])
      };
    })
    .filter(function (x) { return x.name; });
  const accByName = {};
  savAccounts.forEach(function (a) { if (!accByName[a.name]) accByName[a.name] = a.id; });
  const savTx = readRows_(getSheet_(ss, SAV_TX_SHEET, SAV_TX_HEADERS), SAV_TX_HEADERS.length, 5)
    .map(function (r) {
      const type = String(r.v[1]).trim();
      return {
        id: r.id, date: toDate_(r.v[0], tz), type: (type === 'ถอน' || type.toLowerCase() === 'out') ? 'out' : 'in',
        acc: String(r.v[6] || '').trim() || accByName[String(r.v[2]).trim()] || '',
        amount: Math.abs(toNum_(r.v[3])), note: String(r.v[4] || '')
      };
    })
    .filter(function (x) { return x.amount && x.date; });

  return {
    fixedExpenses: fixedExpenses, customIncomes: customIncomes, loans: loans,
    savAccounts: savAccounts, savTx: savTx, updatedAt: getUpdatedAt_(ss)
  };
}

function writeAll_(ss, data) {
  const fixed = (data.fixedExpenses || []).map(function (f) {
    return [String(f.month || ''), String(f.name || ''), toNum_(f.amount), !!f.paid, String(f.id || Utilities.getUuid())];
  });
  fixed.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  const incomes = (data.customIncomes || []).map(function (i) {
    const used = (Array.isArray(i.used) ? i.used : []).map(String).filter(function (m) { return /^\d{4}-\d{2}$/.test(m); });
    return [String(i.name || ''), toNum_(i.amount), String(i.id || Utilities.getUuid()), used.join(', ')];
  });

  writeTable_(getSheet_(ss, FIXED_SHEET, FIXED_HEADERS), FIXED_HEADERS.length, fixed, { textCols: [1, 2, 5], moneyCol: 3, checkCol: 4 });
  writeTable_(getSheet_(ss, INCOME_SHEET, INCOME_HEADERS), INCOME_HEADERS.length, incomes, { textCols: [1, 3, 4], moneyCol: 2 });

  // แอปเวอร์ชันเก่าไม่ได้ส่งข้อมูลผ่อนชำระ/เงินฝากมา → ไม่แตะแผ่นงานเหล่านั้น (กันข้อมูลหาย)
  if (Array.isArray(data.loans)) {
    const loanRows = [], payRows = [];
    data.loans.forEach(function (l) {
      const id = String(l.id || Utilities.getUuid());
      const months = Math.min(600, Math.max(1, Math.round(toNum_(l.months)) || 1));
      const name = String(l.name || ''), start = String(l.start || ''), dueDay = Math.round(toNum_(l.dueDay)) || 1;
      const extras = parseExtras_(l.extras);
      const perPeriod = loanActual_({ amount: toNum_(l.amount), extras: extras });
      loanRows.push([name, String(l.icon || ''), months, toNum_(l.amount), start, dueDay, String(l.note || ''), id, formatExtras_(extras), perPeriod]);
      const pays = {};
      (Array.isArray(l.pays) ? l.pays : []).forEach(function (p) { pays[Math.round(toNum_(p.n))] = p; });
      for (let n = 1; n <= months; n++) {
        const p = pays[n] || {};
        const hasActual = p.actual !== null && p.actual !== undefined && String(p.actual) !== '';
        // งวดที่จ่ายแล้วแสดงยอดจ่ายจริงเสมอ (อ่านกลับ: ถ้าเท่ากับยอดปกติ จะถือว่าไม่ได้แก้)
        const actual = hasActual ? toNum_(p.actual) : (p.paid ? perPeriod : '');
        payRows.push([name, n, dueDate_(start, n, dueDay), !!p.paid, String(p.date || ''), String(p.note || ''), id, actual]);
      }
    });
    writeTable_(getSheet_(ss, LOAN_SHEET, LOAN_HEADERS), LOAN_HEADERS.length, loanRows, { textCols: [1, 2, 5, 7, 8, 9], moneyCol: [4, 10] });
    writeTable_(getSheet_(ss, LOAN_PAY_SHEET, LOAN_PAY_HEADERS), LOAN_PAY_HEADERS.length, payRows, { textCols: [1, 3, 5, 6, 7], checkCol: 4, moneyCol: 8 });
  }
  if (Array.isArray(data.savAccounts) && Array.isArray(data.savTx)) {
    const accName = {};
    const accRows = data.savAccounts.map(function (a) {
      const id = String(a.id || Utilities.getUuid());
      accName[id] = String(a.name || '');
      return [String(a.name || ''), String(a.icon || ''), toNum_(a.goal), String(a.note || ''), id, String(a.kind || ''), toNum_(a.monthly)];
    });
    const txRows = data.savTx.map(function (t) {
      const acc = String(t.acc || '');
      return [String(t.date || ''), t.type === 'out' ? 'ถอน' : 'ฝาก', accName[acc] || '', Math.abs(toNum_(t.amount)), String(t.note || ''), String(t.id || Utilities.getUuid()), acc];
    });
    txRows.sort(function (a, b) { return a[0] < b[0] ? 1 : a[0] > b[0] ? -1 : 0; }); // ใหม่สุดอยู่บน
    writeTable_(getSheet_(ss, SAV_ACC_SHEET, SAV_ACC_HEADERS), SAV_ACC_HEADERS.length, accRows, { textCols: [1, 2, 4, 5, 6], moneyCol: [3, 7] });
    writeTable_(getSheet_(ss, SAV_TX_SHEET, SAV_TX_HEADERS), SAV_TX_HEADERS.length, txRows, { textCols: [1, 2, 3, 5, 6, 7], moneyCol: 4 });
  }

  const at = toNum_(data.updatedAt) || Date.now();
  setUpdatedAt_(ss, at);
  return at;
}

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // สร้างหัวตาราง (หรือเติมหัวคอลัมน์ใหม่ให้ชีตที่สร้างด้วยโค้ดเวอร์ชันเก่า)
  if (headers && (sh.getLastRow() === 0 || sh.getRange(1, headers.length).getValue() === '')) {
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

// อ่านแถวข้อมูลแบบไม่สร้าง id (ใช้กับแผ่นงานที่อ้างอิง id ของแผ่นงานอื่น)
function readValues_(sh, width) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, width).getValues().filter(function (v) {
    return !v.every(function (c) { return c === '' || c === false; });
  });
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
  [].concat(opt.moneyCol || []).forEach(function (c) { sh.getRange(2, c, n, 1).setNumberFormat('#,##0.00'); });
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

function toDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

// วันครบกำหนดของงวดที่ n (ถ้าเดือนนั้นมีวันไม่ถึง จะใช้วันสุดท้ายของเดือน)
function dueDate_(start, n, dueDay) {
  const m = String(start).match(/^(\d{4})-(\d{2})/);
  if (!m) return '';
  const t = Number(m[1]) * 12 + Number(m[2]) - 1 + (n - 1);
  const y = Math.floor(t / 12), mo = t % 12 + 1;
  const d = Math.min(dueDay, new Date(y, mo, 0).getDate());
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

// "2026-10, 2026-11" → ['2026-10', '2026-11']
function toMonthList_(v, tz) {
  if (v instanceof Date) return [Utilities.formatDate(v, tz, 'yyyy-MM')];
  return String(v || '').split(/[,\s]+/)
    .map(function (s) { return toMonth_(s, tz); })
    .filter(function (s, i, a) { return s && a.indexOf(s) === i; });
}

// "ดอกเบี้ย=150; ประกัน=50" หรือ [{ name, amount }] → [{ name, amount }]
function parseExtras_(v) {
  let arr = [];
  if (Array.isArray(v)) arr = v;
  else String(v || '').split(/[;\n]+/).forEach(function (x) {
    const i = x.lastIndexOf('=');
    if (i >= 0) arr.push({ name: x.slice(0, i), amount: x.slice(i + 1) });
  });
  return arr.map(function (e) {
    return { name: String(e.name || '').replace(/[=;\n\r]/g, ' ').replace(/\s+/g, ' ').trim() || 'ค่าใช้จ่ายเพิ่ม', amount: toNum_(e.amount) };
  }).filter(function (e) { return e.amount; });
}

function formatExtras_(extras) {
  return extras.map(function (e) { return e.name + '=' + e.amount; }).join('; ');
}

// จ่ายจริงต่องวด = ค่างวดปกติ + ค่าใช้จ่ายเพิ่ม
function loanActual_(l) {
  return toNum_(l.amount) + (l.extras || []).reduce(function (s, e) { return s + toNum_(e.amount); }, 0);
}

function toNum_(v) {
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,฿\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
