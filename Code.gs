// ============================================================
// MEGA FARM CAMBODIA — Labor Daily App · Google Apps Script
// Container-bound script (Extensions → Apps Script จาก Sheet "MGF Labor Daily Report")
// Deploy: Execute as Me · Access: Anyone
// ============================================================
// โครงสร้าง Sheet:
//   LaborLog     : ID, BatchID, Timestamp, EditedAt, Date, Plot, Contractor,
//                  NumWorkers, Hours, RatePerHourUSD, OTHours, OTMultiplier,
//                  Supervisor, TotalUSD, EnteredBy, Status (ACTIVE/EDITED)
//   AppConfig    : Key, Value  → PIN, LINE_TOKEN, LINE_TARGET, OT_MULTIPLIER
//   Plots        : รายชื่อแปลง (คอลัมน์ A ตั้งแต่แถว 2) — แก้ได้ตรงในชีต
//   Contractors  : รายชื่อผู้รับเหมา (คอลัมน์ A ตั้งแต่แถว 2) — แก้ได้ตรงในชีต
//
// หมายเหตุ: Date เก็บเป็น "@" (Plain Text) รูปแบบ "YYYY-MM-DD" เท่านั้น — ห้ามให้
// Sheet ตีความเป็น Date object เด็ดขาด เพราะ locale ของ Google Sheet (US) จะสลับ
// วัน/เดือนสำหรับวันที่ ≤ 12 (บั๊กเดียวกับที่เจอใน Rainfall Record มาก่อน)
// ============================================================

const LOG_SHEET_NAME    = 'LaborLog';
const CFG_SHEET_NAME    = 'AppConfig';
const PLOTS_SHEET_NAME  = 'Plots';
const CONTR_SHEET_NAME  = 'Contractors';
const DEFAULT_PIN       = '1234';
const DEFAULT_OT_MULT   = 1.5;
const TZ                = 'Asia/Phnom_Penh';

const LOG_HEADERS = [
  'ID','BatchID','Timestamp','EditedAt','Date','Plot','Contractor',
  'NumWorkers','Hours','RatePerHourUSD','OTHours','OTMultiplier',
  'Supervisor','TotalUSD','EnteredBy','Status','Note'
];

// ── Response helper ──────────────────────────────────────────
function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ──────────────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter || {};
  try {
    switch (p.action) {
      case 'getPlots':       return out({ plots: getListValues(PLOTS_SHEET_NAME) });
      case 'getContractors': return out({ contractors: getListValues(CONTR_SHEET_NAME) });
      case 'getConfig':      return out({ otMultiplier: getOtMultiplier() });
      case 'getBatch':       return out(getBatchForDate(p.date));
      case 'getRecentLog':   return out(getRecentLog(parseInt(p.limit || 30)));
      case 'ping':            return out({ ok: true, time: new Date().toISOString() });
      default:                return out({ error: 'unknown action: ' + p.action });
    }
  } catch (err) {
    return out({ error: err.message + ' | ' + err.stack });
  }
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'verifyPin':  return out({ valid: verifyPin(body.pin) });
      case 'changePin':  return out(changePin(body.oldPin, body.newPin));
      case 'saveDay':    return out(saveDay(body.date, body.rows, body.pin, body.enteredBy));
      case 'deleteRow':  return out(deleteRow(body.id, body.pin));
      case 'saveConfig': return out(saveConfig(body.key, body.value, body.pin));
      default:           return out({ error: 'unknown action' });
    }
  } catch (err) {
    return out({ error: err.message });
  }
}

// ══════════════════════════════════════════════════════════════
// SHEET ACCESS / BOOTSTRAP
// ══════════════════════════════════════════════════════════════
function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getLogSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(LOG_HEADERS);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#2C4E38').setFontColor('#fff');
    sh.setFrozenRows(1);
    sh.getRange('E2:E').setNumberFormat('@'); // Date column = plain text, กัน locale สลับวัน/เดือน
  }
  return sh;
}

function getCfgSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(CFG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG_SHEET_NAME);
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['PIN', DEFAULT_PIN]);
    sh.appendRow(['LINE_TOKEN', '']);
    sh.appendRow(['LINE_TARGET', '']);
    sh.appendRow(['OT_MULTIPLIER', DEFAULT_OT_MULT]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2C4E38').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPlotsSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(PLOTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PLOTS_SHEET_NAME);
    sh.appendRow(['Plot']);
    sh.getRange(1, 1).setFontWeight('bold').setBackground('#2C4E38').setFontColor('#fff');
    sh.setFrozenRows(1);
    // แก้/เพิ่ม/ลบชื่อแปลงได้ตรงนี้เลย ไม่ต้องแก้โค้ด
    ['T2V','T2SW','T4S','T4N','A2','A3','M2','M4(P9)','CO2','IRR-A2','C2-JF2','CNBD1'].forEach(v => sh.appendRow([v]));
  }
  return sh;
}

function getContractorsSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(CONTR_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONTR_SHEET_NAME);
    sh.appendRow(['Contractor']);
    sh.getRange(1, 1).setFontWeight('bold').setBackground('#2C4E38').setFontColor('#fff');
    sh.setFrozenRows(1);
    // แก้/เพิ่ม/ลบชื่อผู้รับเหมาได้ตรงนี้เลย ไม่ต้องแก้โค้ด
    ['Chhan Chhoeuk','Theang Thim','Da'].forEach(v => sh.appendRow([v]));
  }
  return sh;
}

// เรียกครั้งเดียวจาก Apps Script editor (Run) หลังวางโค้ดใหม่ เพื่อสร้างชีตทั้งหมดล่วงหน้า
function initSheets() {
  getLogSheet(); getCfgSheet(); getPlotsSheet(); getContractorsSheet();
  return 'ok';
}

function getListValues(sheetName) {
  const ss = getSS();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v.length > 0);
}

// ══════════════════════════════════════════════════════════════
// CONFIG / PIN
// ══════════════════════════════════════════════════════════════
function getCfgValue(key) {
  const rows = getCfgSheet().getDataRange().getValues();
  const r = rows.find(r => r[0] === key);
  return r ? String(r[1]).trim() : '';
}

function setCfgValue(key, value) {
  const sh = getCfgSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

function getOtMultiplier() {
  const v = parseFloat(getCfgValue('OT_MULTIPLIER'));
  return isNaN(v) ? DEFAULT_OT_MULT : v;
}

function getStoredPin() {
  const v = getCfgValue('PIN');
  return v || DEFAULT_PIN;
}

function verifyPin(pin) { return String(pin) === getStoredPin(); }

function changePin(oldPin, newPin) {
  if (!verifyPin(oldPin)) return { error: 'PIN เดิมไม่ถูกต้อง' };
  if (!/^\d{4}$/.test(String(newPin))) return { error: 'PIN ใหม่ต้องเป็นตัวเลข 4 หลัก' };
  setCfgValue('PIN', String(newPin));
  return { success: true };
}

function saveConfig(key, value, pin) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };
  const allowed = ['LINE_TOKEN', 'LINE_TARGET', 'OT_MULTIPLIER'];
  if (allowed.indexOf(key) === -1) return { error: 'ไม่อนุญาตให้แก้ค่านี้' };
  setCfgValue(key, value);
  return { success: true };
}

// ══════════════════════════════════════════════════════════════
// CORE: SAVE DAY (สร้างใหม่ครั้งแรก / อัปเดตถ้ามี batch ของวันนั้นอยู่แล้ว)
// ══════════════════════════════════════════════════════════════
function saveDay(isoDate, rows, pin, enteredBy) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return { error: 'รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)' };
  if (!rows || !rows.length) return { error: 'ไม่มีรายการให้บันทึก' };

  const otMult = getOtMultiplier();
  const sh = getLogSheet();
  const who = enteredBy || 'App';
  const now = new Date();

  // หา batch ที่ active อยู่แล้วของวันนี้ (ถ้ามี = โหมดอัปเดต)
  const existing = findActiveBatchForDate(sh, isoDate);
  const isUpdate = !!existing.batchId;
  const batchId  = existing.batchId || (isoDate + '_' + Date.now());

  // ถ้าอัปเดต: mark แถวเดิมของ batch นี้เป็น EDITED (เก็บ audit trail ไว้ ไม่ลบ)
  if (isUpdate) {
    existing.rowIndexes.forEach(rowIdx => {
      sh.getRange(rowIdx, LOG_HEADERS.indexOf('Status') + 1).setValue('EDITED');
      sh.getRange(rowIdx, LOG_HEADERS.indexOf('EditedAt') + 1).setValue(now.toISOString());
    });
  }

  // เขียนแถวใหม่ (เวอร์ชันล่าสุด) ทั้งหมดของวันนี้
  const savedRows = [];
  let grandTotal = 0;
  rows.forEach(r => {
    const numWorkers = parseFloat(r.numWorkers) || 0;
    const hours      = parseFloat(r.hours) || 0;
    const rate       = parseFloat(r.rate) || 0;
    const otHours    = parseFloat(r.otHours) || 0;
    const total      = +(numWorkers * (hours * rate + otHours * rate * otMult)).toFixed(2);
    grandTotal += total;

    const id = Date.now() + '_' + Math.floor(Math.random() * 10000);
    sh.appendRow([
      id, batchId, now.toISOString(), '', isoDate,
      r.plot || '', r.contractor || '',
      numWorkers, hours, rate, otHours, otMult,
      r.supervisor || '', total, who, 'ACTIVE', r.note || ''
    ]);

    savedRows.push({
      id, plot: r.plot || '', contractor: r.contractor || '',
      numWorkers, hours, rate, otHours, otMultiplier: otMult,
      supervisor: r.supervisor || '', total, note: r.note || ''
    });
  });

  sendLineSummary(isoDate, savedRows, grandTotal, isUpdate ? 'UPDATE' : 'NEW');

  return {
    success: true, batchId, isUpdate,
    rows: savedRows, grandTotal: +grandTotal.toFixed(2)
  };
}

// หา batch ที่ยัง ACTIVE อยู่ของวันที่ที่กำหนด (คืน batchId + row index ทั้งหมดของ batch นั้น)
function findActiveBatchForDate(sh, isoDate) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { batchId: null, rowIndexes: [] };

  const dateCol   = LOG_HEADERS.indexOf('Date') + 1;
  const statusCol = LOG_HEADERS.indexOf('Status') + 1;
  const batchCol  = LOG_HEADERS.indexOf('BatchID') + 1;

  const data = sh.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  let batchId = null;
  const rowIndexes = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[dateCol - 1]) === isoDate && row[statusCol - 1] === 'ACTIVE') {
      batchId = row[batchCol - 1];
    }
  }
  if (!batchId) return { batchId: null, rowIndexes: [] };

  for (let i = 0; i < data.length; i++) {
    if (data[i][batchCol - 1] === batchId && data[i][statusCol - 1] === 'ACTIVE') {
      rowIndexes.push(i + 2); // 1-based + header row
    }
  }
  return { batchId, rowIndexes };
}

function getBatchForDate(isoDate) {
  if (!isoDate) return { batchId: null, rows: [] };
  const sh = getLogSheet();
  const { batchId, rowIndexes } = findActiveBatchForDate(sh, isoDate);
  if (!batchId) return { batchId: null, rows: [] };

  const rows = rowIndexes.map(rowIdx => {
    const r = sh.getRange(rowIdx, 1, 1, LOG_HEADERS.length).getValues()[0];
    return rowToObj(r);
  });
  return { batchId, rows };
}

function rowToObj(r) {
  const o = {};
  LOG_HEADERS.forEach((h, i) => { o[h] = r[i]; });
  return {
    id: String(o.ID), plot: o.Plot, contractor: o.Contractor,
    numWorkers: parseFloat(o.NumWorkers) || 0, hours: parseFloat(o.Hours) || 0,
    rate: parseFloat(o.RatePerHourUSD) || 0, otHours: parseFloat(o.OTHours) || 0,
    otMultiplier: parseFloat(o.OTMultiplier) || DEFAULT_OT_MULT,
    supervisor: o.Supervisor, total: parseFloat(o.TotalUSD) || 0,
    note: o.Note || ''
  };
}

// ── ลบรายการเดียว (mark DELETED, ไม่ลบแถวจริงเพื่อ audit trail) ──
function deleteRow(id, pin) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };
  const sh = getLogSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { error: 'ไม่พบรายการ' };
  const idCol = LOG_HEADERS.indexOf('ID') + 1;
  const statusCol = LOG_HEADERS.indexOf('Status') + 1;
  const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.getRange(i + 2, statusCol).setValue('DELETED');
      return { success: true };
    }
  }
  return { error: 'ไม่พบรายการ ID: ' + id };
}

// ══════════════════════════════════════════════════════════════
// RECENT LOG (audit view — ล่าสุดของแต่ละวัน)
// ══════════════════════════════════════════════════════════════
function getRecentLog(limit) {
  const sh = getLogSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { records: [] };
  const data = sh.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();

  const byDate = {};
  data.forEach(r => {
    const o = rowFullToObj(r);
    if (o.Status !== 'ACTIVE') return;
    if (!byDate[o.Date]) byDate[o.Date] = { date: o.Date, rows: [], total: 0 };
    byDate[o.Date].rows.push(o);
    byDate[o.Date].total += parseFloat(o.TotalUSD) || 0;
  });

  const days = Object.values(byDate)
    .sort((a, b) => b.date.localeCompare(a.date))
    .slice(0, limit || 30)
    .map(d => ({ date: d.date, total: +d.total.toFixed(2), count: d.rows.length }));

  return { records: days, updatedAt: new Date().toISOString() };
}

function rowFullToObj(r) {
  const o = {};
  LOG_HEADERS.forEach((h, i) => { o[h] = r[i]; });
  return o;
}

// ══════════════════════════════════════════════════════════════
// LINE MESSAGING API
// ══════════════════════════════════════════════════════════════
function sendLineSummary(isoDate, rows, grandTotal, mode) {
  const token  = getCfgValue('LINE_TOKEN');
  const target = getCfgValue('LINE_TARGET');
  if (!token || !target) return;

  const [y, m, d] = isoDate.split('-');
  const dateLabel = `${parseInt(d)}/${parseInt(m)}/${y}`;
  const headEmoji = mode === 'UPDATE' ? '🔄 ปรับปรุงข้อมูลคนงาน' : '🆕 รายงานคนงานรายวัน';

  let lines = '';
  let totalWorkers = 0;
  rows.forEach(r => {
    totalWorkers += r.numWorkers;
    const otPart = r.otHours > 0 ? ` (+OT ${r.otHours}ชม.)` : '';
    const notePart = r.note ? `\n   📝 ${r.note}` : '';
    lines += `📍 ${r.plot} — ${r.contractor}\n` +
             `   คน ${r.numWorkers} | ${r.hours}ชม.${otPart} | คุม: ${r.supervisor}\n` +
             `   ค่าจ้าง $${r.total.toFixed(2)}${notePart}\n`;
  });

  const msg = `${headEmoji}\n📅 ${dateLabel}\n\n${lines}\n` +
              `👥 รวมคนงาน: ${totalWorkers} คน\n💰 ยอดรวม: $${grandTotal.toFixed(2)}`;

  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: target, messages: [{ type: 'text', text: msg }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.log('LINE summary error: ' + e.message);
  }
}
