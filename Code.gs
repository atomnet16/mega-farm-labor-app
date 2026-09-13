// ============================================================
// MEGA FARM CAMBODIA — Labor Daily App · Google Apps Script
// Container-bound script (Extensions → Apps Script จาก Sheet "MGF Labor Daily Report")
// Deploy: Execute as Me · Access: Anyone
// ============================================================
// โครงสร้าง Sheet:
//   LaborLog     : ID, BatchID, Timestamp, EditedAt, Date, Activity, Plot, Contractor,
//                  NumWorkers, Hours, OTHours, Supervisor, EnteredBy, Status (ACTIVE/EDITED)
//   AppConfig    : Key, Value  → PIN, LINE_TOKEN, LINE_TARGET
//   Plots        : รายชื่อแปลง (คอลัมน์ A ตั้งแต่แถว 2) — แก้ได้ตรงในชีต
//   Contractors  : รายชื่อผู้รับเหมา (คอลัมน์ A ตั้งแต่แถว 2) — แก้ได้ตรงในชีต
//   Activities   : รายชื่อกิจกรรม (คอลัมน์ A ตั้งแต่แถว 2) — แก้ได้ตรงในชีต
//
// หมายเหตุ: Date เก็บเป็น "@" (Plain Text) รูปแบบ "YYYY-MM-DD" เท่านั้น — ห้ามให้
// Sheet ตีความเป็น Date object เด็ดขาด เพราะ locale ของ Google Sheet (US) จะสลับ
// วัน/เดือนสำหรับวันที่ ≤ 12 (บั๊กเดียวกับที่เจอใน Rainfall Record มาก่อน)
//
// Workflow: กิจกรรม/แปลง/ผู้รับเหมา/จำนวนคนคีย์ตอนเช้าได้โดยยังไม่ต้องรู้ชั่วโมง —
// saveDay() upsert ตาม (Date) เสมอ ดังนั้นกลับมาแก้ Hours/OTHours ทีหลังในวันเดียวกัน
// ก็ใช้ endpoint เดิม ไม่ต้องมี endpoint แยกสำหรับ "อัปเดตชั่วโมง"
// ============================================================

const LOG_SHEET_NAME    = 'LaborLog';
const CFG_SHEET_NAME    = 'AppConfig';
const PLOTS_SHEET_NAME  = 'Plots';
const CONTR_SHEET_NAME  = 'Contractors';
const ACT_SHEET_NAME    = 'Activities';
const DEFAULT_PIN       = '1234';
const TZ                = 'Asia/Phnom_Penh';

const LOG_HEADERS = [
  'ID','BatchID','Timestamp','EditedAt','Date','Activity','Plot','Contractor',
  'NumWorkers','Hours','OTHours','Supervisor','EnteredBy','Status'
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
      case 'getActivities':  return out({ activities: getListValues(ACT_SHEET_NAME) });
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

function getActivitiesSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(ACT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACT_SHEET_NAME);
    sh.appendRow(['Activity']);
    sh.getRange(1, 1).setFontWeight('bold').setBackground('#2C4E38').setFontColor('#fff');
    sh.setFrozenRows(1);
    // แก้/เพิ่ม/ลบชื่อกิจกรรมได้ตรงนี้เลย ไม่ต้องแก้โค้ด
    ['วัชพืช (Weeding)','ปลูก (Planting)','เก็บเกี่ยว (Harvesting)','ชลประทาน (Irrigation)',
     'บำรุงรักษา (Maintenance)','หว่าน (Sowing)','พ่นสาร (Spraying)','อื่น ๆ'].forEach(v => sh.appendRow([v]));
  }
  return sh;
}

// เรียกครั้งเดียวจาก Apps Script editor (Run) หลังวางโค้ดใหม่ เพื่อสร้างชีตทั้งหมดล่วงหน้า
function initSheets() {
  getLogSheet(); getCfgSheet(); getPlotsSheet(); getContractorsSheet(); getActivitiesSheet();
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
  const allowed = ['LINE_TOKEN', 'LINE_TARGET'];
  if (allowed.indexOf(key) === -1) return { error: 'ไม่อนุญาตให้แก้ค่านี้' };
  setCfgValue(key, value);
  return { success: true };
}

// ══════════════════════════════════════════════════════════════
// CORE: SAVE DAY (สร้างใหม่ครั้งแรก / อัปเดตถ้ามี batch ของวันนั้นอยู่แล้ว)
// ใช้ endpoint เดียวกันทั้งตอนคีย์เช้า (Hours=0) และตอนกลับมาเติมชั่วโมงเย็น
// ══════════════════════════════════════════════════════════════
function saveDay(isoDate, rows, pin, enteredBy) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return { error: 'รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)' };
  if (!rows || !rows.length) return { error: 'ไม่มีรายการให้บันทึก' };

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
  rows.forEach(r => {
    const numWorkers = parseFloat(r.numWorkers) || 0;
    const hours      = parseFloat(r.hours) || 0;
    const otHours    = parseFloat(r.otHours) || 0;

    const id = Date.now() + '_' + Math.floor(Math.random() * 10000);
    sh.appendRow([
      id, batchId, now.toISOString(), '', isoDate,
      r.activity || '', r.plot || '', r.contractor || '',
      numWorkers, hours, otHours, r.supervisor || '', who, 'ACTIVE'
    ]);

    savedRows.push({
      id, activity: r.activity || '', plot: r.plot || '', contractor: r.contractor || '',
      numWorkers, hours, otHours, supervisor: r.supervisor || ''
    });
  });

  sendLineSummary(isoDate, savedRows, isUpdate ? 'UPDATE' : 'NEW');

  return { success: true, batchId, isUpdate, rows: savedRows };
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
    if (normalizeDateCell(row[dateCol - 1]) === isoDate && row[statusCol - 1] === 'ACTIVE') {
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

// แปลง cell ของ Date ให้เป็น "YYYY-MM-DD" ไม่ว่าจะเก็บเป็น text หรือ Date object
// (Sheet บางครั้งตีความ string วันที่เป็น Date object เองแม้ตั้ง format เป็น Plain Text ไว้แล้ว —
// นี่คือสาเหตุที่ทำให้ getBatch หาข้อมูลของวันนั้นไม่เจอหลังกดส่ง)
function normalizeDateCell(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, TZ, 'yyyy-MM-dd');
  }
  return String(cell).trim();
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
    id: String(o.ID), activity: o.Activity, plot: o.Plot, contractor: o.Contractor,
    numWorkers: parseFloat(o.NumWorkers) || 0, hours: parseFloat(o.Hours) || 0,
    otHours: parseFloat(o.OTHours) || 0, supervisor: o.Supervisor
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
    const dateKey = normalizeDateCell(o.Date);
    if (o.Status !== 'ACTIVE') return;
    if (!byDate[dateKey]) byDate[dateKey] = { date: dateKey, rows: [], totalHours: 0, totalWorkers: 0, pending: 0 };
    byDate[dateKey].rows.push(o);
    byDate[dateKey].totalHours += parseFloat(o.Hours) || 0;
    byDate[dateKey].totalWorkers += parseFloat(o.NumWorkers) || 0;
    if ((parseFloat(o.Hours) || 0) <= 0) byDate[dateKey].pending++;
  });

  const days = Object.values(byDate)
    .sort((a, b) => b.date.localeCompare(a.date))
    .slice(0, limit || 30)
    .map(d => ({
      date: d.date, count: d.rows.length,
      totalHours: +d.totalHours.toFixed(1), totalWorkers: d.totalWorkers, pending: d.pending
    }));

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
function sendLineSummary(isoDate, rows, mode) {
  const token  = getCfgValue('LINE_TOKEN');
  const target = getCfgValue('LINE_TARGET');
  if (!token || !target) return;

  const [y, m, d] = isoDate.split('-');
  const dateLabel = `${parseInt(d)}/${parseInt(m)}/${y}`;
  const headEmoji = mode === 'UPDATE' ? '🔄 ปรับปรุงข้อมูลคนงาน' : '🆕 รายงานคนงานรายวัน';

  // บรรทัดต่อแปลง+กิจกรรม (รวมจำนวนคนของผู้รับเหมาหลายรายที่ทำงานเดียวกันแบบ 9+4=13)
  // ตามด้วยสรุปรวมคนต่อผู้รับเหมา — รายละเอียดชั่วโมง/OT ดูใน Sheet เท่านั้น
  const groups = {};
  const groupOrder = [];
  const workersByContractor = {};
  const contractorOrder = [];
  let totalWorkers = 0;

  rows.forEach(r => {
    const gKey = r.plot + '|' + r.activity;
    if (!groups[gKey]) {
      groups[gKey] = { plot: r.plot, activity: r.activity, parts: [], supervisors: [] };
      groupOrder.push(gKey);
    }
    groups[gKey].parts.push({ contractor: r.contractor || 'ไม่ระบุผู้รับเหมา', numWorkers: r.numWorkers });
    if (r.supervisor && groups[gKey].supervisors.indexOf(r.supervisor) === -1) {
      groups[gKey].supervisors.push(r.supervisor);
    }

    const cName = r.contractor || 'ไม่ระบุผู้รับเหมา';
    if (!(cName in workersByContractor)) { workersByContractor[cName] = 0; contractorOrder.push(cName); }
    workersByContractor[cName] += r.numWorkers;
    totalWorkers += r.numWorkers;
  });

  const plotLines = groupOrder.map(key => {
    const g = groups[key];
    const sum = g.parts.reduce((a, p) => a + p.numWorkers, 0);
    const multiContractor = new Set(g.parts.map(p => p.contractor)).size > 1;
    const countLabel = g.parts.length > 1
      ? `${g.parts.map(p => multiContractor ? `${p.numWorkers}(${p.contractor})` : `${p.numWorkers}`).join('+')}=${sum}`
      : `${sum}`;
    const supLabel = g.supervisors.length ? g.supervisors.join('/') : '-';
    return `📍 ${g.plot} — ${g.activity} — ${countLabel} คน — ผู้ควบคุม: ${supLabel}`;
  }).join('\n');

  const contractorLines = contractorOrder.map(name => `👷 ${name} — ${workersByContractor[name]} คน`).join('\n');

  const msg = `${headEmoji}\n📅 ${dateLabel}\n\n${plotLines}\n\n${contractorLines}\n\n` +
              `👥 รวมคนงาน: ${totalWorkers} คน`;

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
