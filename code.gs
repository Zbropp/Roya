const SHEET_NAME = 'Students';
const TEST_PREFIX = 'T';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('رؤية التعليمية')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const baseHeaders = ['الكود','اسم الطالب','النوع','الرقم الشخصي','رقم ولي الأمر'];
  const current = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const headers = current.filter(String);

  if (headers.length === 0) {
    sh.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
  } else {
    const missing = baseHeaders.filter(h => !headers.includes(h));
    if (missing.length) {
      sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 5))
    .setFontWeight('bold');
  return true;
}

function getSheet_() {
  setupSheet();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getData() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (!values.length) return {headers: [], rows: []};
  const headers = values[0].map(String);
  const rows = values.slice(1).filter(r => r.some(v => v !== ''));
  return {headers, rows};
}

function addStudent(data) {
  if (!data || !data.name || !data.gender) throw new Error('الاسم والنوع مطلوبان.');
  const sh = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sh.getLastRow();
    const all = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, Math.max(5, sh.getLastColumn())).getValues() : [];
    const prefix = data.gender === 'M' ? 'S1M' : 'S1F';
    let max = 0;
    all.forEach(r => {
      const c = String(r[0] || '');
      if (c.indexOf(prefix) === 0) {
        const n = parseInt(c.substring(prefix.length), 10);
        if (!isNaN(n)) max = Math.max(max, n);
      }
    });
    const code = prefix + (max + 1);

    sh.appendRow([
      code,
      String(data.name).trim(),
      data.gender === 'M' ? 'طالب' : 'طالبة',
      String(data.personal || '').trim(),
      String(data.guardian || '').trim()
    ]);
    return {ok: true, code};
  } finally {
    lock.releaseLock();
  }
}

function addTest() {
  const sh = getSheet_();
  const headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  let max = 0;
  headers.forEach(h => {
    const m = String(h).match(/^T(\d+)$/i);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  const next = 'T' + (max + 1);
  sh.getRange(1, sh.getLastColumn() + 1).setValue(next);
  return next;
}

function updateGrade(code, test, grade) {
  if (!code || !test) throw new Error('بيانات الدرجة غير مكتملة.');
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);
  const col = headers.indexOf(test);
  if (col === -1) throw new Error('الاختبار غير موجود.');
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(code)) { row = i + 1; break; }
  }
  if (row === -1) throw new Error('الطالب غير موجود.');
  const v = grade === '' || grade === null ? '' : Number(grade);
  if (grade !== '' && isNaN(v)) throw new Error('الدرجة يجب أن تكون رقمًا.');
  sh.getRange(row, col + 1).setValue(v);
  return true;
}

function getStudent(code) {
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(code)) {
      const obj = {};
      headers.forEach((h, j) => obj[h] = data[i][j]);
      return obj;
    }
  }
  return null;
}
