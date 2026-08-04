/**
 * דשבורד הלוואות החברה - קוד צד שרת (Google Apps Script)
 * -----------------------------------------------------
 * הקוד הזה רץ בתוך Google Sheets ומשמש כ-"שרת" קטן:
 * הוא שומר את כל ההלוואות בגיליון, ומגיש אותן כ-JSON לדף האינטרנט.
 *
 * איך מתקינים - ראו את הקובץ README.md
 */

const SHEET_NAME = 'הלוואות';

// סדר השדות קבוע - כל שינוי כאן חייב להתאים גם לשינוי ב-index.html
const FIELDS = [
  'id', 'bankName', 'originalAmount', 'annualRate', 'monthlyPayment',
  'paymentDay', 'startDate', 'endDate', 'status',
  'manualBalance', 'manualBalanceDate'
];

const HEADERS = [
  'מזהה', 'שם בנק', 'סכום הלוואה מקורי', 'ריבית שנתית (%)', 'תשלום חודשי',
  'יום תשלום בחודש', 'תאריך תחילת הלוואה', 'תאריך סיום צפוי', 'סטטוס',
  'יתרה מעודכנת ידנית', 'תאריך העדכון הידני'
];

const DATE_FIELDS = ['startDate', 'endDate', 'manualBalanceDate'];
const NUMBER_FIELDS = ['originalAmount', 'annualRate', 'monthlyPayment', 'paymentDay', 'manualBalance'];

function doGet(e) {
  if (!checkToken(e.parameter && e.parameter.token)) return unauthorized();
  try {
    const sheet = getSheet();
    const range = sheet.getDataRange().getValues();
    const rows = range.slice(1).filter(r => r[0] !== '' && r[0] !== null);
    const loans = rows.map(rowToObject);
    return jsonResponse({ ok: true, loans: loans });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken(body.token)) return unauthorized();
    const sheet = getSheet();

    if (body.action === 'add') {
      const loan = body.loan || {};
      loan.id = Utilities.getUuid();
      sheet.appendRow(objectToRow(loan));
      return jsonResponse({ ok: true, id: loan.id });
    }

    if (body.action === 'update') {
      const rowIndex = findRowById(sheet, body.loan.id);
      if (rowIndex === -1) return jsonResponse({ ok: false, error: 'לא נמצאה הלוואה עם המזהה הזה' });
      const row = objectToRow(body.loan);
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      return jsonResponse({ ok: true });
    }

    if (body.action === 'delete') {
      const rowIndex = findRowById(sheet, body.id);
      if (rowIndex === -1) return jsonResponse({ ok: false, error: 'לא נמצאה הלוואה עם המזהה הזה' });
      sheet.deleteRow(rowIndex);
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: 'פעולה לא מוכרת: ' + body.action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ---------- הרשאה ----------
// הטוקן נשמר ב-Script Properties (Project Settings > Script Properties) ולא כאן בקוד,
// כדי שלא ייחשף בריפו הפומבי ב-GitHub. יש להגדיר שם מפתח בשם API_TOKEN.
// שימו לב: זו הגנה חלקית בלבד מפני סריקות אוטומטיות של כתובות פתוחות - היא לא
// מסתירה את הטוקן ממי שבאמת פותח את דף ה-index.html וקורא את קוד המקור שלו (Ctrl+U),
// כי דף סטטי חייב לשלוח את הטוקן בעצמו כדי להתאמת. הגנה אמיתית דורשת שרת ביניים.

function checkToken(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return !!expected && token === expected;
}

function unauthorized() {
  return jsonResponse({ ok: false, error: 'גישה לא מורשית - טוקן שגוי או חסר' });
}

// ---------- עזרים ----------

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowById(sheet, id) {
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // +2: דילוג על שורת הכותרות ואינדקס מ-1
  }
  return -1;
}

function rowToObject(row) {
  const obj = {};
  FIELDS.forEach((field, i) => {
    let value = row[i];
    if (DATE_FIELDS.indexOf(field) !== -1) {
      value = formatDateCell(value);
    } else if (NUMBER_FIELDS.indexOf(field) !== -1) {
      value = (value === '' || value === null || value === undefined) ? null : Number(value);
    } else {
      value = (value === null || value === undefined) ? '' : String(value);
    }
    obj[field] = value;
  });
  return obj;
}

function objectToRow(obj) {
  return FIELDS.map(field => {
    const value = obj[field];
    if (value === undefined || value === null) return '';
    return value;
  });
}

function formatDateCell(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
