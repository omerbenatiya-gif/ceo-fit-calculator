// ============================================================
// CEO FIT — סקר צרכים
// הדבקה: Extensions → Apps Script בתוך גיליון "CEO-FIT סקר צרכים"
// Deploy: Deploy → New deployment (או Manage deployments → Edit הקיים)
//         Type: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

const SURVEY_SHEET_NAME = 'תשובות';
const LEADS_SHEET_NAME  = 'ליידים מסקר';

// 20 עמודות: timestamp + q1 + (q2a,q2b) + (q3a,q3b) + q4...q9 + q10 + q11 + q12...q17
const SURVEY_HEADERS = [
  'חותמת זמן',
  'ספר על עצמך ועל העסק',            // q1
  'מספר לקוחות',                      // q2 — לקוחות
  'הכנסה חודשית',                     // q2 — הכנסה
  'שעות ביום בעסק',                   // q3 — שעות
  'מה שואב זמן ואנרגיה',              // q3 — מה שואב
  'עם מי גר',                         // q4
  'עבודה נוספת מעבר לעסק',            // q5
  'זוגיות ותמיכה בתהליך',             // q6
  'מה הכי רוצה לשפר בגוף',            // q7
  'איך גוף החלומות ישפיע על חייו',    // q8
  'קשר בין מראה פיזי להצלחה בעסק',   // q9
  'מה מקשה להתמיד באימונים',          // q10
  'מה מקשה להתמיד בתזונה',            // q11
  'כמה פעמים בשבוע אוכל בחוץ',        // q12
  'מה כן ולא עבד בתהליכים קודמים',   // q13
  'כמה השקיע בעבר ומה חסר',           // q14
  'מה חייב להיות בתהליך ליווי',       // q15
  'מה גורם להשקיע כסף בעצמו',         // q16
  'מה הכי חשוב בבחירת מאמן'           // q17
];

const LEADS_HEADERS = ['חותמת זמן', 'שם', 'טלפון'];

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'survey_cta') {
      writeCTARow(data);
    } else {
      writeSurveyRow(data);
    }
    output.setContent(JSON.stringify({ success: true }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }
  return output;
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#102036')
      .setFontColor('#3DF0FF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    for (var i = 2; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 240);
    }
  }
  return sheet;
}

function writeSurveyRow(data) {
  const sheet = getOrCreateSheet_(SURVEY_SHEET_NAME, SURVEY_HEADERS);

  // q2 is stored as "clients|income"
  const q2parts = (data.q2 || '').split('|');
  const q2clients = q2parts[0] || '';
  const q2income  = q2parts[1] || '';

  // q3 is stored as "hours|energyText"
  const q3parts  = (data.q3 || '').split('|');
  const q3hours  = q3parts[0] || '';
  const q3energy = q3parts.slice(1).join('|');

  sheet.appendRow([
    new Date(),
    data.q1  || '',
    q2clients, q2income,
    q3hours,   q3energy,
    data.q4  || '',
    data.q5  || '',
    data.q6  || '',
    data.q7  || '',
    data.q8  || '',
    data.q9  || '',
    data.q10 || '',
    data.q11 || '',
    data.q12 || '',
    data.q13 || '',
    data.q14 || '',
    data.q15 || '',
    data.q16 || '',
    data.q17 || ''
  ]);
}

function writeCTARow(data) {
  const sheet = getOrCreateSheet_(LEADS_SHEET_NAME, LEADS_HEADERS);
  sheet.appendRow([new Date(), data.name || '', data.phone || '']);
}
