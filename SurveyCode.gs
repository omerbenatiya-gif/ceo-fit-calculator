// ============================================================
// CEO FIT — סקר צרכים
// הוראות הפעלה:
//   1. פתח את הסקריפט (Extensions → Apps Script)
//   2. מחק את כל הקוד הקיים
//   3. הדבק קוד זה
//   4. שמור (Ctrl+S)
//   5. הרץ את הפונקציה setupSheets פעם אחת (▶ Run → setupSheets)
//      → זה יצור את שני הגיליונות עם הכותרות הנכונות
//   6. פרסם מחדש: Deploy → Manage deployments → Edit → New version → Deploy
// ============================================================

const SURVEY_SHEET = 'תשובות סקר';
const LEADS_SHEET  = 'ליידים מסקר';

// כותרות גיליון התשובות — כל שאלה בעמודה נפרדת
// שאלה 2 מפוצלת ל-2 עמודות (לקוחות + הכנסה)
// שאלה 3 מפוצלת ל-2 עמודות (שעות + מה שואב זמן)
// שאלות 10+11 נפרדות למרות שמופיעות על אותו מסך
const SURVEY_HEADERS = [
  'חותמת זמן',
  'ש1 | ספר על עצמך ועל העסק',
  'ש2 | מספר לקוחות',
  'ש2 | הכנסה חודשית (שח)',
  'ש3 | שעות ביום בעסק',
  'ש3 | מה שואב זמן ואנרגיה',
  'ש4 | עם מי גר',
  'ש5 | עבודה נוספת מעבר לעסק',
  'ש6 | זוגיות ותמיכה בתהליך',
  'ש7 | מה הכי רוצה לשפר בגוף',
  'ש8 | איך גוף החלומות ישפיע על חייו ועל העסק',
  'ש9 | קשר בין מראה פיזי להצלחה בעסק',
  'ש10 | מה מקשה להתמיד באימונים',
  'ש11 | מה מקשה להתמיד בתזונה',
  'ש12 | כמה פעמים בשבוע אוכל בחוץ',
  'ש13 | מה כן ולא עבד בתהליכים קודמים',
  'ש14 | כמה השקיע בעבר ומה חסר',
  'ש15 | מה חייב להיות בתהליך ליווי',
  'ש16 | מה גורם להשקיע כסף בעצמו',
  'ש17 | מה הכי חשוב בבחירת מאמן'
];

const LEADS_HEADERS = ['חותמת זמן', 'שם', 'טלפון'];

// ─── הגדרת גיליונות ───────────────────────────────────────
// הרץ פונקציה זו פעם אחת ידנית אחרי הדבקת הקוד!
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createOrReplaceSheet_(ss, SURVEY_SHEET, SURVEY_HEADERS, 220);
  createOrReplaceSheet_(ss, LEADS_SHEET,  LEADS_HEADERS,  180);
  SpreadsheetApp.flush();
  Logger.log('הגיליונות נוצרו בהצלחה');
}

function createOrReplaceSheet_(ss, name, headers, colWidth) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    // אם קיים — מנקה ומאתחל מחדש
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(name);
  }
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight('bold')
    .setBackground('#102036')
    .setFontColor('#3DF0FF')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 50);
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, colWidth);
  }
  return sheet;
}

// ─── נקודת כניסה HTTP ─────────────────────────────────────
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'survey_cta') {
      writeCTARow_(data);
    } else {
      writeSurveyRow_(data);
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

// ─── כתיבת שורת תשובות ────────────────────────────────────
function writeSurveyRow_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SURVEY_SHEET);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SURVEY_SHEET); }

  // q2 = "clients|income"
  const q2 = (data.q2 || '').split('|');
  // q3 = "hours|energyText"
  const q3 = (data.q3 || '').split('|');

  sheet.appendRow([
    new Date(),
    data.q1  || '',          // ש1
    q2[0]    || '',          // ש2 לקוחות
    q2[1]    || '',          // ש2 הכנסה
    q3[0]    || '',          // ש3 שעות
    q3.slice(1).join('|') || '',  // ש3 מה שואב
    data.q4  || '',          // ש4
    data.q5  || '',          // ש5
    data.q6  || '',          // ש6
    data.q7  || '',          // ש7
    data.q8  || '',          // ש8
    data.q9  || '',          // ש9
    data.q10 || '',          // ש10 אימונים
    data.q11 || '',          // ש11 תזונה
    data.q12 || '',          // ש12
    data.q13 || '',          // ש13
    data.q14 || '',          // ש14
    data.q15 || '',          // ש15
    data.q16 || '',          // ש16
    data.q17 || ''           // ש17
  ]);
}

// ─── כתיבת ליד (שם + טלפון) ──────────────────────────────
function writeCTARow_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(LEADS_SHEET); }

  sheet.appendRow([new Date(), data.name || '', data.phone || '']);
}
