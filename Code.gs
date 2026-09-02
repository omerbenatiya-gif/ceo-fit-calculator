// ============================================================
// CEO FIT — Google Apps Script לאוטומציית לידים
// ============================================================
// הדבקה: script.google.com → פרויקט חדש → Code.gs (מחק את הקוד הקיים)
// Deploy: Extensions → Apps Script → Deploy → New deployment
//         Type: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_LEADS      = 'לידים';
const SHEET_UPSELLS    = 'ליווי';
const SURVEY_SHEET_NAME = 'CEO-FIT סקר צרכים';
const SENDER_EMAIL     = 'service@joinceofit.com';
const EMAIL_SUBJECT    = 'התוכנית האישית שלך | C.E.O fit 💪';

// ← שנה ל-false אחרי אישור העיצוב (יחזור לדיליי 20 דקות)
const SEND_IMMEDIATELY = true;

// ============================================================
// נקודת כניסה: POST מהאתר
// ============================================================
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'upsell') {
      writeUpsellRow(data);
      sendUpsellNotification(data);
    } else if (data.type === 'survey') {
      writeSurveyRow(data);
    } else {
      writeLeadRow(data);
      schedulePlanEmail(data);
    }

    output.setContent(JSON.stringify({ success: true }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }

  return output;
}

// CORS preflight (GET)
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// כתיבה לגוגל שיט — לידים
// ============================================================
function writeLeadRow(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LEADS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LEADS);
    const headers = [
      'תאריך ושעה', 'שם מלא', 'טלפון', 'מייל', 'גיל',
      'מטרה', 'רמת פעילות', 'גובה (ס"מ)', 'משקל (ק"ג)',
      'BMI', 'קטגוריית BMI', 'למה לשנות את הגוף'
    ];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight('bold')
         .setBackground('#102036')
         .setFontColor('#3DF0FF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(12, 300);
  }

  const activityMap = {
    '1.2':   'יושב רוב היום',
    '1.375': 'פעילות קלה',
    '1.55':  'פעילות בינונית',
    '1.725': 'פעילות גבוהה'
  };
  const goalMap = { cut: 'להתחטב', bulk: 'לעלות במסה' };

  sheet.appendRow([
    new Date(),
    data.name          || '',
    data.phone         || '',
    data.email         || '',
    data.age           || '',
    goalMap[data.goal] || data.goal || '',
    activityMap[data.activity] || data.activity || '',
    data.height        || '',
    data.weight        || '',
    data.bmi           || '',
    data.bmiCategory   || '',
    data.why           || '',
  ]);
}

// ============================================================
// כתיבה לגוגל שיט — בקשות ליווי
// ============================================================
function writeUpsellRow(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_UPSELLS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UPSELLS);
    const headers = ['תאריך ושעה', 'שם מלא', 'טלפון', 'כמה זמן ניסית', 'למה לשנות עכשיו'];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight('bold')
         .setBackground('#102036')
         .setFontColor('#8AFFD6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 300);
  }

  sheet.appendRow([
    new Date(),
    data.name   || '',
    data.phone  || '',
    data.time   || '',
    data.reason || '',
  ]);
}

// ============================================================
// שליחת התראה לאחר בקשת ליווי
// ============================================================
function sendUpsellNotification(data) {
  GmailApp.sendEmail(
    SENDER_EMAIL,
    '🔥 בקשת ליווי חדשה | CEO FIT',
    '',
    {
      from:     SENDER_EMAIL,
      name:     'C.E.O FIT',
      replyTo:  SENDER_EMAIL,
      htmlBody: '<div dir="rtl" style="font-family:Arial;padding:20px;">' +
        '<h2 style="color:#102036;">בקשת ליווי חדשה!</h2>' +
        '<p><strong>שם:</strong> ' + (data.name || '') + '</p>' +
        '<p><strong>טלפון:</strong> ' + (data.phone || '') + '</p>' +
        '<p><strong>כמה זמן ניסית:</strong> ' + (data.time || '') + '</p>' +
        '<p><strong>למה לשנות עכשיו:</strong> ' + (data.reason || '') + '</p>' +
        '</div>',
    }
  );
}

// ============================================================
// תזמון / שליחה מיידית של מייל התוכנית
// ============================================================
function schedulePlanEmail(data) {
  if (SEND_IMMEDIATELY) {
    // מצב בדיקה — שולח מיידית ללא טריגר
    sendPlanEmailToLead_(data);
    return;
  }

  // מצב ייצור — דיליי 20 דקות
  const props   = PropertiesService.getScriptProperties();
  const leadKey = 'lead_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
  props.setProperty(leadKey, JSON.stringify(data));

  const trigger = ScriptApp.newTrigger('sendPlanEmail')
    .timeBased()
    .after(20 * 60 * 1000)
    .create();

  props.setProperty('trigger_' + trigger.getUniqueId(), leadKey);
}

// ============================================================
// מטפל בטריגר (מצב ייצור בלבד)
// ============================================================
function sendPlanEmail(e) {
  const props     = PropertiesService.getScriptProperties();
  const triggerId = e.triggerUid;
  const leadKey   = props.getProperty('trigger_' + triggerId);
  if (!leadKey) return;

  const dataStr = props.getProperty(leadKey);
  if (!dataStr) return;

  sendPlanEmailToLead_(JSON.parse(dataStr));

  // ניקוי
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getUniqueId() === triggerId; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  props.deleteProperty('trigger_' + triggerId);
  props.deleteProperty(leadKey);
}

// ============================================================
// שליחת מייל התוכנית (לוגיקה משותפת)
// ============================================================
function sendPlanEmailToLead_(data) {
  if (!data.email) return;

  GmailApp.sendEmail(
    data.email,
    EMAIL_SUBJECT,
    '',
    {
      from:     SENDER_EMAIL,
      name:     'C.E.O FIT',
      htmlBody: buildFullPlanEmail_(data),
      replyTo:  SENDER_EMAIL,
    }
  );
}

// ============================================================
// חישובי תזונה
// ============================================================
function calcNutrition_(data) {
  const h      = parseFloat(data.height)   || 175;
  const w      = parseFloat(data.weight)   || 78;
  const factor = parseFloat(data.activity) || 1.375;
  const goal   = data.goal || 'cut';

  const age      = parseFloat(data.age) || 40;
  const bmr      = Math.round(10 * w + 6.25 * h - 5 * age + 5);
  const tdee     = Math.round(bmr * factor);
  const calories = goal === 'cut'
    ? Math.round(tdee * 0.80)
    : Math.round(tdee * 1.10);

  const protein = Math.round(w * 2.2);
  const fat     = Math.round(w * 0.8);
  const carbs   = Math.round((calories - protein * 4 - fat * 9) / 4);

  return { bmr, tdee, calories, protein, fat, carbs };
}

// ============================================================
// HTML מייל מלא — תוכנית אישית
// כל הסגנונות inline + bgcolor לתאימות לקוחות מייל
// ============================================================
function buildFullPlanEmail_(data) {
  const iscut = (data.goal || 'cut') === 'cut';
  const n     = calcNutrition_(data);
  const name  = data.name || 'חבר';
  const bmi   = data.bmi || parseFloat((data.weight / Math.pow(data.height / 100, 2)).toFixed(1));
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const goal  = iscut ? 'להתחטב' : 'לעלות במסה';

  return '<!DOCTYPE html>' +
  '<html dir="rtl" lang="he">' +
  '<head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/></head>' +

  '<body bgcolor="#071828" style="margin:0;padding:0;background-color:#071828;font-family:Arial,sans-serif;direction:rtl;">' +

  // ── outer centering table ──
  '<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#071828" style="background-color:#071828;">' +
  '<tr><td align="center" style="padding:20px 8px 40px;">' +

  // ── 600px card ──
  '<table width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;border:1px solid #1E3A52;">' +

  // ── HEADER ──
  '<tr><td bgcolor="#071828" align="center" style="background-color:#071828;padding:28px 20px 0;">' +
    '<div style="font-size:34px;font-weight:900;letter-spacing:6px;line-height:1;">' +
      '<span style="color:#EAF6FF;">C.E.O</span>' +
      '<span style="color:#8AFFD6;"> fit</span>' +
    '</div>' +
    '<div style="color:#6B8299;font-size:11px;margin-top:10px;letter-spacing:1px;">התוכנית האישית שלך &bull; ' + today + '</div>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;">' +
      '<tr>' +
        '<td width="50%" height="3" bgcolor="#3DF0FF" style="background-color:#3DF0FF;font-size:0;line-height:0;"> </td>' +
        '<td width="50%" height="3" bgcolor="#8AFFD6" style="background-color:#8AFFD6;font-size:0;line-height:0;"> </td>' +
      '</tr>' +
    '</table>' +
  '</td></tr>' +

  // ── BODY ──
  '<tr><td bgcolor="#071828" style="background-color:#071828;padding:20px 16px;">' +

    // פתיח
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;">' +
    '<tr><td bgcolor="#0D2035" style="background-color:#0D2035;padding:16px 18px;border:1px solid #1E3A52;">' +
      '<div style="font-size:15px;font-weight:bold;color:#3DF0FF;margin-bottom:6px;">הינה תוכנית האימונים ותפריט התזונה שלך, ' + name + '! &#128170;</div>' +
      '<div style="color:#A0B8CC;font-size:12px;line-height:1.7;">' +
        'מטרה: <b style="color:#3DF0FF;">' + goal + '</b>' +
        '&nbsp;&nbsp;|&nbsp;&nbsp;' +
        'יעד קלוריות יומי: <b style="color:#3DF0FF;">' + n.calories + ' קק"ל</b>' +
      '</div>' +
    '</td></tr>' +
    '</table>' +

    // הנתונים שלך
    emailSectionHeader_('הנתונים שלך') +
    '<table width="100%" cellpadding="0" cellspacing="3" style="margin-bottom:8px;">' +
    '<tr>' +
      statCell_('גיל',     data.age || '', 'שנים') +
      statCell_('גובה',    data.height,    'ס"מ') +
      statCell_('משקל',    data.weight,    'ק"ג') +
      statCell_('BMI',     bmi,            data.bmiCategory || '') +
    '</tr>' +
    '</table>' +

    // מה הגוף שלך צריך
    emailSectionHeader_('מה הגוף שלך צריך') +
    '<table width="100%" cellpadding="0" cellspacing="3" style="margin-bottom:16px;">' +
    '<tr>' +
      statCell_('קלוריות', n.calories,   'קק"ל') +
      statCell_('חלבון',   n.protein,    'גרם') +
      statCell_('פחמימות', n.carbs,      'גרם') +
      statCell_('שומן',    n.fat,        'גרם') +
    '</tr>' +
    '</table>' +

    // תפריט
    emailSectionHeader_('תפריט תזונה יומי | ' + (iscut ? 'חיטוב' : 'מסה')) +
    (iscut ? mealPlanCut_(n) : mealPlanBulk_(n)) +

    // דגשי תזונה
    emailSectionHeader_('דגשי תזונה') +
    (iscut ? principlesCut_() : principlesBulk_()) +

    // אימונים
    emailSectionHeader_('תוכנית אימונים | ' + (iscut ? 'חיטוב' : 'מסה')) +
    workoutMethodBox_() +
    (iscut ? workoutCut_() : workoutBulk_()) +

  '</td></tr>' +

  // ── FOOTER ──
  '<tr><td bgcolor="#071828" align="center" style="background-color:#071828;padding:16px 20px;border-top:1px solid #1E3A52;">' +
    '<div style="font-size:18px;font-weight:900;letter-spacing:4px;line-height:1;">' +
      '<span style="color:#EAF6FF;">C.E.O</span>' +
      '<span style="color:#8AFFD6;"> fit</span>' +
    '</div>' +
    '<div style="color:#6B8299;font-size:11px;margin-top:4px;">תוכנית אישית לבעלי עסקים | ' + SENDER_EMAIL + '</div>' +
  '</td></tr>' +

  '</table>' + // סוף 600px
  '</td></tr>' +
  '</table>' + // סוף outer

  '</body></html>';
}

// ── כותרת סקציה למייל ──
function emailSectionHeader_(title) {
  return '<table width="100%" cellpadding="0" cellspacing="0" style="margin:18px 0 10px;">' +
    '<tr><td bgcolor="#071828" style="background-color:#071828;border-right:4px solid #3DF0FF;padding:3px 10px;">' +
      '<b style="font-size:13px;color:#3DF0FF;">' + title + '</b>' +
    '</td></tr>' +
    '</table>';
}

// ── תא סטטיסטיקה ──
function statCell_(label, value, unit) {
  return '<td bgcolor="#0D2035" align="center" style="background-color:#0D2035;border:1px solid #1E3A52;padding:10px 4px;">' +
    '<div style="color:#6B8299;font-size:10px;margin-bottom:3px;">' + label + '</div>' +
    '<b style="font-size:16px;color:#3DF0FF;">' + value + '</b><br/>' +
    '<span style="color:#6B8299;font-size:10px;">' + unit + '</span>' +
  '</td>';
}

// ── תפריט חיטוב ──
function mealPlanCut_(n) {
  var op = '<b style="color:#8AFFD6;font-size:10px;">';
  var cl = '</b>&nbsp;&nbsp;';
  return buildMealTable_([
    ['ארוחת בוקר',
      op + 'אופציה 1:' + cl + 'חביתה מ3 ביצים עם ירקות ו2 פרוסות לחם שיפון<br/>' +
      op + 'אופציה 2:' + cl + 'גביע יוגורט יווני 0% עם חצי כוס גרנולה ללא סוכר ופרי<br/>' +
      op + 'אופציה 3:' + cl + 'גביע קוטג\' 5% עם 2 פרוסות לחם מלא וירקות',
      Math.round(n.calories * 0.25)],
    ['ארוחת צהריים',
      op + 'אופציה 1:' + cl + 'חתיכת חזה עוף בגריל עם כוס אורז מלא וסלט ירקות<br/>' +
      op + 'אופציה 2:' + cl + 'פחית טונה במים עם 2 פרוסות לחם מלא וירקות<br/>' +
      op + 'אופציה 3:' + cl + 'חביתה מ3 ביצים עם בטטה בינונית וסלט ירקות',
      Math.round(n.calories * 0.35)],
    ['נשנוש אמצע יום',
      op + 'אופציה 1:' + cl + 'חצי גביע קוטג\' 5% עם חופן אגוזים<br/>' +
      op + 'אופציה 2:' + cl + 'תפוח עם 2 כפות חמאת בוטנים טבעית<br/>' +
      op + 'אופציה 3:' + cl + 'גביע יוגורט יווני 0% עם חופן שקדים',
      Math.round(n.calories * 0.10)],
    ['ארוחת ערב',
      op + 'אופציה 1:' + cl + 'חתיכת סלמון אפויה עם ירקות מוקפצים (מינימום שמן)<br/>' +
      op + 'אופציה 2:' + cl + 'חתיכת חזה עוף בגריל עם ירקות מאודים ומרק ירקות<br/>' +
      op + 'אופציה 3:' + cl + '4 ביצים קשות עם סלט ירקות גדול וכף טחינה גולמית',
      Math.round(n.calories * 0.30)],
  ]);
}

// ── תפריט מסה ──
function mealPlanBulk_(n) {
  var op = '<b style="color:#8AFFD6;font-size:10px;">';
  var cl = '</b>&nbsp;&nbsp;';
  return buildMealTable_([
    ['ארוחת בוקר',
      op + 'אופציה 1:' + cl + '3 ביצים שלמות עם קערת שיבולת שועל, בננה וכוס חלב 3%<br/>' +
      op + 'אופציה 2:' + cl + 'חביתה מ3 ביצים עם 3 פרוסות לחם מלא, חצי אבוקדו וכוס חלב<br/>' +
      op + 'אופציה 3:' + cl + 'גביע יוגורט יווני 2% עם כוס גרנולה, בננה ו2 כפות אגוזים',
      Math.round(n.calories * 0.25)],
    ['ארוחת צהריים',
      op + 'אופציה 1:' + cl + 'חתיכת חזה עוף גדולה עם 1.5 כוס אורז לבן וירקות מבושלים<br/>' +
      op + 'אופציה 2:' + cl + 'חתיכת בשר בקר רזה עם 2 תפוחי אדמה בינוניים וסלט<br/>' +
      op + 'אופציה 3:' + cl + 'חתיכת סלמון עם כוס פסטה מלאה וירקות מאודים',
      Math.round(n.calories * 0.35)],
    ['נשנוש אמצע יום',
      op + 'אופציה 1:' + cl + 'שייק חלבון עם בננה וכפית חמאת בוטנים<br/>' +
      op + 'אופציה 2:' + cl + '3 פרוסות לחם מלא עם 3 כפות חמאת שקדים<br/>' +
      op + 'אופציה 3:' + cl + 'גביע יוגורט 3% עם חצי כוס גרנולה וחופן אגוזים',
      Math.round(n.calories * 0.10)],
    ['ארוחת ערב',
      op + 'אופציה 1:' + cl + 'חתיכת בקר רזה או סלמון עם תפוח אדמה אפוי בינוני וסלט<br/>' +
      op + 'אופציה 2:' + cl + 'חתיכת חזה עוף בתנור עם בטטה בינונית וירקות מאודים<br/>' +
      op + 'אופציה 3:' + cl + '3 ביצים שלמות עם 3 פרוסות גבינה לבנה 5% ו2 פרוסות לחם',
      Math.round(n.calories * 0.30)],
  ]);
}

// ── בניית טבלת ארוחות ──
function buildMealTable_(meals) {
  var rows = meals.map(function(m, i) {
    var bg = (i % 2 === 0) ? '#0D2035' : '#0A1A28';
    return '<tr>' +
      '<td bgcolor="' + bg + '" style="background-color:' + bg + ';padding:8px 10px;border:1px solid #071828;white-space:nowrap;">' +
        '<b style="color:#8AFFD6;font-size:11px;">' + m[0] + '</b>' +
      '</td>' +
      '<td bgcolor="' + bg + '" style="background-color:' + bg + ';padding:8px 10px;border:1px solid #071828;color:#C8D8E8;font-size:12px;line-height:1.8;">' +
        m[1] +
      '</td>' +
      '<td bgcolor="' + bg + '" align="center" style="background-color:' + bg + ';padding:8px 10px;border:1px solid #071828;white-space:nowrap;">' +
        '<b style="color:#3DF0FF;font-size:12px;">' + m[2] + '</b>' +
      '</td>' +
    '</tr>';
  }).join('');

  return '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #1E3A52;margin-bottom:8px;">' +
    '<tr><td colspan="3" bgcolor="#102036" style="background-color:#102036;padding:6px 10px;border-bottom:1px solid #1E3A52;">' +
      '<span style="color:#8AFFD6;font-size:10px;font-weight:bold;">&#9998;&nbsp;לבחור אופציה אחת מכל ארוחה</span>' +
    '</td></tr>' +
    '<tr>' +
      '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:11px;padding:7px 10px;border:1px solid #1E3A52;text-align:right;">ארוחה</th>' +
      '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:11px;padding:7px 10px;border:1px solid #1E3A52;text-align:right;">מה לאכול</th>' +
      '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:11px;padding:7px 10px;border:1px solid #1E3A52;text-align:center;">קק"ל</th>' +
    '</tr>' +
    rows +
  '</table>';
}

// ── דגשי תזונה ──
function principlesList_(items) {
  var rows = items.map(function(t, i) {
    var bg = (i % 2 === 0) ? '#0D2035' : '#0A1A28';
    return '<tr>' +
      '<td bgcolor="' + bg + '" align="center" style="background-color:' + bg + ';padding:7px 8px;border:1px solid #071828;color:#3DF0FF;font-size:13px;width:18px;">&#10022;</td>' +
      '<td bgcolor="' + bg + '" style="background-color:' + bg + ';padding:7px 10px;border:1px solid #071828;color:#A0B8CC;font-size:12px;line-height:1.6;">' + t + '</td>' +
    '</tr>';
  }).join('');
  return '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #1E3A52;margin-bottom:8px;">' + rows + '</table>';
}

function principlesCut_() {
  return principlesList_([
    'לשתות 2.5 עד 3 ליטר מים ביום. חשוב לחילוף חומרים ולצמצום תחושת רעב',
    'לא לדלג על ארוחות. ארוחות קטנות ותכופות שומרות על רמת סוכר יציבה',
    'להגביל פחמימות בשעות הלילה. אחרי 18:00 עדיף חלבון וירקות',
    'להוסיף 30 דקות הליכה ארבע פעמים בשבוע לתהליך החיטוב',
    'חלבון בכל ארוחה. זה מה שמגן על השרירים בזמן חיטוב',
    'לקרוא תוויות ולבדוק את סך הקלוריות, החלבון ואת השומן הרווי',
    'שבת ופינוק? ארוחה אחת חופשית בסדר. פשוט לחזור לשגרה בסעודה הבאה',
  ]);
}

function principlesBulk_() {
  return principlesList_([
    'לאכול כל 3 עד 4 שעות. גוף שבונה שרירים צריך אספקה קבועה של אנרגיה',
    'חלבון בכל ארוחה. מינימום 30 גרם לארוחה',
    'פחמימות סביב האימון, לפני ואחרי, לביצועים ולהתאוששות מקסימלית',
    'לא לפחד משומן בריא. אגוזים, שמן זית וסלמון חשובים להורמונים',
    'לשתות לפחות 3 ליטר מים ביום. הידרציה קריטית לבניית שרירים',
    'שינה של 7 עד 8 שעות. 70% מבניית השריר קורה בשינה',
    'אם לא עולים 0.3 עד 0.5 ק"ג לשבוע יש להוסיף 100 עד 200 קלוריות ליום',
  ]);
}

// ── תיבת הסבר שיטת Full Body ──
function workoutMethodBox_() {
  return '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #1E3A52;margin-bottom:14px;">' +
    '<tr><td bgcolor="#0A1A28" style="background-color:#0A1A28;padding:14px 16px;border-right:4px solid #8AFFD6;">' +
      '<div style="color:#8AFFD6;font-size:12px;font-weight:bold;margin-bottom:8px;">&#x2605;&nbsp;שיטת Full Body | למה זה עובד לבעלי עסקים?</div>' +
      '<div style="color:#A0B8CC;font-size:11px;line-height:1.8;">' +
        '&#x2022;&nbsp;<b style="color:#C8D8E8;">3 אימונים בשבוע</b> בלבד. כל אימון מכסה את כל הגוף&nbsp;&nbsp;' +
        '&#x2022;&nbsp;<b style="color:#C8D8E8;">45 דקות מספיקות</b>. ללא בזבוז זמן<br/>' +
        '&#x2022;&nbsp;כל שריר מאומן <b style="color:#C8D8E8;">3 פעמים בשבוע</b>. זה הסוד לתוצאות מהירות&nbsp;&nbsp;' +
        '&#x2022;&nbsp;<b style="color:#C8D8E8;">A / B / C</b>. שלושה אימונים משלימים לסירוגין<br/>' +
        '&#x2022;&nbsp;אין חדר כושר? <b style="color:#C8D8E8;">אימון גיבוי ביתי</b> מחכה לך בתחתית' +
      '</div>' +
    '</td></tr>' +
  '</table>';
}

// ── בלוק אימון ──
function wBlock_(title, rows, showCols) {
  var showColumns = (showCols !== false);

  var headerRow = showColumns
    ? '<tr>' +
        '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:10px;padding:6px 8px;border:1px solid #1E3A52;text-align:right;">תרגיל</th>' +
        '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:10px;padding:6px 8px;border:1px solid #1E3A52;text-align:center;width:56px;">סטים</th>' +
        '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:10px;padding:6px 8px;border:1px solid #1E3A52;text-align:center;width:66px;">חזרות</th>' +
        '<th bgcolor="#102036" style="background-color:#102036;color:#3DF0FF;font-size:10px;padding:6px 8px;border:1px solid #1E3A52;text-align:right;">דגשים</th>' +
      '</tr>'
    : '';

  var dataRows = rows.map(function(r, i) {
    var bg  = (i % 2 === 0) ? '#0D2035' : '#0A1A28';
    var ex  = r[0] || '';
    var sets = r[1] || '';
    var reps = r[2] || '';
    var tip  = r[3] || '';

    if (!showColumns) {
      return '<tr><td bgcolor="' + bg + '" colspan="4" style="background-color:' + bg + ';padding:6px 10px;border:1px solid #071828;color:#A0B8CC;font-size:11px;">' + ex + '</td></tr>';
    }
    return '<tr>' +
      '<td bgcolor="' + bg + '" style="background-color:' + bg + ';padding:6px 8px;border:1px solid #071828;color:#C8D8E8;font-size:11px;">' + ex + '</td>' +
      '<td bgcolor="' + bg + '" align="center" style="background-color:' + bg + ';padding:6px 8px;border:1px solid #071828;"><b style="color:#3DF0FF;font-size:11px;">' + sets + '</b></td>' +
      '<td bgcolor="' + bg + '" align="center" style="background-color:' + bg + ';padding:6px 8px;border:1px solid #071828;color:#8AFFD6;font-size:11px;">' + reps + '</td>' +
      '<td bgcolor="' + bg + '" style="background-color:' + bg + ';padding:6px 8px;border:1px solid #071828;color:#6B8299;font-size:10px;">' + tip + '</td>' +
    '</tr>';
  }).join('');

  return '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #1E3A52;margin-bottom:10px;">' +
    '<tr><th bgcolor="#0D2035" colspan="4" style="background-color:#0D2035;color:#8AFFD6;font-size:12px;font-weight:bold;padding:8px 10px;border-bottom:1px solid #1E3A52;text-align:right;">' + title + '</th></tr>' +
    headerRow +
    dataRows +
  '</table>';
}

// ── תוכנית חיטוב ──
function workoutCut_() {
  return wBlock_('עקרונות האימון', [
    ['45 דקות | 3 פעמים בשבוע | יום מנוחה בין אימון לאימון'],
    ['מנוחה בין סטים: 45 עד 90 שניות | חזרות מורכב 8 עד 15, פשוט 12 עד 30'],
    ['חימום לפני כל שריר גדול (חזה / רגליים / גב). סט קל במשקל נמוך'],
    ['להוסיף הליכה של 30 דקות ארבע פעמים בשבוע לתהליך החיטוב'],
  ], false) +
  wBlock_('אימון A | דגש חזה', [
    ['לחיצת חזה במוט (ספסל ישר)',        '3 סטים', '8 עד 15',    'הורד באיטיות, מרפקים 45 מעלות'],
    ['פתיחות חזה בפולי',                  '2 סטים', '12 עד 30',   'תנועה רחבה, כיווץ בחזה'],
    ['סקוואט עם מוט',                     '3 סטים', '8 עד 15',    'ירכיים מקביל לרצפה, דחיפה מהעקבים'],
    ['חתירה עם מוט (כיפוף קדימה)',        '3 סטים', '8 עד 15',    'גב ישר, משוך לבטן תחתונה'],
    ['לחיצת כתפיים עם משקולות',           '3 סטים', '8 עד 15',    'מעל הראש, גב ישר'],
    ['כפיפת מרפק קדמי',                   '2 סטים', '12 עד 30',   'מרפקים צמודים, בלי תנופה'],
    ['פשיטת מרפק בפולי',                  '2 סטים', '12 עד 30',   'יישר עד הסוף, כווץ'],
    ['פלאנק',                             '2 סטים', 'מקסימום',    'גוף ישר, בטן כווצה'],
  ]) +
  wBlock_('אימון B | דגש רגליים', [
    ['לחיצת רגליים במכונה',               '3 סטים', '8 עד 15',  'ירידה מבוקרת, בלי לנעול ברכיים'],
    ['דדליפט רומני',                       '3 סטים', '8 עד 15',  'רגליים כמעט ישרות, מתיחה באחורי הירך'],
    ['לחיצת חזה בשיפוע עם משקולות',       '3 סטים', '8 עד 15',  'שיפוע קל, דחוף למעלה ופנימה'],
    ['משיכת פולי עליון לחזה',             '3 סטים', '8 עד 15',  'משוך לעצם החזה, כתפיים אחורה'],
    ['הרמות צד עם משקולות',               '2 סטים', '12 עד 30', 'עד גובה הכתפיים, מרפק מכופף קלות'],
    ['כפיפת מרפק פטיש',                   '2 סטים', '12 עד 30', 'אחיזה ניטרלית, בלי תנופה'],
    ['הרמות רגליים בשכיבה',               '2 סטים', '12 עד 30', 'הורד איטי בלי לגעת ברצפה'],
  ]) +
  wBlock_('אימון C | דגש גב', [
    ['חתירה חד צדדית עם ספסל',            '3 סטים', '8 עד 15',  'גב ישר, משוך למותן, כיווץ בגב'],
    ['משיכת פולי תחתון לבטן',             '3 סטים', '8 עד 15',  'גב זקוף, כתפיים אחורה'],
    ['מספריים בהליכה עם משקולות',         '3 סטים', '8 עד 15',  'צעד גדול, ברך אחורית לרצפה'],
    ['לחיצת חזה במכונה',                  '3 סטים', '8 עד 15',  'גב צמוד למשענת, דחיפה מבוקרת'],
    ['לחיצת כתפיים במכונה',               '3 סטים', '8 עד 15',  'ירידה איטית, שליטה מלאה'],
    ['פשיטת מרפק מעל הראש',               '2 סטים', '12 עד 30', 'מרפקים צמודים לראש'],
    ['כפיפות בטן',                         '2 סטים', '12 עד 30', 'הרם חלק עליון, כווץ בטן'],
  ]) +
  wBlock_('אימון גיבוי ביתי | משקל גוף', [
    ['סקוואט משקל גוף',                    '3 סטים', '15 עד 25',        'ירכיים מקביל לרצפה, גב ישר'],
    ['מספריים בהליכה',                      '2 סטים', '12 עד 20 לכל רגל', 'צעד גדול, שיווי משקל'],
    ['שכיבות סמיכה',                        '3 סטים', 'מקסימום',          'גוף ישר, חזה כמעט נוגע ברצפה'],
    ['חתירה הפוכה מתחת לשולחן',             '3 סטים', 'מקסימום',          'גוף ישר, משוך חזה לשולחן'],
    ['לחיצת כתפיים הפוכה, רגליים גבוהות',  '3 סטים', '10 עד 20',          'ידיים ברוחב כתפיים, ראש למטה'],
    ['מקבילים על קצה כיסא',                '2 סטים', '12 עד 20',          'ידיים מאחור, הורד בכיפוף מרפקים'],
    ['פלאנק עם הרמות רגליים לסירוגין',      '2 סטים', 'מקסימום',          'גוף ישר, הרם רגל אחת בכל פעם'],
  ]);
}

// ============================================================
// סקר צרכים — גיליון נפרד
// ============================================================
function getSurveySheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SURVEY_SHEET_ID');
  let sheet;
  if (id) {
    sheet = SpreadsheetApp.openById(id).getActiveSheet();
  } else {
    const ss = SpreadsheetApp.create(SURVEY_SHEET_NAME);
    id = ss.getId();
    props.setProperty('SURVEY_SHEET_ID', id);
    sheet = ss.getActiveSheet();
    sheet.setName('תשובות');
    const headers = [
      'חותמת זמן',
      'ספר על עצמך ועל העסק',
      'כמה לקוחות ומה ההכנסה החודשית',
      'כמה שעות ביום בעסק ומה שואב זמן',
      'עם מי אתה גר',
      'עבודה נוספת מעבר לעסק',
      'זוגיות ותמיכה בתהליך',
      'מה הכי רוצה לשנות בגוף',
      'למה דווקא עכשיו',
      'מה יאפשר לך בעוד 6 חודשים',
      'קשר בין מראה פיזי להצלחה בעסק',
      'מה מקשה להתמיד באימונים',
      'מה מקשה להתמיד בתזונה',
      'כמה פעמים בשבוע אוכל בחוץ',
      'תהליכים שניסית בעבר',
      'כמה השקעת בעבר ומה חסר',
      'מה חייב להיות בתהליך ליווי',
      'מה גורם לך להשקיע כסף בעצמך',
      'מה הכי חשוב בבחירת מאמן',
      'מה מרגיש כשמדמיין את הגוף הרצוי',
      'דבר אחד שהיית פותר בכושר',
      'דבר אחד שמונע ממך להגיע לגוף הרצוי'
    ];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#102036')
      .setFontColor('#3DF0FF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    for (var i = 2; i <= 22; i++) { sheet.setColumnWidth(i, 240); }
  }
  return sheet;
}

function writeSurveyRow(data) {
  var sheet = getSurveySheet_();
  sheet.appendRow([
    new Date(),
    data.q1  || '', data.q2  || '', data.q3  || '', data.q4  || '',
    data.q5  || '', data.q6  || '', data.q7  || '', data.q8  || '',
    data.q9  || '', data.q10 || '', data.q11 || '', data.q12 || '',
    data.q13 || '', data.q14 || '', data.q15 || '', data.q16 || '',
    data.q17 || '', data.q18 || '', data.q19 || '', data.q20 || '', data.q21 || ''
  ]);
}

// ── תוכנית מסה ──
function workoutBulk_() {
  return wBlock_('עקרונות האימון', [
    ['45 דקות | 3 פעמים בשבוע | יום מנוחה בין אימון לאימון'],
    ['מנוחה בין סטים: 90 עד 120 שניות | חזרות מורכב 6 עד 12, פשוט 10 עד 15'],
    ['חימום לפני כל שריר גדול (חזה / רגליים / גב). סט קל במשקל נמוך'],
    ['עדיפות לעומסים גבוהים עם טכניקה נקייה. איכות לפני כמות'],
  ], false) +
  wBlock_('אימון A | דגש חזה', [
    ['לחיצת חזה במוט (ספסל ישר)',        '4 סטים', '6 עד 12',    'הורד באיטיות, מרפקים 45 מעלות'],
    ['פתיחות חזה בפולי',                  '3 סטים', '10 עד 15',   'תנועה רחבה, כיווץ בחזה'],
    ['סקוואט עם מוט',                     '4 סטים', '6 עד 12',    'ירכיים מקביל לרצפה, דחיפה מהעקבים'],
    ['חתירה עם מוט (כיפוף קדימה)',        '4 סטים', '6 עד 12',    'גב ישר, משוך לבטן תחתונה'],
    ['לחיצת כתפיים עם משקולות',           '3 סטים', '6 עד 12',    'מעל הראש, גב ישר'],
    ['כפיפת מרפק קדמי',                   '3 סטים', '10 עד 15',   'מרפקים צמודים, בלי תנופה'],
    ['פשיטת מרפק בפולי',                  '3 סטים', '10 עד 15',   'יישר עד הסוף, כווץ'],
    ['פלאנק',                             '2 סטים', 'מקסימום',    'גוף ישר, בטן כווצה'],
  ]) +
  wBlock_('אימון B | דגש רגליים', [
    ['לחיצת רגליים במכונה',               '4 סטים', '6 עד 12',  'ירידה מבוקרת, בלי לנעול ברכיים'],
    ['דדליפט רומני',                       '4 סטים', '6 עד 12',  'רגליים כמעט ישרות, מתיחה באחורי הירך'],
    ['לחיצת חזה בשיפוע עם משקולות',       '4 סטים', '6 עד 12',  'שיפוע קל, דחוף למעלה ופנימה'],
    ['משיכת פולי עליון לחזה',             '4 סטים', '6 עד 12',  'משוך לעצם החזה, כתפיים אחורה'],
    ['הרמות צד עם משקולות',               '3 סטים', '10 עד 15', 'עד גובה הכתפיים, מרפק מכופף קלות'],
    ['כפיפת מרפק פטיש',                   '3 סטים', '10 עד 15', 'אחיזה ניטרלית, בלי תנופה'],
    ['הרמות רגליים בשכיבה',               '2 סטים', '12 עד 15', 'הורד איטי בלי לגעת ברצפה'],
  ]) +
  wBlock_('אימון C | דגש גב', [
    ['חתירה חד צדדית עם ספסל',            '4 סטים', '6 עד 12',  'גב ישר, משוך למותן, כיווץ בגב'],
    ['משיכת פולי תחתון לבטן',             '4 סטים', '6 עד 12',  'גב זקוף, כתפיים אחורה'],
    ['מספריים בהליכה עם משקולות',         '4 סטים', '6 עד 12',  'צעד גדול, ברך אחורית לרצפה'],
    ['לחיצת חזה במכונה',                  '4 סטים', '6 עד 12',  'גב צמוד למשענת, דחיפה מבוקרת'],
    ['לחיצת כתפיים במכונה',               '3 סטים', '6 עד 12',  'ירידה איטית, שליטה מלאה'],
    ['פשיטת מרפק מעל הראש',               '3 סטים', '10 עד 15', 'מרפקים צמודים לראש'],
    ['כפיפות בטן',                         '2 סטים', '12 עד 15', 'הרם חלק עליון, כווץ בטן'],
  ]) +
  wBlock_('אימון גיבוי ביתי | משקל גוף', [
    ['סקוואט משקל גוף',                    '3 סטים', '15 עד 25',        'ירכיים מקביל לרצפה, גב ישר'],
    ['מספריים בהליכה',                      '2 סטים', '12 עד 20 לכל רגל', 'צעד גדול, שיווי משקל'],
    ['שכיבות סמיכה',                        '3 סטים', 'מקסימום',          'גוף ישר, חזה כמעט נוגע ברצפה'],
    ['חתירה הפוכה מתחת לשולחן',             '3 סטים', 'מקסימום',          'גוף ישר, משוך חזה לשולחן'],
    ['לחיצת כתפיים הפוכה, רגליים גבוהות',  '3 סטים', '10 עד 20',          'ידיים ברוחב כתפיים, ראש למטה'],
    ['מקבילים על קצה כיסא',                '2 סטים', '12 עד 20',          'ידיים מאחור, הורד בכיפוף מרפקים'],
    ['פלאנק עם הרמות רגליים לסירוגין',      '2 סטים', 'מקסימום',          'גוף ישר, הרם רגל אחת בכל פעם'],
  ]);
}
