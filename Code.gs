// ============================================================
// بوابة المدارس 2026 — Google Apps Script Backend
// ============================================================

var SPREADSHEET_ID = '189_NrjhdDdUjUD53ZuEQtyL7GzOEFJugwx5wrXducdw';
var SHEETS = { AUTH: 'مصادقة', DATA: 'البيانات' };

// ---- نقطة الدخول: POST ----
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case 'login':          return handleLogin(d);
      case 'submit':         return handleSubmit(d);
      case 'update':         return handleUpdate(d);
      case 'changePassword': return handleChangePassword(d);
      default: return jsonResponse({ success: false, error: 'إجراء غير معروف' });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ---- نقطة الدخول: GET ----
function doGet(e) {
  var action    = e.parameter.action;
  var school_id = e.parameter.school_id;
  try {
    switch (action) {
      case 'getData':    return getSchoolData(school_id);
      case 'submitted':  return getSubmittedList();
      case 'aggregate':  return getAggregate();
      default: return jsonResponse({ success: false });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ============================================================
// المصادقة
// ============================================================

function handleLogin(d) {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var authSheet = getOrCreateAuthSheet(ss);
  var rows      = authSheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(d.school_id).trim()) {
      // تحويل كلاهما لـ String قبل المقارنة (Sheets يخزن الأرقام كـ number)
      if (String(rows[i][1]).trim() === String(d.password).trim()) {
        authSheet.getRange(i + 1, 3).setValue(new Date().toISOString());
        return jsonResponse({ success: true, firstLogin: false });
      }
      return jsonResponse({ success: false, error: 'كلمة المرور غير صحيحة' });
    }
  }

  // أول تسجيل دخول → الكود هو كلمة المرور الافتراضية
  if (String(d.password).trim() === String(d.school_id).trim()) {
    // نخزّن كلمة المرور كـ نص صريح بإضافة apostrophe prefix
    authSheet.appendRow(["'" + d.school_id, "'" + d.school_id, new Date().toISOString()]);
    return jsonResponse({ success: true, firstLogin: true });
  }
  return jsonResponse({ success: false, error: 'كلمة المرور غير صحيحة' });
}

function handleChangePassword(d) {
  if (!validateAuth(d.school_id, d.oldPassword))
    return jsonResponse({ success: false, error: 'كلمة المرور الحالية غير صحيحة' });

  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var authSheet = getOrCreateAuthSheet(ss);
  var rows      = authSheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(d.school_id).trim()) {
      // نخزّن كلمة المرور الجديدة كـ نص صريح
      authSheet.getRange(i + 1, 2).setValue("'" + d.newPassword);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ success: false, error: 'المدرسة غير موجودة' });
}

// ============================================================
// حفظ البيانات (جديد) مع قفل مانع للتضارب
// ============================================================

function handleSubmit(d) {
  if (!validateAuth(d.school_id, d.password))
    return jsonResponse({ success: false, error: 'غير مصرح' });

  var lock    = LockService.getScriptLock();
  var hasLock = lock.tryLock(15000); // ننتظر 15 ثانية قبل الاستسلام
  if (!hasLock)
    return jsonResponse({ success: false, locked: true, message: 'الخادم مشغول، سيتم إعادة المحاولة تلقائياً…' });

  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = getOrCreateDataSheet(ss);
    var rows      = dataSheet.getDataRange().getValues();

    // منع التكرار
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(d.school_id))
        return jsonResponse({ success: false, alreadyExists: true });
    }

    dataSheet.appendRow(buildRow(d, '1'));
    dataSheet.appendRow(buildRow(d, '2'));
    SpreadsheetApp.flush(); // إجبار الكتابة الفورية
    return jsonResponse({ success: true });

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// تعديل البيانات مع قفل
// ============================================================

function handleUpdate(d) {
  if (!validateAuth(d.school_id, d.password))
    return jsonResponse({ success: false, error: 'غير مصرح' });

  var lock    = LockService.getScriptLock();
  var hasLock = lock.tryLock(15000);
  if (!hasLock)
    return jsonResponse({ success: false, locked: true, message: 'الخادم مشغول، سيتم إعادة المحاولة تلقائياً…' });

  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = getOrCreateDataSheet(ss);
    var rows      = dataSheet.getDataRange().getValues();
    var found     = { '1': false, '2': false };

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(d.school_id)) {
        var gr   = String(rows[i][4]);
        var newR = buildRow(d, gr);
        dataSheet.getRange(i + 1, 1, 1, newR.length).setValues([newR]);
        found[gr] = true;
      }
    }
    // أضف الصف لو غائب (حالة نادرة)
    if (!found['1']) dataSheet.appendRow(buildRow(d, '1'));
    if (!found['2']) dataSheet.appendRow(buildRow(d, '2'));

    SpreadsheetApp.flush();
    return jsonResponse({ success: true });

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// استرجاع البيانات
// ============================================================

function getSchoolData(school_id) {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = getOrCreateDataSheet(ss);
  var rows      = dataSheet.getDataRange().getValues();
  var headers   = rows[0];
  var found     = [];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(school_id)) {
      var obj = {};
      headers.forEach(function (h, idx) { obj[h] = rows[i][idx]; });
      found.push(obj);
    }
  }
  return jsonResponse({ success: true, exists: found.length > 0, data: found });
}

function getSubmittedList() {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = getOrCreateDataSheet(ss);
  var rows      = dataSheet.getDataRange().getValues();
  var ids       = {};
  for (var i = 1; i < rows.length; i++)
    if (rows[i][1]) ids[String(rows[i][1])] = true;
  return jsonResponse({ success: true, submitted: Object.keys(ids) });
}

function getAggregate() {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = getOrCreateDataSheet(ss);
  var rows      = dataSheet.getDataRange().getValues();
  var headers   = rows[0];
  var numStart  = 5; // أول عمود رقمي (بعد الحقول التعريفية)
  var agg       = {}; // {grade: {type: {field: value, count: n}}}

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    var grade = String(rows[i][4]);
    var type  = String(rows[i][3]);
    if (!agg[grade]) agg[grade] = {};
    if (!agg[grade][type]) {
      agg[grade][type] = { count: 0 };
      headers.slice(numStart).forEach(function (h) { agg[grade][type][h] = 0; });
    }
    agg[grade][type].count++;
    headers.slice(numStart).forEach(function (h, idx) {
      agg[grade][type][h] += Number(rows[i][numStart + idx]) || 0;
    });
  }
  return jsonResponse({ success: true, aggregate: agg });
}

// ============================================================
// دوال مساعدة
// ============================================================

function validateAuth(school_id, password) {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var authSheet = getOrCreateAuthSheet(ss);
  var rows      = authSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(school_id).trim())
      return String(rows[i][1]).trim() === String(password).trim();
  }
  // لم يُسجَّل بعد → الكود هو الباسوورد الافتراضي
  return String(password).trim() === String(school_id).trim();
}

// بناء صف واحد للبيانات (صف دراسي محدد)
function buildRow(d, grade) {
  var g = 'g' + grade + '_';
  var n = function (k) { return Number(d[g + k]) || 0; };

  return [
    new Date().toISOString(), d.school_id, d.school_name, d.school_type, grade,
    // استمارة 1
    n('enrolled_m'),   n('enrolled_f'),
    n('pct60_m'),      n('pct60_f'),
    n('init_pres_m'),  n('init_pres_f'),
    n('init_pass_m'),  n('init_pass_f'),
    n('init_fail_m'),  n('init_fail_f'),
    n('init_abs_m'),   n('init_abs_f'),
    n('fin_pres_m'),   n('fin_pres_f'),
    n('fin_pass_m'),   n('fin_pass_f'),
    n('fin_fail_m'),   n('fin_fail_f'),
    n('fin_abs_m'),    n('fin_abs_f'),
    // استمارة 2
    n('rel_pres'),  n('rel_pass'),  n('rel_fail'),
    n('ara_pres'),  n('ara_pass'),  n('ara_fail'),
    n('eng_pres'),  n('eng_pass'),  n('eng_fail'),
    n('math_pres'), n('math_pass'), n('math_fail')
  ];
}

function getOrCreateAuthSheet(ss) {
  var s = ss.getSheetByName(SHEETS.AUTH);
  if (!s) {
    s = ss.insertSheet(SHEETS.AUTH);
    s.appendRow(['كود المدرسة', 'كلمة المرور', 'آخر دخول']);
    s.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0F6E56').setFontColor('#fff');
    s.setFrozenRows(1);
  }
  return s;
}

function getOrCreateDataSheet(ss) {
  var s = ss.getSheetByName(SHEETS.DATA);
  if (!s) {
    s = ss.insertSheet(SHEETS.DATA);
    var h = [
      'تاريخ_الإرسال','كود_المدرسة','اسم_المدرسة','نوع_التعليم','الصف',
      'مقيدين_ذ','مقيدين_ي',
      'حاضرين60_ذ','حاضرين60_ي',
      'مبدئي_حاضرين_ذ','مبدئي_حاضرين_ي',
      'مبدئي_اجتازوا_ذ','مبدئي_اجتازوا_ي',
      'مبدئي_لم_يجتازوا_ذ','مبدئي_لم_يجتازوا_ي',
      'مبدئي_غائب_ذ','مبدئي_غائب_ي',
      'نهائي_حاضرين_ذ','نهائي_حاضرين_ي',
      'نهائي_اجتازوا_ذ','نهائي_اجتازوا_ي',
      'نهائي_لم_يجتازوا_ذ','نهائي_لم_يجتازوا_ي',
      'نهائي_غائب_ذ','نهائي_غائب_ي',
      'ديني_حاضر','ديني_اجتازوا','ديني_لم_يجتازوا',
      'عربي_حاضر','عربي_اجتازوا','عربي_لم_يجتازوا',
      'انجليزي_حاضر','انجليزي_اجتازوا','انجليزي_لم_يجتازوا',
      'رياضيات_حاضر','رياضيات_اجتازوا','رياضيات_لم_يجتازوا'
    ];
    s.appendRow(h);
    s.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#0F6E56').setFontColor('#fff');
    s.setFrozenRows(1);
    s.setColumnWidths(1, h.length, 120);
  }
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
