var SPREADSHEET_ID = '189_NrjhdDdUjUD53ZuEQtyL7GzOEFJugwx5wrXducdw';
var SHEET_NAME = 'نتائج المدارس';

// ----------------------------------------------------------------
// doPost: تستقبل بيانات مدرسة جديدة
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    // ✅ منع التكرار: تحقق من كود المدرسة قبل الحفظ
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][1]) === String(data.school_id)) {
        return jsonResponse({ success: false, duplicate: true,
          message: 'هذه المدرسة قامت برفع بياناتها مسبقاً' });
      }
    }

    // التحقق من صحة البيانات
    var v = validateData(data);
    if (!v.ok) return jsonResponse({ success: false, error: v.message });

    // حفظ البيانات
    sheet.appendRow([
      new Date(),
      data.school_id,
      data.school_name,
      data.school_type || '',
      Number(data.g1_total),
      Number(data.g1_passed),
      Number(data.g1_remedial),
      Number(data.g2_total),
      Number(data.g2_passed),
      Number(data.g2_remedial)
    ]);

    return jsonResponse({ success: true });

  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ----------------------------------------------------------------
// doGet: يرجع قائمة المدارس التي أرسلت بالفعل (للـ frontend)
// ----------------------------------------------------------------
function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    // action=submitted → يرجع فقط كودات المدارس المرسلة
    var submittedIds = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1]) submittedIds.push(String(rows[i][1]));
    }
    return jsonResponse({ success: true, submitted: submittedIds });

  } catch(err) {
    return jsonResponse({ success: false, error: err.message, submitted: [] });
  }
}

// ----------------------------------------------------------------
// دوال مساعدة
// ----------------------------------------------------------------
function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['تاريخ الإرسال','كود المدرسة','اسم المدرسة','نوع التعليم',
                   'إجمالي أول','ناجح أول','علاجي أول',
                   'إجمالي ثاني','ناجح ثاني','علاجي ثاني'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
         .setFontWeight('bold')
         .setBackground('#0F6E56')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 130);
  }
  return sheet;
}

function validateData(data) {
  if (!data.school_name || data.school_name.length < 2)
    return { ok: false, message: 'اسم المدرسة غير صحيح' };
  var nums = ['g1_total','g1_passed','g1_remedial','g2_total','g2_passed','g2_remedial'];
  for (var i = 0; i < nums.length; i++) {
    if (isNaN(Number(data[nums[i]])))
      return { ok: false, message: 'بيانات رقمية غير صحيحة' };
  }
  return { ok: true };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
