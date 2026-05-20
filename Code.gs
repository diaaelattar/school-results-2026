// ============================================================
// Google Apps Script — بوابة نتائج المدارس 2026
// ارفع هذا الكود في Google Apps Script وانشره كـ Web App
// ============================================================

// معرّف الـ Spreadsheet — ستضعه أنت بعد إنشاء الـ Sheet
var SPREADSHEET_ID = 'ضع_هنا_معرف_الـ_Sheet';
var SHEET_NAME = 'نتائج المدارس';

// ----------------------------------------------------------------
// doPost: يُستدعى لما مدرسة ترسل بياناتها
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // التحقق من البيانات قبل الحفظ
    var validation = validateData(data);
    if (!validation.ok) {
      return jsonResponse({ success: false, error: validation.message });
    }

    var sheet = getOrCreateSheet();

    // إضافة صف جديد بالبيانات
    sheet.appendRow([
      new Date(),                    // تاريخ ووقت الإرسال
      data.school_name,              // اسم المدرسة
      data.g1_total,                 // إجمالي الصف الأول
      data.g1_passed,                // ناجحو الصف الأول
      data.g1_remedial,              // علاجي الصف الأول
      data.g2_total,                 // إجمالي الصف الثاني
      data.g2_passed,                // ناجحو الصف الثاني
      data.g2_remedial               // علاجي الصف الثاني
    ]);

    return jsonResponse({ success: true, message: 'تم الحفظ بنجاح' });

  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ----------------------------------------------------------------
// doGet: يُستدعى لما لوحة الإدارة تطلب كل البيانات
// ----------------------------------------------------------------
function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    // أول صف هو العناوين، نتخطاه
    var headers = rows[0];
    var data = rows.slice(1).map(function(row) {
      return {
        submitted_at: row[0],
        school_name:  row[1],
        g1_total:     row[2],
        g1_passed:    row[3],
        g1_remedial:  row[4],
        g2_total:     row[5],
        g2_passed:    row[6],
        g2_remedial:  row[7]
      };
    });

    return jsonResponse({ success: true, data: data });

  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ----------------------------------------------------------------
// دوال مساعدة
// ----------------------------------------------------------------

// يجيب الـ Sheet الموجودة أو ينشئها لو مش موجودة
function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // إضافة صف العناوين بتنسيق جميل
    var headers = ['تاريخ الإرسال','اسم المدرسة',
                   'إجمالي أول','ناجح أول','علاجي أول',
                   'إجمالي ثاني','ناجح ثاني','علاجي ثاني'];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold')
         .setBackground('#0F6E56').setFontColor('#ffffff');
    sheet.setFrozenRows(1); // تثبيت صف العناوين
  }
  return sheet;
}

// التحقق من صحة البيانات المُرسلة
function validateData(data) {
  if (!data.school_name || data.school_name.length < 3)
    return { ok: false, message: 'اسم المدرسة غير صحيح' };
  var fields = ['g1_total','g1_passed','g1_remedial','g2_total','g2_passed','g2_remedial'];
  for (var i = 0; i < fields.length; i++) {
    if (isNaN(Number(data[fields[i]])) || data[fields[i]] === '')
      return { ok: false, message: 'بيانات رقمية غير صحيحة' };
  }
  return { ok: true };
}

// تحويل الاستجابة لـ JSON مع headers صحيحة
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
