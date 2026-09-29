/**
 * ZoubirNasriShop — استقبال الطلبات من المتجر وتسجيلها في ورقة Orders.
 * يجب أن يكون اسم الورقة الأولى في الجدول: Orders
 */

var SECRET = "ZN"; // يجب أن يطابق التوكن المُرسَل من index.html

var HEADERS = [
  "المرجع", "التاريخ", "الزبون", "الهاتف", "الهاتف 2",
  "رمز الولاية", "الولاية", "البلدية", "العنوان", "نوع التوصيل",
  "المنتج", "الكمية", "سعر الوحدة", "التوصيل", "المجموع",
  "الحالة", "ملاحظات"
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.token !== SECRET) {
      return ContentService.createTextOutput("unauthorized");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    if (!sheet) {
      sheet = ss.insertSheet("Orders");
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }

    var rows = data.rows || [];
    rows.forEach(function (row) {
      sheet.appendRow(row);
    });

    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message);
  }
}

/** اختبار سريع يدوي من محرر Apps Script (Run → testDoPost) */
function testDoPost() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        token: SECRET,
        rows: [["ZNTEST01", new Date().toLocaleString(), "زبون تجريبي", "0500000000", "", "58", "سطيف", "عين ولمان", "عنوان تجريبي", "المنزل", "منتج تجريبي", 1, 1000, 600, 1600, "pending", ""]]
      })
    }
  };
  doPost(fake);
}
