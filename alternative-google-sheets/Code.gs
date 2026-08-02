/**
 * دليل العسيرات — بديل اختياري: استقبال البيانات في Google Sheet بدل GitHub
 * (استخدم هذا فقط لو قررت لاحقًا إنك ميتوكل مش على GitHub API، ده أأمن تقنيًا
 * لأن مفتاح الوصول بيفضل على سيرفر جوجل ومايظهرش في كود الموقع أبدًا)
 * راجع تعليمات التركيب في تعليقات الكود بتاع هذا الملف في المحادثة السابقة،
 * أو اسأل في أي وقت لو حبيت ترجع لهذه الطريقة.
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
    Object.keys(data).forEach(key => { if (headers.indexOf(key) === -1) headers.push(key); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const row = headers.map(h => Array.isArray(data[h]) ? data[h].join(', ') : (data[h] !== undefined ? data[h] : ''));
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
function doGet(e) {
  return ContentService.createTextOutput('دليل العسيرات — نقطة استقبال البيانات تعمل بنجاح ✅');
}
