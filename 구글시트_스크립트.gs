// =============================================
// 한나의 가계부 - Google Apps Script
// 구글 시트에 붙여넣고 웹앱으로 배포하세요
// =============================================

const SHEET_NAME = '가계부데이터';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    let data;
    if (!sheet || sheet.getLastRow() < 3) {
      data = { transactions: [] };
    } else {
      const raw = sheet.getRange(3, 1).getValue();
      data = raw ? JSON.parse(raw) : { transactions: [] };
    }
    // JSONP 지원 (file:// 프로토콜 CORS 우회)
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return response(data);
  } catch (err) {
    return response({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    sheet.clearContents();
    sheet.getRange(1, 1).setValue('최종업데이트');
    sheet.getRange(1, 2).setValue('거래수');
    sheet.getRange(2, 1).setValue(new Date().toLocaleString('ko-KR'));

    // 폼 제출(parameter.data) 또는 JSON body(postData.contents) 모두 지원
    let raw = '';
    if (e.parameter && e.parameter.data) {
      raw = e.parameter.data;
    } else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }

    const data = JSON.parse(raw);
    sheet.getRange(2, 2).setValue(data.transactions ? data.transactions.length : 0);

    // 데이터 JSON을 3행 1열에 저장
    sheet.getRange(3, 1).setValue(raw);

    return response({ status: 'ok', count: data.transactions ? data.transactions.length : 0 });
  } catch (err) {
    return response({ error: err.toString() });
  }
}

function response(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
