// ============================================================
// 비서앱 동기화 - Google Apps Script
// 사용법: 아래 코드를 복사해서 Google Apps Script에 붙여넣고
//         "웹앱으로 배포" > 액세스: 모든 사용자(익명 포함)
// ============================================================

const SHEET_NAME = 'data';
const VALID_KEYS = ['todos','routines','catNames','books','hallOfFame','memos','movies','stories'];

// 가져오기 (GET) - JSONP 방식
function doGet(e) {
  const action = (e.parameter.action || '').trim();
  const callback = (e.parameter.callback || 'callback').trim();
  let result;
  if (action === 'get') {
    result = getAllData();
  } else {
    result = { error: '알 수 없는 action: ' + action };
  }
  const output = callback + '(' + JSON.stringify(result) + ')';
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// 올리기 (POST) - fetch no-cors 방식
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'set' && body.data) {
      const sheet = getOrCreateSheet();
      VALID_KEYS.forEach(key => {
        if (body.data[key] !== undefined && body.data[key] !== null) {
          setKey(sheet, key, JSON.stringify(body.data[key]));
        }
      });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllData() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < 1) return result;
  const rows = sheet.getRange(1, 1, lastRow, 2).getValues();
  rows.forEach(row => {
    const key = row[0];
    const valueStr = row[1];
    if (VALID_KEYS.includes(key) && valueStr) {
      try { result[key] = JSON.parse(valueStr); }
      catch(e) { result[key] = valueStr; }
    }
  });
  return result;
}

function setKey(sheet, key, valueStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const keys = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
    const rowIndex = keys.indexOf(key);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 1, 2).setValue(valueStr);
      return;
    }
  }
  sheet.appendRow([key, valueStr]);
}

function getOrCreateSheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (!ssId) {
    ss = SpreadsheetApp.create('비서앱_데이터');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  } else {
    try { ss = SpreadsheetApp.openById(ssId); }
    catch(e) {
      ss = SpreadsheetApp.create('비서앱_데이터');
      props.setProperty('SPREADSHEET_ID', ss.getId());
    }
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}
