// ======================================================
// 한나의 비서앱 - 가계부 구글 시트 연동 Apps Script
// ======================================================
// 사용 방법:
// 1. 구글 시트 새로 만들기
// 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 저장(Ctrl+S) 후 → 배포 → 새 배포
// 4. 유형: 웹 앱 / 액세스: 모든 사용자 → 배포
// 5. 생성된 URL을 비서앱 가계부 탭에 붙여넣고 URL 저장
// ======================================================

const SHEET_NAME = '가계부';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['id','date','type','cat','amount','memo']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'get') {
    return ContentService
      .createTextOutput(JSON.stringify(getAllEntries()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'set') {
      saveAllEntries(body.data);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', count: body.data.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'add') {
      addEntry(body.entry);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllEntries() {
  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    id:     r[0],
    date:   r[1],
    type:   r[2],
    cat:    r[3],
    amount: r[4],
    memo:   r[5]
  }));
}

function saveAllEntries(data) {
  const sheet = getOrCreateSheet();
  // 헤더 제외하고 전체 초기화
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (data.length === 0) return;
  const rows = data.map(e => [e.id, e.date, e.type, e.cat, e.amount, e.memo || '']);
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

function addEntry(entry) {
  const sheet = getOrCreateSheet();
  sheet.appendRow([entry.id, entry.date, entry.type, entry.cat, entry.amount, entry.memo || '']);
}
