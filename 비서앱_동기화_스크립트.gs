// ======================================================
// 한나의 비서앱 - 데이터 동기화 Apps Script
// ======================================================
// 사용 방법:
// 1. 구글 시트 새로 만들기
// 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 저장(Ctrl+S) 후 → 배포 → 새 배포
// 4. 유형: 웹 앱 / 액세스: 모든 사용자 → 배포
// 5. 생성된 URL을 비서앱 캘린더 탭 → 데이터 동기화에 붙여넣고 URL 저장
// ======================================================

const PROP_KEY = 'BISSA_DATA';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'get') {
    const stored = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
    const data = stored ? JSON.parse(stored) : {};
    return ContentService
      .createTextOutput(JSON.stringify(data))
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
      PropertiesService.getScriptProperties().setProperty(
        PROP_KEY,
        JSON.stringify(body.data)
      );
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
