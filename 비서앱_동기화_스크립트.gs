// ======================================================
// 한나의 비서앱 - 데이터 동기화 Apps Script (2026-05-31 업데이트)
// ======================================================

const PROP_PREFIX = 'BISSA_';

// 앱과 동일하게 유지 (app.html SYNC_KEYS와 항상 일치)
const SYNC_KEYS = [
  // 할일·루틴
  'todos','routines','catNames','oneThing','routineLog',
  // 독서
  'books','hallOfFame','giveUpShelf',
  // 하리원서
  'ortProgress','ortCovers','ortTitles','ortNums','customSeries',
  // 미디어
  'movies','cultures','youtubeUploads',
  // 창작
  'stories','drafts','writingTime','quotes',
  // 감사·반성·걱정
  'gratitude','reflections','worries',
  // 건강
  'sleepLog','phoneUsage','moods','moodEnergy','moodReason','weightLog','mealLog',
  // 자기탐색
  'selfSnapshots','futureMe','futureAdvices',
  // 메아리셋
  'echo','quarter','ceoDaily',
  // 기타
  'memos','visionBoard','mandalart',
  // 탭 설정
  'tabConfig'
];

function doGet(e) {
  var action   = e.parameter.action || 'get';
  var callback = e.parameter.callback;
  var props    = PropertiesService.getScriptProperties();

  // ── 전체 가져오기 (action=get) ──
  var data = {};
  SYNC_KEYS.forEach(function(key) {
    var val = props.getProperty(PROP_PREFIX + key);
    if (val) {
      try { data[key] = JSON.parse(val); } catch(err) {}
    }
  });
  var json = JSON.stringify(data);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'set') {
      var props = PropertiesService.getScriptProperties();
      var saved = [];
      var skipped = [];

      SYNC_KEYS.forEach(function(key) {
        if (body.data[key] === undefined) return;
        try {
          var val = JSON.stringify(body.data[key]);
          // PropertiesService 9KB 제한 처리
          if (val.length > 9000) {
            // 배열이면 최근 200개로 자름
            if (Array.isArray(body.data[key])) {
              val = JSON.stringify(body.data[key].slice(-200));
            }
          }
          // 그래도 크면 skip
          if (val.length > 9000) {
            skipped.push(key);
            return;
          }
          props.setProperty(PROP_PREFIX + key, val);
          saved.push(key);
        } catch(err) {
          skipped.push(key + '(' + err.message + ')');
        }
      });

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', saved: saved.length, skipped: skipped })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
