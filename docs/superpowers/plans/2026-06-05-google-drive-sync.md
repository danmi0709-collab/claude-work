# Google Drive 자동 동기화 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 비서앱 데이터를 Google Drive appDataFolder에 자동 저장해 PC와 모바일 간 동기화한다.

**Architecture:** 기존 Google Calendar OAuth 토큰에 `drive.appdata` 스코프를 추가하고, `save()` 함수에 훅을 걸어 3초 디바운스 후 Drive에 업로드한다. 앱 시작 시 Drive에서 데이터를 다운로드해 로컬과 병합한다.

**Tech Stack:** Vanilla JS, Google Drive REST API v3, Google Identity Services (GSI), localStorage

---

## 파일 변경 목록

- **Modify:** `app.html` — 아래 모든 변경사항이 이 단일 파일에 적용됨

---

## Task 1: 동기화 상태 UI 추가

**Files:**
- Modify: `app.html` — 헤더 영역에 syncStatusIcon 추가

- [ ] **Step 1: 헤더에 동기화 아이콘 엘리먼트 추가**

`app.html` 헤더 div 안에서 날짜/인사말 텍스트 바로 뒤에 추가:

```html
<div id="syncStatusIcon" style="font-size:11px;color:#aaa;margin-top:2px;min-height:16px"></div>
```

헤더 구조 확인 후 `<div class="header">` 안의 마지막 자식으로 추가한다.

- [ ] **Step 2: setSyncStatus 함수 추가**

`app.html` 의 `// 🔮 미래의 나` 섹션 바로 위에 새 섹션으로 삽입:

```javascript
// ============================================================
// ☁️ Google Drive 자동 동기화
// ============================================================
let driveUploadTimer = null;

function setSyncStatus(status) {
  const el = document.getElementById('syncStatusIcon');
  if (!el) return;
  const icons   = { syncing: '🔄', synced: '☁️', offline: '⚠️' };
  const labels  = { syncing: '동기화 중', synced: '동기화됨', offline: '오프라인' };
  const colors  = { syncing: '#4285f4', synced: '#aaa', offline: '#f59e0b' };
  el.textContent = icons[status] + ' ' + labels[status];
  el.style.color = colors[status];
}
```

- [ ] **Step 3: 브라우저에서 앱 열어 헤더에 아이콘 영역이 생겼는지 확인**

http://localhost:3000/app.html 열어서 헤더 아래에 빈 공간(나중에 아이콘 들어갈 자리)이 있는지 육안 확인.

- [ ] **Step 4: 커밋**

```bash
git add app.html
git commit -m "feat: 동기화 상태 UI 아이콘 추가 (syncStatusIcon)"
```

---

## Task 2: save() 함수에 _syncMeta 훅 추가

**Files:**
- Modify: `app.html` — line ~2301의 `save` 함수

- [ ] **Step 1: 기존 save 함수 위치 확인**

app.html 에서 다음 코드를 찾는다 (line ~2301):

```javascript
const save = (key, val) => {
  try { localStorage.setItem(key, JSON.stringify(val)); }
  catch(e) { if (e.name === 'QuotaExceededError' || e.code === 22) showToast('저장 공간이 부족해요. 사진 크기를 줄여보세요 📦'); }
};
```

- [ ] **Step 2: save 함수를 _syncMeta 업데이트 + scheduleUpload 호출로 교체**

```javascript
const save = (key, val) => {
  try {
    localStorage.setItem(key, JSON.stringify(val));
    // _syncMeta 자체 저장 시 재귀 방지
    if (key !== '_syncMeta' && key !== 'driveFileId') {
      const meta = JSON.parse(localStorage.getItem('_syncMeta') || '{}');
      meta[key] = Date.now();
      localStorage.setItem('_syncMeta', JSON.stringify(meta));
      scheduleUpload();
    }
  }
  catch(e) { if (e.name === 'QuotaExceededError' || e.code === 22) showToast('저장 공간이 부족해요. 사진 크기를 줄여보세요 📦'); }
};
```

- [ ] **Step 3: scheduleUpload 함수 추가 (Task 1에서 만든 Drive 섹션 안에)**

```javascript
function scheduleUpload() {
  if (!gcalToken) return; // 구글 연결 안 됐으면 무시
  if (driveUploadTimer) clearTimeout(driveUploadTimer);
  driveUploadTimer = setTimeout(uploadToDrive, 3000);
}
```

- [ ] **Step 4: 앱 열고 콘솔에서 동작 확인**

브라우저 개발자 도구 콘솔에서:
```javascript
save('testKey', 'hello');
JSON.parse(localStorage.getItem('_syncMeta'))
// → { testKey: <타임스탬프> } 형태가 나와야 함
localStorage.removeItem('testKey');
```

- [ ] **Step 5: 커밋**

```bash
git add app.html
git commit -m "feat: save()에 _syncMeta 타임스탬프 훅 + scheduleUpload 연결"
```

---

## Task 3: Drive 업로드 함수 구현

**Files:**
- Modify: `app.html` — Drive 섹션에 uploadToDrive 추가

- [ ] **Step 1: uploadToDrive 함수 추가 (Drive 섹션 안에)**

```javascript
async function uploadToDrive() {
  if (!gcalToken) return;
  setSyncStatus('syncing');
  try {
    // 데이터 수집 (SYNC_KEYS 전체)
    const payload = {};
    SYNC_KEYS.forEach(k => {
      const v = load(k, null);
      if (v !== null) payload[k] = v;
    });
    payload._syncMeta = JSON.parse(localStorage.getItem('_syncMeta') || '{}');

    const fileId = load('driveFileId', null);
    const body = JSON.stringify(payload);

    if (fileId) {
      // 기존 파일 업데이트
      const res = await fetch(
        `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': 'Bearer ' + gcalToken,
            'Content-Type': 'application/json'
          },
          body
        }
      );
      if (!res.ok) throw new Error('upload failed: ' + res.status);
    } else {
      // 새 파일 생성 (multipart)
      const meta = JSON.stringify({ name: 'bissa-data.json', parents: ['appDataFolder'] });
      const boundary = '-------bissa_boundary';
      const delimiter = '\r\n--' + boundary + '\r\n';
      const closeDelimiter = '\r\n--' + boundary + '--';
      const multipartBody =
        delimiter + 'Content-Type: application/json\r\n\r\n' + meta +
        delimiter + 'Content-Type: application/json\r\n\r\n' + body +
        closeDelimiter;

      const res = await fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
        {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer ' + gcalToken,
            'Content-Type': 'multipart/related; boundary=' + boundary
          },
          body: multipartBody
        }
      );
      if (!res.ok) throw new Error('create failed: ' + res.status);
      const data = await res.json();
      if (data.id) {
        localStorage.setItem('driveFileId', JSON.stringify(data.id));
      }
    }
    setSyncStatus('synced');
  } catch(e) {
    console.warn('Drive 업로드 실패:', e);
    setSyncStatus('offline');
  }
}
```

- [ ] **Step 2: 커밋**

```bash
git add app.html
git commit -m "feat: uploadToDrive 구현 (신규 파일 생성 + 업데이트)"
```

---

## Task 4: Drive 다운로드 + 병합 함수 구현

**Files:**
- Modify: `app.html` — Drive 섹션에 downloadAndMerge 추가

- [ ] **Step 1: ARRAY_KEYS 상수 정의 (Drive 섹션 상단에)**

```javascript
const DRIVE_ARRAY_KEYS = new Set([
  'routines','books','memos','movies','stories','drafts',
  'gratitude','reflections','todos','routineLog','futureAdvices','selfSnapshots'
]);
```

- [ ] **Step 2: downloadAndMerge 함수 추가**

```javascript
async function downloadAndMerge() {
  if (!gcalToken) return;
  setSyncStatus('syncing');
  try {
    // 1. 파일 ID 찾기
    let fileId = load('driveFileId', null);
    if (!fileId) {
      const searchRes = await fetch(
        "https://www.googleapis.com/drive/v3/files?spaces=appDataFolder&q=name%3D'bissa-data.json'&fields=files(id)",
        { headers: { 'Authorization': 'Bearer ' + gcalToken } }
      );
      if (!searchRes.ok) throw new Error('search failed');
      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        fileId = searchData.files[0].id;
        localStorage.setItem('driveFileId', JSON.stringify(fileId));
      } else {
        // 첫 사용 — 로컬 데이터를 Drive에 올리기만 함
        setSyncStatus('synced');
        await uploadToDrive();
        return;
      }
    }

    // 2. 파일 다운로드
    const res = await fetch(
      `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
      { headers: { 'Authorization': 'Bearer ' + gcalToken } }
    );
    if (!res.ok) throw new Error('download failed: ' + res.status);
    const remote = await res.json();

    // 3. 병합
    const localMeta  = JSON.parse(localStorage.getItem('_syncMeta') || '{}');
    const remoteMeta = remote._syncMeta || {};

    SYNC_KEYS.forEach(key => {
      const localVal  = load(key, null);
      const remoteVal = remote[key];
      if (remoteVal === undefined) return; // 원격에 없으면 패스

      const localTime  = localMeta[key]  || 0;
      const remoteTime = remoteMeta[key] || 0;

      if (DRIVE_ARRAY_KEYS.has(key) && Array.isArray(remoteVal)) {
        // 배열: id 기준 union, 충돌 시 더 최신 타임스탬프 쪽 우선
        const base    = Array.isArray(localVal) ? localVal : [];
        const winner  = remoteTime >= localTime ? remoteVal : base;
        const loser   = remoteTime >= localTime ? base : remoteVal;
        const byId    = {};
        loser.forEach(item  => { if (item && item.id != null) byId[item.id] = item; });
        winner.forEach(item => { if (item && item.id != null) byId[item.id] = item; });
        // id 없는 항목도 보존 (winner 기준)
        const noId = winner.filter(item => item == null || item.id == null);
        save(key, [...Object.values(byId), ...noId]);
      } else {
        // 단순값 / 날짜별 객체: 더 최신 것 사용
        if (remoteTime > localTime) {
          save(key, remoteVal);
        }
      }
    });

    setSyncStatus('synced');
    showToast('☁️ 동기화 완료!');
  } catch(e) {
    console.warn('Drive 다운로드 실패:', e);
    setSyncStatus('offline');
  }
}
```

- [ ] **Step 3: 커밋**

```bash
git add app.html
git commit -m "feat: downloadAndMerge 구현 (배열 union 병합 + 날짜별 최신 우선)"
```

---

## Task 5: 인증 스코프 확장 + 앱 시작 시 동기화 연결

**Files:**
- Modify: `app.html` — gcalAutoConnect, setGcalConnected 함수

- [ ] **Step 1: gcalAutoConnect의 scope에 drive.appdata 추가**

기존:
```javascript
scope: 'https://www.googleapis.com/auth/calendar.readonly',
```

변경 후:
```javascript
scope: 'https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/drive.appdata',
```

- [ ] **Step 2: setGcalConnected에 downloadAndMerge 호출 추가**

기존 `setGcalConnected` 함수 끝부분 `fetchCalendarEvents(); renderSchedGcal();` 바로 뒤에 추가:

```javascript
// Drive 동기화 시작 (다운로드 후 병합)
downloadAndMerge().catch(() => setSyncStatus('offline'));
```

- [ ] **Step 3: 캘린더 연결 안내문에 Drive 언급 추가**

`app.html`에서 캘린더 설정 안내 HTML 영역을 찾아 기존 안내 텍스트 끝에 한 줄 추가:

```html
<p style="font-size:12px;color:#888;margin-top:8px">🔒 캘린더 읽기 + Drive 자동 저장 권한이 요청됩니다.</p>
```

- [ ] **Step 4: 브라우저에서 수동 테스트**

1. http://localhost:3000/app.html 열기
2. 캘린더 탭 → 구글 캘린더 연결 버튼 클릭
3. 권한 요청 팝업에 `calendar.readonly`와 `drive.appdata` 두 권한이 표시되는지 확인
4. 연결 후 헤더에 "☁️ 동기화됨" 표시 확인
5. 브라우저 콘솔에 Drive 관련 에러 없는지 확인

- [ ] **Step 5: 커밋**

```bash
git add app.html
git commit -m "feat: Drive 스코프 추가 + 연결 시 자동 다운로드 병합 시작"
```

---

## Task 6: 기존 캘린더 연결자 재인증 안내 + 마무리

**Files:**
- Modify: `app.html` — 기존 연결된 사용자 토스트 안내

- [ ] **Step 1: gcalAutoConnect 자동 연결 성공 후 Drive 초기화 호출 확인**

`gcalAutoConnect` 의 `callback` 안에서 `setGcalConnected(resp.access_token)` 이 호출되고,
`setGcalConnected` 에서 `downloadAndMerge()` 가 호출되므로 자동 연결 시에도 Drive 동기화가 시작된다.
별도 수정 불필요 — 코드 흐름만 확인.

- [ ] **Step 2: 앱 배포 (Vercel)**

```bash
vercel deploy --prod
```

- [ ] **Step 3: 모바일에서 테스트**

1. PC에서 데이터 몇 개 입력 (메모, 루틴 체크 등) → 헤더에 "☁️ 동기화됨" 확인
2. 모바일에서 https://bissa-peach.vercel.app 열기
3. 구글 캘린더 연결 → 헤더에 "동기화 완료!" 토스트 확인
4. PC에서 입력한 데이터가 모바일에 보이는지 확인

- [ ] **Step 4: 최종 커밋**

```bash
git add app.html
git commit -m "feat: Google Drive 자동 동기화 완성 (PC↔모바일 자동 병합)"
```
