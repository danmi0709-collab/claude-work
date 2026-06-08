# Google Drive 자동 동기화 설계

**날짜:** 2026-06-05  
**목표:** 비서앱 데이터를 PC와 모바일 간 자동으로 동기화한다.

---

## 1. 전체 구조

- 앱 데이터 전체를 Google Drive **appDataFolder**(앱 전용 숨김 폴더)에 `bissa-data.json` 파일 하나로 저장
- 앱 열 때: Drive → 다운로드 → 로컬과 병합 → 적용
- 데이터 바뀔 때: 3초 디바운스 후 Drive에 자동 업로드
- 구글 캘린더 연결 시 Drive 권한도 동시 요청 (추가 로그인 없음)

## 2. 인증 변경

- 기존 `gcalAutoConnect` / `handleGcalAuth` 의 OAuth scope에 `https://www.googleapis.com/auth/drive.appdata` 추가
- 기존 `gcalToken` 을 Drive API 호출에도 재사용 (같은 토큰)
- 기존에 캘린더를 연결해둔 사용자는 재연결 시 새 권한 자동 요청됨

## 3. 업로드 흐름

1. 기존 `save(key, value)` 함수 호출 시 dirty 플래그 세팅 + `_syncMeta[key] = Date.now()` 갱신
2. 3초 디바운스 타이머 시작 (이미 타이머 돌면 리셋)
3. 타이머 만료 시: 전체 localStorage 데이터 + `_syncMeta` 수집 → JSON 직렬화 → Drive에 PUT
4. Drive 파일 ID는 로컬에 `driveFileId`로 저장 (첫 업로드 시 생성, 이후 업데이트)

## 4. 다운로드 및 병합 흐름

앱 시작 시 (gcalToken 준비된 후):
1. Drive에서 `bissa-data.json` 다운로드
2. 로컬 데이터와 키별 병합:
   - **배열형 키** (routines, books, memos, movies, stories, drafts 등): 아이템 `id` 기준 union — 한쪽에만 있는 항목은 무조건 추가. 같은 `id`가 양쪽에 있으면 해당 키의 `_syncMeta[key]` 타임스탬프가 더 최신인 기기의 항목을 사용
   - **날짜별 객체형 키** (sleepLog, weightLog, moods, moodEnergy 등): 날짜 키별로 `_syncMeta` 비교해 최신 것 사용
   - **단순 값** (설정, 프로필 등): `_syncMeta` 타임스탬프 비교해 최신 것 사용
3. 병합 결과를 localStorage에 저장 → 현재 탭 리렌더

## 5. UI 변경

- **헤더 오른쪽 상단**에 작은 동기화 아이콘 추가:
  - ☁️ `동기화됨` (회색, 평상시)
  - 🔄 `동기화 중` (파란색, 업로드/다운로드 중)
  - ⚠️ `오프라인` (주황색, Drive 연결 안 됨)
- 캘린더 연결 안내문에 "Google Drive에도 자동 저장됩니다" 한 줄 추가

## 6. 오류 처리

- Drive 요청 실패 시: 로컬 데이터 유지, 재시도 없이 아이콘만 ⚠️로 변경
- 토큰 만료 시: 기존 캘린더 재연결 플로우와 동일하게 자동 재시도
- Drive 파일 없음(첫 사용): 새 파일 생성으로 처리

## 7. 배열형 키 목록 (병합 대상)

`routines`, `books`, `memos`, `movies`, `stories`, `drafts`, `gratitude`, `reflections`, `todos`, `routineLog`, `futureAdvices`, `selfSnapshots`

## 8. 단순값/날짜별 키 목록

`sleepLog`, `weightLog`, `moods`, `moodEnergy`, `moodReason`, `phoneUsage`, `mealLog`, `writingTime`, `futureMe`, `claudeApiKey`, `gcalClientId`, `driveFileId`, `_syncMeta`
