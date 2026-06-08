# NotebookLM 연결 설정 완료 (2026-05-28)

## 설치 상태
- **라이브러리**: notebooklm-py v0.5.0 설치됨
- **브라우저**: Playwright Chromium (`C:\playwright-browsers`)
- **인증 계정**: danmi0709@gmail.com
- **인증 저장 위치**: `C:\notebooklm\profiles\default\storage_state.json`

## 환경변수 (실행 시 필수)

| 변수명 | 값 |
|--------|-----|
| `PLAYWRIGHT_BROWSERS_PATH` | `C:\playwright-browsers` |
| `NOTEBOOKLM_HOME` | `C:\notebooklm` |
| `PYTHONUTF8` | `1` |

> **중요**: 한글 경로 문제로 위 환경변수가 없으면 실행 안 됩니다.

## 사용 방법

### 방법 1: 래퍼 스크립트 사용 (권장)
CLAUDE 폴더의 `notebooklm.ps1`을 통해 실행:
```powershell
# PowerShell에서
.\notebooklm.ps1 list                    # 노트북 목록
.\notebooklm.ps1 ask "이 내용 요약해줘"  # 질문
.\notebooklm.ps1 create "새 노트북"      # 노트북 만들기
```

### 방법 2: Claude Code에서 직접 사용
Claude에게 이렇게 말하면 됩니다:
- "notebooklm으로 위정11 노트북 요약해줘"
- "notebooklm에서 마인드맵 JSON 추출해줘"
- "유튜브 링크를 notebooklm 소스로 추가해줘"

### 방법 3: 에이전트 스킬 설치 (선택사항)
PowerShell을 직접 열어서 아래 명령어 실행:
```powershell
$env:PLAYWRIGHT_BROWSERS_PATH = "C:\playwright-browsers"
$env:NOTEBOOKLM_HOME = "C:\notebooklm"
$env:PYTHONUTF8 = "1"
notebooklm skill install --scope user --target claude
```

## 주요 명령어 모음

```powershell
# 노트북 목록 조회
notebooklm list

# 특정 노트북 선택
notebooklm use "위정11"

# 질문하기
notebooklm ask "이 챕터의 핵심 내용은?"

# 소스 추가
notebooklm source add-url "https://유튜브링크"
notebooklm source add "파일.pdf"

# 콘텐츠 생성
notebooklm generate podcast          # 팟캐스트(오디오)
notebooklm generate slide-deck       # 슬라이드덱
notebooklm generate mind-map         # 마인드맵
notebooklm generate quiz             # 퀴즈

# 다운로드
notebooklm download audio            # 오디오 다운
notebooklm download mind-map         # 마인드맵 JSON
notebooklm download slide-deck       # 슬라이드

# 노트북 요약
notebooklm summary
```

## 현재 등록된 노트북 (2026-05-28 기준)
- 위정11, 위정10, 위정7, 위정1
- 학이1718k, 학이 11\12, 학이7
- 논어, 논어1, 논어3, 논어4
- 쓰라림3
- 기타 영문 노트북 다수

이 비서한테 다음에 또 시킬 만한 일: notebooklm으로 위정 시리즈 노트북에서 마인드맵 JSON 추출해서 시각화하기
