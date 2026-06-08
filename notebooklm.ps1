# NotebookLM 실행 래퍼 - 환경변수 자동 설정
# 사용법: .\notebooklm.ps1 list
#         .\notebooklm.ps1 ask "질문 내용"
#         .\notebooklm.ps1 login

$env:PLAYWRIGHT_BROWSERS_PATH = "C:\playwright-browsers"
$env:NOTEBOOKLM_HOME = "C:\notebooklm"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

notebooklm @args
