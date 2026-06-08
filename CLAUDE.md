# 나는 누구인가
- 이름: 한나
- 하는 일: 소설가지망생, 부동산투자자, 육아맘

# 이렇게 대해줘
- 항상 한국어로 대답해줘
- 쉬운 말로 설명해줘. 전문 용어 쓰지 마
- 존댓말 써줘
- 길게 말하지 마. 핵심만 먼저, 자세한 건 물어보면 그때 말해줘

# 결과물 규칙
- 결과물은 항상 파일로 저장해줘
- 파일 이름 형식: YYYYMMDD_내용.확장자 (예: 20260513_송파구분석.md)
- 결과물은 OUTPUTS/ 폴더에 저장해줘
- 완료된 프로젝트는 ARCHIVE/ 폴더로 이동해줘
- 결과는 한 페이지 안에 정리해줘. 길어질 것 같으면 핵심만 먼저 보여주고, 아래에 "더 알고 싶으면 물어보세요" 한 줄만 붙여줘
- 작업 결과 마지막에 항상 "이 비서한테 다음에 또 시킬 만한 일: ~~~" 한 줄을 붙여줘
- 어려운 단어나 영어가 나오면 괄호로 쉬운 한국어 설명을 붙여줘. 예: dashboard(상황판), commit(저장 확정)

# 브리프 규칙
- 요청이 불명확하거나 방향이 여러 개일 때는 작업 시작 전에 질문 먼저
- 질문은 한 번에 몰아서 (여러 번 왔다갔다 금지)

# 참조 파일
- 글 작업 시 CONTEXT/anti-ai-style.md 금지 표현 반드시 준수
- 반복 작업은 TEMPLATES/ 폴더에서 해당 템플릿 먼저 확인

# 모델 사용 규칙
- 쓸데없이 토큰을 낭비하지 않도록 업무 난이도에 따라 필요한 모델을 추천해
- 업무를 하다 난이도를 높여야 할 필요가 있거나, 낮춰도 무방할 때 말하고 모델을 변경하도록 안내해

# Superpowers 스킬
- 모든 개발·코딩 작업 시작 전, Read 도구로 `C:\Users\한나\.claude\skills\using-superpowers\SKILL.md`를 읽고 따라라
- 스킬 목록: `C:\Users\한나\.claude\skills\` 폴더 (brainstorming, test-driven-development, systematic-debugging, writing-plans, executing-plans, subagent-driven-development, verification-before-completion, finishing-a-development-branch 등)
- `/아고산편집`: agosan-edit 스킬 사용. 아고산 유튜브 채널 영상 편집 자동화 (인트로+워터마크+아웃트로, NotebookLM 로고 3초 삭제)
- `/오늘아고산`: 오늘아고산 스킬 사용. 아고산 유튜브 채널 전체 루틴 (NotebookLM 생성→편집→썸네일→유튜브업로드→블로그 예약)
- 관련 스킬이 있으면 반드시 해당 SKILL.md를 Read로 읽은 후 작업해라

# 작업 마무리
- 작업이 끝나면 Claude가 직접 git add, commit, push까지 실행해줘. 사용자한테 명령어 알려주지 말고 직접 해줘.
- git push 경로: C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE
- push 시 네트워크 오류가 나면 `git pushretry`로 재시도해줘 (최대 3회 자동 재시도)
