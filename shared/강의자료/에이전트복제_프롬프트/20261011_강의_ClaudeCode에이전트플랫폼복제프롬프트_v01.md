# Claude Code로 "부동산 투자팀 6-에이전트 플랫폼" 만들기 — 강의용 프롬프트 모음 v01

- 작성일: 2026-10-11
- 용도: 강의 시연과 실습. 수강생이 빈 폴더에서 Claude Code에 단계별로 붙여넣어 같은 구조를 다시 만든다.
- 이 문서에는 특정 회사·거래·자산 정보가 없다. 예시 딜은 가상의 `DEMO-LOGIS`(가상 물류센터)를 쓴다.

---

## 0. 완성되면 이런 구조가 된다

```
CLAUDE.md                    ← 팀 규칙 + 에이전트 라우팅표 (Claude가 매 세션 읽음)
.claude/agents/              ← 서브에이전트 6개 (역할·담당 폴더·산출 형식)
.claude/skills/              ← /new-project · /project-status · /ic-package
.claude/settings.json        ← SessionStart 훅 (세션 시작 시 프로젝트 현황 출력)
platform/templates/project/  ← 신규 딜 표준 폴더 (00~08)
platform/scripts/            ← new_project.sh · list_projects.sh · validate_project.sh · extract_text.py
platform/standards/          ← 명명 규칙 · 단위·지표 정의 · 보고 스타일 · 법률 체크리스트
projects/<코드명>/            ← 딜별 작업 공간 (REGISTRY.md에 자동 등록)
shared/                      ← 시장자료 · 과거 선례 · 투자 전략 기준
```

| # | 에이전트 | 하는 일 | 산출 폴더 |
|---|---|---|---|
| 1 | data-steward | 원본 접수·등록·텍스트 추출·버전 비교·이상치 (단일 진실 원천) | 00_원본자료 |
| 2 | legal-counsel | 계약서·약정서 검토, 협상 포인트, 법률 의견서 | 01, 02/법률, 05/약정서, 08 |
| 3 | research-analyst | IM 검증, 투심 PPT·Word, 1-pager, 발표 스크립트, 시장 실사 | 03, 02/시장 |
| 4 | pmo-scheduler | 마스터 일정, 회의록, 액션아이템 | 07 |
| 5 | finance-ir | 수익률·민감도, 투자자·대주 Q&A, 재무약정 | 04, 05, 02/재무·감정 |
| 6 | decision-advisor | 의사결정 아젠다, 옵션 비교, 결의안 | 06 |

**핵심 설계 4가지** (강의에서 강조할 점)
1. **역할 분리**: 에이전트마다 담당 폴더와 산출 형식이 정해져 있어 결과물이 섞이지 않는다.
2. **단일 진실 원천**: 수치는 data-steward가 정제한 `processed/` 값만 인용한다. 그래서 문서마다 숫자가 달라지지 않는다.
3. **전파 규칙**: 각 에이전트는 결과 끝에 "전파 필요 사항"을 적는다. 메인 Claude가 이를 다음 담당 에이전트에게 넘긴다.
4. **자동화**: 훅(세션 시작 현황)과 스킬(/new-project 등)로 반복 작업을 명령어 하나로 만든다.

---

## 1. 준비

```bash
mkdir re-invest-platform && cd re-invest-platform
git init
claude
```

- 첫 설계 단계는 **Plan 모드**(Shift+Tab 두 번)에서 시작하는 것을 권한다. 계획을 먼저 확인한 뒤 실행하게 하면 잘못 만든 것을 되돌릴 일이 줄어든다.
- 단계가 하나 끝날 때마다 `git commit`을 해 두면 실습 중 되돌리기 쉽다.

---

## 2. 단계별 프롬프트 (복사해서 순서대로 붙여넣기)

### STEP 1. 전체 설계 받기 (Plan 모드)

```text
부동산 자산운용사 투자팀(운용역)이 딜 검토부터 투자심의·계약·클로징까지
모든 업무를 처리하는 작업 플랫폼을 이 저장소에 만들고 싶어. 언어는 한국어로 해 줘.

요구사항:
- 6개 전문 서브에이전트가 업무를 나눠 맡는다:
  data-steward(원본 자료 관리), legal-counsel(법률), research-analyst(문서·리서치),
  pmo-scheduler(일정·회의록), finance-ir(재무·투자자 대응), decision-advisor(의사결정 지원)
- 딜마다 projects/<코드명>/ 아래에 같은 표준 폴더를 쓴다:
  00_원본자료, 01_설립_인허가, 02_실사보고서, 03_IM_투자심의, 04_Equity투자자,
  05_대출업무, 06_사내보고, 07_일정_회의록, 08_매매_임대차계약서
- CLAUDE.md에 "요청 성격 → 담당 에이전트 → 산출 폴더" 라우팅표를 둔다.
- 신규 딜 생성, 현황 점검, 투심 패키지 생성은 슬래시 커맨드(스킬)로 만든다.
- 세션을 시작하면 전체 프로젝트 현황이 자동으로 출력되게 한다.

지금은 파일을 만들지 말고, 전체 디렉터리 구조와 파일 목록,
에이전트별 책임·입력·산출 폴더, 에이전트 간 데이터 흐름(의존 순서)을 계획으로 보여줘.
```

> 강의 포인트: 계획이 나오면 수강생에게 "어느 에이전트가 어떤 폴더에 쓰는지"를 확인시킨다. 고칠 점이 있으면 이 단계에서 말로 고치게 한다.

---

### STEP 2. CLAUDE.md와 표준 문서

```text
계획대로 진행해. 먼저 다음을 만들어 줘.

1) CLAUDE.md
   - 플랫폼 소개(2줄), 디렉터리 구조
   - 에이전트 라우팅표 (요청 성격 | 에이전트 | 산출 폴더)
   - 다중 영역 요청의 실행 순서: 데이터 → (재무 ∥ 법률) → 리서치 → 의사결정 → 일정 반영.
     독립 작업은 한 메시지에서 병렬 실행
   - 자동 작동 규칙:
     ① 세션 시작 시 현황 출력, 미등록 원본이 있으면 접수 처리를 제안
     ② 사용자가 새 딜을 언급하면 /new-project 실행
     ③ 00_원본자료/raw/는 신규 접수함 — 새 파일이 보이면 항상 data-steward 먼저 실행
        (INDEX 등록 → 텍스트 추출 → 업무 폴더로 이동)
     ④ 산출물을 만들면 PROJECT.md "현재 상태 요약" 표 갱신
     ⑤ 에이전트 결과의 "전파 필요 사항"은 지정된 에이전트에게 후속 작업으로 넘김
     ⑥ 같은 자료가 여러 판본이면 작성일자가 가장 최신인 파일을 기준으로 함
   - 공통 원칙:
     원본 파일은 절대 수정하지 않음 / 모든 수치에 출처 파일·버전 표기 /
     추정은 [추정], 미접수는 [미접수] / 수익을 약속하는 표현("보장", "확정") 금지 /
     법률 검토는 내부 검토이며 외부 법무법인 확인이 필요하다고 항상 명시

2) platform/standards/
   - naming.md: 파일명 YYYYMMDD_<유형>_<제목>_v##.<ext>, 유형 목록
   - units_and_metrics.md: 단위(면적 ㎡+평, 금액 백만원/억원, % 소수 2자리)와
     지표 정의(NOI, Cap Rate, IRR, Equity Multiple, DSCR, LTV, WALE, 공실률 — 분모 명시 규칙 포함)
   - report_style.md: 결론 먼저, 한 장 한 메시지, 리스크는 리스크→완화→잔여 3열
   - legal_checklist.md: 매매계약·임대차·대출약정·펀드/신탁·MOU별 검토 항목

3) shared/README.md, shared/strategy.md(투자 전략 기준 빈 양식), shared/market/, shared/precedents/
```

---

### STEP 3. 프로젝트 템플릿과 스크립트

```text
다음을 만들어 줘.

1) platform/templates/project/ — 신규 딜 표준 폴더.
   각 폴더에 README.md(용도·담당 에이전트·파일 규칙)를 넣고, 아래 추적 파일의 빈 양식을 만든다:
   - PROJECT.md: 딜 개요 표, 핵심 지표 표, 현재 상태 요약 표(영역 | 담당 | 상태 | 최종 갱신), 기준 일정·기준 재무모델 지정란
   - 00_원본자료/: raw/, text/, processed/, history/, INDEX.md, CHANGELOG.md, DATA_DICTIONARY.md
   - 01_설립_인허가/CHECKLIST.md, 02_실사보고서/DD_ISSUE_LOG.md(심각도·상태)
   - 04_Equity투자자/ASSUMPTIONS.md · QA_BANK.md · INVESTORS.md
   - 05_대출업무/LENDERS.md · COVENANTS.md
   - 06_사내보고/DECISION_LOG.md, 07_일정_회의록/TIMELINE.md · ACTION_ITEMS.md

2) platform/scripts/
   - new_project.sh <코드명> "<자산명>" [자산유형] [투자유형] [담당자]
     : 템플릿 복사, PROJECT.md 자리표시자 치환, projects/REGISTRY.md에 한 줄 등록.
       같은 코드명이 있으면 중단
   - list_projects.sh: REGISTRY의 모든 프로젝트에 대해 단계·상태·원본 수·INDEX 등록 수·
     미완료 액션 수를 한 줄씩 출력하고, raw/에 INDEX 미등록 파일이 있으면 ⚠ 경고
   - validate_project.sh <코드명>: 필수 폴더·파일 존재 여부를 점검해 누락 목록 출력
   - extract_text.py <코드명>: 프로젝트 안의 pdf·docx·xlsx·pptx를 텍스트로 추출해
     00_원본자료/text/<원본파일명>.txt로 저장 (맨 윗줄에 원본 경로 표기).
     텍스트가 거의 없는 PDF는 [스캔-OCR필요]로 표시. 이미 추출된 파일은 건너뜀

3) 스크립트를 실제로 실행해 TEST 프로젝트로 동작을 확인하고, 확인이 끝나면 TEST는 지워 줘.
```

---

### STEP 4. 서브에이전트 6개

```text
.claude/agents/에 서브에이전트 6개를 만들어 줘. 모든 파일은 같은 골격을 따른다:

---
name: <영문 이름>
description: "<#번호 역할 한 줄>. <언제 쓰는지>. 예: '<사용자 요청 예시 3개>'"
tools: <필요한 도구만>
model: inherit
---
본문 섹션: 역할 정의(1문단) / 담당 폴더(입력·산출) / 핵심 원칙 / 작업 순서 /
산출물 형식(마크다운 골격) / 협업(누구에게 무엇을 넘기는지) / 하지 말아야 할 것.
결과 보고 마지막에는 항상 "전파 필요 사항(→ 담당 에이전트)"을 적게 한다.

에이전트별 요구사항:
1. data-steward (Read, Write, Edit, Bash, Glob, Grep)
   - raw/ 신규 파일 → INDEX 등록 → extract_text.py 실행 → processed/에 핵심 수치 표 →
     이전 판본이 있으면 history/에 비교표(항목, 이전값, 현재값, 변동률) → CHANGELOG 기록 →
     원본을 해당 업무 폴더(01~08)로 이동. 원본은 절대 수정하지 않음
   - 전 판본 대비 ±10% 이상 변동, 합계 불일치, 단위 혼용은 "이상치" 섹션으로 별도 표시
2. legal-counsel (… + WebSearch, WebFetch)
   - 계약 유형별 체크리스트(standards/legal_checklist.md) 기준으로 검토
   - 의견서 형식: 검토 요약(결론 먼저) → 조항별 표(조항 | 내용 | 리스크 상·중·하 | 수정안) →
     협상 포인트 → 외부 법무법인 확인 필요 사항
   - 발견사항은 DD_ISSUE_LOG.md에, 재무약정은 COVENANTS.md에 등록
3. research-analyst (… + Skill, WebSearch, WebFetch)
   - IM 검증 메모(IM 주장 vs 실사 결과), 투심 PPT·Word, 1-pager, 발표 스크립트, 시장 실사
   - 결론형 슬라이드 제목, 모든 수치에 출처 각주, 수치는 processed/ 값만 사용
     (계산이 필요하면 finance-ir에 요청)
   - PPT를 만들기 전에 outline.md(슬라이드별 제목·메시지·데이터)를 먼저 작성.
     표는 실제 표 객체로만 만든다
4. pmo-scheduler (Read, Write, Edit, Glob, Grep, Bash)
   - TIMELINE.md: 단계별 마일스톤·기한·담당·상태·의존관계, 표준 리드타임으로 역산
   - 회의 메모를 주면 회의록(결정사항·액션아이템·미결사항) 작성 → ACTION_ITEMS.md 갱신
   - 일정이 밀릴 때의 영향(크리티컬 패스)을 분석
5. finance-ir (Read, Write, Edit, Bash, Glob, Grep)
   - ASSUMPTIONS.md, models/ 수익률·민감도, QA_BANK.md 질문-답변 뱅크, LENDERS.md 대주 조건 비교
   - 답변 구조: 결론 1문장 → 근거 수치 표 → 가정 → 리스크·단서. 답변 파일은 [대외 발송용]과 [내부 참고]로 구분
   - 데이터 판본이 바뀌면 기존 Q&A를 [갱신 필요]로 표시
6. decision-advisor (Read, Write, Edit, Bash, Glob, Grep, WebSearch)
   - 아젠다(결정 필요 사항 표: 안건 | 결정 유형 | 요청 결정 | 기한 | 미결정 시 영향)
   - 의사결정 보고서 1~2쪽: 결론·권고 → 옵션 비교표 → 근거 → 리스크 → 요청 사항
   - DECISION_LOG.md에 결정 이력(일자·안건·결정·조건·후속조치) 기록
   - shared/strategy.md 기준으로 전략 부합 여부를 판정

만든 뒤 CLAUDE.md 라우팅표와 에이전트별 산출 폴더가 서로 맞는지 대조해 줘.
```

> 강의 포인트: `description`이 라우팅을 결정한다. "언제 쓰는지"와 "요청 예시"가 구체적일수록 메인 Claude가 알맞은 에이전트를 부른다.

---

### STEP 5. 슬래시 커맨드(스킬) 3종

```text
.claude/skills/에 스킬 3개를 만들어 줘 (각 폴더에 SKILL.md, frontmatter에 name·description).

1) new-project — "새 딜", "새 프로젝트", "/new-project <코드명> "<자산명>"" 에 반응
   ① new_project.sh와 validate_project.sh 실행
   ② 사용자가 준 딜 정보(위치·규모·가격·투자 포인트)를 PROJECT.md 딜 개요에 기록
   ③ 6개 에이전트 킥오프를 한 메시지에서 병렬 실행
      (data-steward: 접수 요청 목록 / pmo: 역산 일정 초안 / finance-ir: 초기 가정·예상 질문 10개 /
       legal: 설립·인허가 체크리스트 / research: IM 검토 틀 / decision: 첫 아젠다)
   ④ 결과를 PROJECT.md 상태 표에 반영하고 요약 보고

2) project-status — "현황", "/project-status [코드명]"
   - 코드명이 없으면 list_projects.sh 결과만 출력
   - 코드명이 있으면 에이전트를 부르지 않고 파일만 읽어 점검:
     미등록 원본, 최근 CHANGELOG, DD 이슈 중 심각도 상, 다음 마일스톤 D-day,
     기한이 지난 액션, [갱신 필요] Q&A, 조건부 결정의 이행 여부 → 상태 표 갱신, 후속 조치 제안

3) ic-package — "투심 자료 만들어줘", "/ic-package <코드명>"
   순차 파이프라인: data-steward 검증 → (finance-ir 수익률·민감도·예상 질문 ∥ legal-counsel 투심용 법률 요약)
   → research-analyst 개요·PPT·발표 스크립트 → decision-advisor 결의안 초안 → pmo 일정 반영.
   원본 자료(렌트롤·감정평가·IM)가 하나도 없으면 시작하지 말고 필요한 자료를 안내
```

---

### STEP 6. 세션 시작 훅

```text
.claude/settings.json에 SessionStart 훅을 추가해서, 세션을 시작할 때
platform/scripts/list_projects.sh 결과가 자동으로 출력되게 해 줘.
명령은 "$CLAUDE_PROJECT_DIR" 기준 경로로 쓰고, 추가한 뒤 스크립트가 오류 없이 도는지 확인해 줘.
```

> 확인 방법: Claude Code를 종료했다가 다시 실행해서 첫 화면에 현황 요약이 나오는지 본다.

---

### STEP 7. 시연 — 가상 딜로 전체 흐름 돌려보기

```text
새 딜 시작할게. /new-project DEMO-LOGIS "가상 물류센터" 물류 실물매입 홍길동
경기도 소재 연면적 약 3만㎡ 상온 물류센터, 매입가 [●]억원 검토, 주요 임차인 1개사 장기임대.
```

이어서 가상의 원본 자료를 넣고 접수 흐름을 확인한다. (실습용 가상 렌트롤을 만드는 프롬프트)

```text
projects/DEMO-LOGIS/00_원본자료/raw/에 실습용 가상 렌트롤 엑셀을 하나 만들어 줘.
층·임차인(가상 이름)·면적·보증금·월 임대료·관리비·만기 10행,
그중 1행은 단위를 일부러 평으로 섞어서 이상치 검출을 시연할 수 있게 해 줘.
그다음 접수 규칙대로 처리해.
```

교차 업무 요청 (라우팅과 병렬 실행을 시연):

```text
DEMO-LOGIS: 주요 임차인 임대료가 5% 내려간다고 가정하면
① 수익률이 얼마나 바뀌는지, ② 대출 재무약정(DSCR)에 걸리는지,
③ 임대차계약상 우리가 쓸 수 있는 대응 조항이 있는지 보고,
대표 보고용 1쪽 의사결정 보고서로 정리해 줘.
```

> 강의 포인트: 메인 Claude가 finance-ir와 legal-counsel을 병렬로 부르고, 그 결과를 decision-advisor에게 넘기는지 보여준다. 끝나면 PROJECT.md 상태 표가 갱신됐는지도 확인한다.

---

### STEP 8. 점검·개선 프롬프트

```text
지금까지 만든 플랫폼을 점검해 줘.
1) CLAUDE.md 라우팅표 ↔ 각 에이전트 담당 폴더 ↔ 템플릿 폴더 구조가 서로 맞는지
2) 에이전트 description끼리 겹쳐서 라우팅이 헷갈릴 만한 곳
3) 스크립트를 공백·한글 파일명에서 실행해도 문제가 없는지
4) 원본 불변, 출처 표기, 확약 표현 금지 규칙이 모든 에이전트에 들어 있는지
문제를 표로 정리하고, 고칠 내용을 제안한 뒤 내가 승인하면 반영해.
```

---

## 3. 시간이 부족할 때 — 한 번에 만드는 압축 프롬프트

```text
이 저장소에 "부동산 자산운용사 투자팀용 Claude Code 작업 플랫폼"을 만들어 줘. 한국어로.

[구성]
- CLAUDE.md: 라우팅표(요청 성격→에이전트→산출 폴더), 실행 순서(데이터→재무∥법률→리서치→의사결정→일정),
  자동 규칙(raw/ 신규 파일은 data-steward 먼저, 산출물 생성 시 PROJECT.md 상태 갱신, "전파 필요 사항" 후속 위임,
  최신 작성일 판본 기준), 공통 원칙(원본 불변, 수치 출처 표기, [추정]/[미접수] 태그, "보장·확정" 금지,
  법률 검토는 내부 검토·외부 법무법인 확인 필요 명시, 파일명 YYYYMMDD_<유형>_<제목>_v##)
- .claude/agents/ 6개: data-steward, legal-counsel, research-analyst, pmo-scheduler, finance-ir, decision-advisor.
  각각 frontmatter(name, description에 사용 예시 3개, 최소 tools, model: inherit)와
  본문(역할·담당 폴더·원칙·작업 순서·산출 형식·협업·금지사항·결과 끝 "전파 필요 사항")
- 프로젝트 템플릿 platform/templates/project/: 00_원본자료(raw·text·processed·history, INDEX·CHANGELOG),
  01_설립_인허가 ~ 08_매매_임대차계약서, PROJECT.md 및 각 추적 파일(TIMELINE, ACTION_ITEMS, QA_BANK,
  ASSUMPTIONS, LENDERS, COVENANTS, DD_ISSUE_LOG, DECISION_LOG)
- platform/scripts/: new_project.sh, list_projects.sh(미등록 원본 경고), validate_project.sh, extract_text.py
- platform/standards/: naming, units_and_metrics(지표 정의·분모 명시), report_style, legal_checklist
- .claude/skills/: new-project(스캐폴딩+6개 에이전트 병렬 킥오프), project-status(파일만 읽어 현황 점검),
  ic-package(투심 패키지 순차 파이프라인)
- .claude/settings.json: SessionStart 훅으로 list_projects.sh 실행

[진행 방식]
먼저 계획을 보여주고 내 승인을 받은 뒤 만들어. 다 만들면 TEST 프로젝트로 스크립트·스킬을 실제로 돌려 보고,
문제를 고친 다음 TEST는 지우고, 무엇을 만들었는지 표로 보고해.
```

---

## 4. 실습 중 자주 생기는 문제와 보정 프롬프트

| 증상 | 원인 | 보정 프롬프트 |
|---|---|---|
| 메인 Claude가 에이전트를 안 부르고 직접 처리 | description이 모호함 | `각 에이전트 description에 "언제 쓰는지"와 사용자 요청 예시 3개를 더 구체적으로 넣어 줘.` |
| 문서마다 숫자가 다름 | 원본에서 직접 숫자를 읽음 | `모든 에이전트가 수치는 00_원본자료/processed/ 값만 인용하고 파일명·버전을 붙이도록 규칙을 강화해 줘.` |
| 원본 파일이 수정됨 | 원본 불변 규칙 누락 | `원본 불변 원칙을 CLAUDE.md와 6개 에이전트의 "하지 말아야 할 것"에 모두 넣고, 정제본은 새 파일로만 만들게 해 줘.` |
| 후속 작업이 끊김 | 전파 규칙 미작동 | `에이전트 결과의 "전파 필요 사항"을 읽으면 지정된 에이전트를 바로 이어서 실행하도록 CLAUDE.md 규칙을 고쳐 줘.` |
| 훅이 동작하지 않음 | 경로·실행 권한 | `SessionStart 훅 명령을 "$CLAUDE_PROJECT_DIR" 기준 경로와 bash 실행으로 고치고, 스크립트를 직접 실행해 확인해 줘.` |
| PPT의 표가 도형으로 그려짐 | 생성 규칙 누락 | `PPT의 표는 반드시 실제 표 객체로 만들고, 만든 뒤 PDF로 렌더링해서 겹침·잘림을 확인하라는 규칙을 research-analyst에 추가해 줘.` |

---

## 5. 강의 시 유의사항

- 실습에는 **가상 데이터만** 쓴다. 실제 거래 자료, 회사 내부 양식, 계약서 원문은 넣지 않는다(비밀유지·저작권).
- 회사 고유 양식(투심 PPT 템플릿 등)은 각 회사가 자기 템플릿을 `platform/templates/`에 넣고, 이를 쓰라는 규칙을 CLAUDE.md에 추가하는 방식으로 연결한다.
- 에이전트가 만든 법률·재무 결과물은 초안이다. 최종 판단과 외부 전문가 확인은 사람이 한다는 점을 강조한다.
