# Investment — 부동산 투자 프로젝트 플랫폼

부동산 자산운용사 투자팀을 위한 Claude Code 기반 멀티에이전트 작업 공간입니다. 새 딜(프로젝트)을 하나 추가하면 표준 폴더가 생성되고 6개 전문 에이전트가 자동으로 킥오프합니다.

## 에이전트
| # | 에이전트 | 역할 |
|---|---|---|
| 1 | `data-steward` | 원본 input 데이터 관리, 버전·히스토리 분석 |
| 2 | `legal-counsel` | 계약서·약정서 검토, 조건 협의, 법률 의견서 |
| 3 | `research-analyst` | 보고서, 투심 PPT/Word, 발표 스크립트 |
| 4 | `pmo-scheduler` | 일정 조율, 회의록, 액션아이템 |
| 5 | `finance-ir` | 수익률 분석, 투자자 Q&A 대응 |
| 6 | `decision-advisor` | 의사결정 아젠다, 전략 판단 근거자료 |

## 빠른 시작
```bash
# 1) 새 프로젝트 (Claude Code 안에서)
/new-project GNG-TOWER "강남 G타워" 오피스 매입

# 또는 터미널에서
platform/scripts/new_project.sh GNG-TOWER "강남 G타워" 오피스 매입 홍길동

# 2) 원본 자료를 넣고
cp ~/Downloads/렌트롤.xlsx projects/GNG-TOWER/00_원본자료/raw/20260916_렌트롤_매도자제공_v01.xlsx

# 3) Claude에게
"GNG-TOWER 새 자료 접수 처리해줘"        → data-steward
"SPA 초안 검토해줘"                       → legal-counsel
"투자자가 IRR 근거 물어봤어, 답변 초안"   → finance-ir
"오늘 회의 내용 회의록으로"               → pmo-scheduler
/ic-package GNG-TOWER                     → 투심 패키지 일괄 생성
/project-status GNG-TOWER                 → 현황 점검
```

## 프로젝트 폴더 구조
```
projects/<코드명>/
├── PROJECT.md          프로젝트 메타·상태 요약·핵심 지표
├── 00_원본자료/         raw(접수 원본, 수정 금지) · processed · history · INDEX · CHANGELOG
├── 01_설립_인허가/      설립서류 · 인허가 · 정관_약정 · CHECKLIST
├── 02_실사보고서/       법률 · 시장 · 재무 · 물리 · 감정 · DD_ISSUE_LOG
├── 03_IM_투자심의/      IM · 투심자료 · 발표스크립트
├── 04_Equity투자자/     투자자별 · models · qa_log · QA_BANK · ASSUMPTIONS · INVESTORS
├── 05_대출업무/         텀싯 · 약정서 · 대주QA · LENDERS · COVENANTS
├── 06_사내보고/         아젠다 · 의사결정보고 · DECISION_LOG
└── 07_일정_회의록/      TIMELINE · ACTION_ITEMS · minutes
```

## 설정 파일
- `CLAUDE.md` 에이전트 라우팅과 자동 작동 규칙
- `.claude/settings.json` 세션 시작 시 프로젝트 현황 출력 훅
- `shared/strategy.md` 회사 투자 전략 기준 (먼저 채워 두면 decision-advisor가 정합성 판정에 사용)
