---
name: new-project
description: 신규 부동산 투자 프로젝트를 생성하고 6개 전문 에이전트의 킥오프를 자동 실행한다. 사용자가 "새 프로젝트", "새 딜 시작", "/new-project"라고 하면 사용. 인자: <코드명> "<자산명>" [자산유형] [투자유형] [담당자]
---

# 신규 프로젝트 생성 및 자동 킥오프

## 1. 스캐폴딩
인자를 확인하고 실행한다. 코드명이나 자산명이 없으면 사용자에게 한 번만 묻는다.
```bash
platform/scripts/new_project.sh <코드명> "<자산명>" [자산유형] [투자유형] [담당자]
platform/scripts/validate_project.sh <코드명>
```

## 2. PROJECT.md 딜 개요 채우기
사용자가 프롬프트에서 준 정보(위치, 규모, 매도자, 예상 가격, 투자 포인트)가 있으면 `projects/<코드명>/PROJECT.md`의 "딜 개요"에 기록한다. 없으면 비워 두고 사용자에게 알린다.

## 3. 에이전트 킥오프 (Agent 도구로 병렬 실행)
아래 6개를 **한 메시지에서 동시에** 실행한다. 각 에이전트에게 프로젝트 코드명, PROJECT.md 경로, 사용자가 준 딜 정보를 전달한다.

| 에이전트 | 킥오프 과업 |
|---|---|
| data-steward | `00_원본자료/raw/` 스캔. 파일이 있으면 INDEX 등록·정제·DATA_DICTIONARY 작성. 없으면 접수해야 할 자료 목록을 `00_원본자료/INDEX.md` 하단에 "접수 요청 목록"으로 작성 |
| pmo-scheduler | 딜 정보 기반으로 `TIMELINE.md` 역산 일정 초안 수립(목표일 미정이면 표준 리드타임 적용), ACTION_ITEMS 갱신, README "이번 주 핵심 일정" 작성 |
| finance-ir | `04_Equity투자자/ASSUMPTIONS.md`에 초기 가정(재무모델·IM 접수 시 실제 값), `QA_BANK.md`에 투자자 유형별 예상 질문 10개 시드, `05_대출업무/LENDERS.md`·`COVENANTS.md`에 IM상 대출 조건 기입 |
| legal-counsel | 비히클 구조에 맞춰 `01_설립_인허가/CHECKLIST.md` 항목·근거 법령 구체화, 검토 대상 계약 목록을 `02_실사보고서/법률/README.md`로 작성. 법률실사보고서가 접수돼 있으면 요약과 `DD_ISSUE_LOG.md` 등록 |
| research-analyst | `03_IM_투자심의/IM/`에 1-pager 초안(`YYYYMMDD_1pager_<코드명>_v01.md`) 작성. IM이 접수돼 있으면 IM 요약·검증 메모 포함. 데이터 없는 항목은 `[데이터 대기]` |
| decision-advisor | `06_사내보고/아젠다/`에 첫 아젠다: "검토 착수 여부" Go/No-go 조건 초안 |

## 4. 마무리
- 각 에이전트 결과를 PROJECT.md "현재 상태 요약" 표에 반영한다.
- `projects/REGISTRY.md` 행이 생성됐는지 확인한다.
- 사용자에게: 생성된 경로, 킥오프 결과 요약(에이전트별 1줄), 다음에 할 일(원본 자료 적재)을 보고한다.
