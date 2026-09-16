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
| data-steward | `00_input/raw/` 스캔. 파일이 있으면 INDEX 등록·정제·DATA_DICTIONARY 작성. 없으면 접수해야 할 자료 목록을 `00_input/INDEX.md` 하단에 "접수 요청 목록"으로 작성 |
| pmo-scheduler | 딜 정보 기반으로 `TIMELINE.md` 역산 일정 초안 수립(목표일 미정이면 표준 리드타임 적용), ACTION_ITEMS 갱신, README "이번 주 핵심 일정" 작성 |
| finance-ir | `ASSUMPTIONS.md`에 초기 가정 항목과 필요 데이터 표시, `QA_BANK.md`에 투자자 유형별 예상 질문 10개 시드(답변은 데이터 접수 후) |
| legal-counsel | 투자 유형에 따라 검토 대상 계약 목록과 체크리스트를 `02_legal/ISSUE_LOG.md` 상단에 "검토 대상 문서" 섹션으로 작성 |
| research-analyst | `03_research/reports/`에 1-pager 골격(`YYYYMMDD_보고서_1pager_v00.md`) 작성. 데이터 없는 항목은 `[데이터 대기]` |
| decision-advisor | `06_decision/agendas/`에 첫 아젠다: "검토 착수 여부" Go/No-go 조건 초안 |

## 4. 마무리
- 각 에이전트 결과를 PROJECT.md "현재 상태 요약" 표에 반영한다.
- `projects/REGISTRY.md` 행이 생성됐는지 확인한다.
- 사용자에게: 생성된 경로, 킥오프 결과 요약(에이전트별 1줄), 다음에 할 일(원본 자료 적재)을 보고한다.
