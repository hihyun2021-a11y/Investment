---
name: project-status
description: 특정 프로젝트(또는 전체)의 6개 영역 현황을 점검하고 PROJECT.md 상태 표를 갱신한다. "현황 알려줘", "상태 점검", "/project-status <코드명>"에 사용. 코드명 생략 시 전체 요약.
---

# 프로젝트 현황 점검

## 코드명이 없을 때
`platform/scripts/list_projects.sh`를 실행해 전체 요약을 보여주고 끝낸다.

## 코드명이 있을 때
1. `projects/<코드명>/PROJECT.md`를 읽는다.
2. 아래를 직접 확인한다 (에이전트 호출 없이 파일만 읽어 빠르게):
   - `00_원본자료/raw/` 파일 수 vs `INDEX.md` 등록 수 → 미등록 자료
   - `00_원본자료/CHANGELOG.md` 최근 변경
   - `02_실사보고서/DD_ISSUE_LOG.md` 심각도 高 미해결 이슈
   - `07_일정_회의록/TIMELINE.md` 다음 마일스톤과 D-day, `ACTION_ITEMS.md` 기한 경과 항목
   - `04_Equity투자자/QA_BANK.md` `[갱신 필요]` 항목
   - `06_사내보고/DECISION_LOG.md` 조건부 결정의 조건 이행 여부
3. PROJECT.md "현재 상태 요약" 표를 갱신한다 (최종 갱신일, 상태 한 줄).
4. 필요한 후속 조치가 있으면 담당 에이전트를 지정해 제안한다. 미등록 원본 자료가 있으면 data-steward를 바로 실행한다.

## 출력 형식
```
## <코드명> <자산명> — <단계> (D-xx to <다음 마일스톤>)
| 영역 | 상태 | 조치 필요 |
...
### 즉시 조치 (담당 에이전트)
```
