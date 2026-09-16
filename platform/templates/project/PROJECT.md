# {{NAME}} ({{CODE}})

| 항목 | 내용 |
|---|---|
| 프로젝트 코드 | {{CODE}} |
| 자산명 | {{NAME}} |
| 자산 유형 | {{ASSET_TYPE}} |
| 투자 유형 | {{DEAL_TYPE}} |
| 생성일 | {{DATE}} |
| 담당 운용역 | {{OWNER}} |
| 현재 단계 | 1. 소싱/검토 |
| 상태 | 진행중 |

## 딜 개요
<!-- 위치, 규모, 매도자, 예상 매입가, 핵심 투자 포인트 3줄 -->

## 현재 상태 요약 (에이전트가 갱신)
| 업무 폴더 | 담당 에이전트 | 최종 갱신 | 상태 한 줄 |
|---|---|---|---|
| 00 원본자료 | data-steward | - | 원본 자료 미접수 |
| 01 설립·인허가 | legal-counsel | - | 미착수 |
| 02 실사보고서 | data-steward / legal-counsel / finance-ir | - | 실사 미착수 |
| 03 IM·투자심의 | research-analyst | - | 작성 문서 없음 |
| 04 Equity 투자자 | finance-ir | - | 가정 미설정 |
| 05 대출업무 | finance-ir / legal-counsel | - | 미착수 |
| 06 사내보고 | decision-advisor | - | 결정 안건 없음 |
| 07 일정·회의록 | pmo-scheduler | - | 타임라인 미수립 |

## 핵심 지표 (finance-ir 갱신, 출처 필수)
| 지표 | 값 | 출처 |
|---|---|---|
| 매입가 | - | - |
| 감정가 | - | - |
| Cap Rate | - | - |
| LTV | - | - |
| WALE / 임대율 | - | - |
| Equity 수익률 (트랜치별) | - | - |

## 다음 마일스톤
<!-- pmo-scheduler가 07_일정_회의록/TIMELINE.md에서 복사 -->

## 폴더 안내
| 폴더 | 내용 | 담당 |
|---|---|---|
| `00_원본자료/` | 접수 원본(raw, 수정 금지), 정제 데이터(processed), 히스토리 | data-steward |
| `01_설립_인허가/` | 리츠·SPC 설립, 정관, 영업인가·등록, 인허가 | legal-counsel |
| `02_실사보고서/` | 법률·시장·재무·물리·감정 실사 요약 및 이슈 | data-steward, legal-counsel, finance-ir |
| `03_IM_투자심의/` | IM, 투자심의 자료(PPT/Word), 발표 스크립트 | research-analyst |
| `04_Equity투자자/` | 투자자별 대응, 재무 가정·모델, Q&A 뱅크 | finance-ir |
| `05_대출업무/` | 텀싯, 대출약정, 재무약정 관리, 대주 Q&A | finance-ir, legal-counsel |
| `06_사내보고/` | 의사결정 아젠다, 사내 보고서, 결정 이력 | decision-advisor |
| `07_일정_회의록/` | 마스터 일정, 액션아이템, 회의록 | pmo-scheduler |
