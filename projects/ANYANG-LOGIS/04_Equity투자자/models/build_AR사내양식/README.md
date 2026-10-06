# 재무모델 A&R 사내 표준양식 탭 생성 (ANYANG-LOGIS)

기준 재무모델(.xlsm)에 `A&R(사내양식)` 탭을 추가한다. 원본 모델은 수정하지 않고 새 버전 파일을 만든다.

1. `stage_a.py` — 사내 표준양식(`platform/templates/financial_model/재무모델_AssumptionResult_사내표준양식_261006.xlsx`)의 '실물' 시트를 복제해
   모델 시트(A&R·IM·CF(Y)·IS(FY))를 직접 참조하는 수식으로 채운다 (천원 → 백만원).
2. `stage_b.py` — 1단계 시트를 .xlsm 패키지에 직접 이식한다. 매크로·양식컨트롤·웹확장·메모 등 기존 파트는 바이트 그대로 유지하고
   styles.xml(서식 추가), workbook.xml·rels·[Content_Types]·app.xml(시트 1개 추가)만 변경한다.
3. LibreOffice 재계산(`lo_prof`: OOXMLRecalcMode=0)으로 수식 결과를 구해 `collect.py`가 values.json으로 저장 → `stage_b.py`가 캐시값(<v>)으로 기록.
4. `preview.py` — 시각 점검용 값 치환본 생성.

`run_all.sh`는 작업 폴더 기준 상대경로(`../ar_in/form.xlsx`, `../ar_in/model_v03.xlsm`, `lo_prof/`)를 사용하므로, 실행 전 경로를 맞춘다.
IRR은 연간 CF 기준 XIRR(모델 CF(Y) I열과 동일 기준)이며, 반기 기준 IM 탭 값(Equity 전체 10.21%)과 집계주기 차이가 있다.
