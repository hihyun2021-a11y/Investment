# 00_원본자료 (담당: data-steward)

- `raw/` 접수한 파일을 **그대로** 넣습니다. 내용 수정 금지. 파일명은 `YYYYMMDD_<유형>_<제목>_v##.<ext>` 권장.
- `processed/` 원본에서 추출·표준화한 데이터 (CSV/MD). 팀의 단일 수치 원천(SSOT).
- `history/` 버전 간 비교, 시계열 추이 분석.
- `INDEX.md` 접수 원본 목록 / `CHANGELOG.md` 변경 이력 / `DATA_DICTIONARY.md` 항목 정의·단위.

새 파일을 넣은 뒤 "새 자료 접수 처리해줘"라고 요청하면 data-steward가 등록·정제·비교를 수행합니다.
같은 자료의 새 버전은 삭제하지 말고 새 파일로 추가합니다.
