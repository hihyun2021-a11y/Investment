# -*- coding: utf-8 -*-
"""신규투자 업무 단계(모듈)와 산출물 종류(subtype) 정의.
samples: 저장소 루트 기준 glob 패턴 — 안양물류센터(ANYANG-LOGIS)·표준서식에서 기본 샘플로 자동 연결된다(없는 파일은 건너뜀).
fields: 화면 입력칸. 'text' 한 줄, 'area' 여러 줄.
view: 모듈 화면 위쪽에 붙는 전용 패널(sourcing·schedule·ic·procurement·distribution). 생성 패널은 모든 모듈에 공통."""

P = "projects/ANYANG-LOGIS/"
STD = "shared/강의자료/표준서식_Clean/"

PROJECT_FIELDS = [
    {"key": "project", "label": "프로젝트(코드명·자산명)", "type": "text", "placeholder": "예: ANYANG-LOGIS / 안양물류센터"},
]

MODULES = [
    {
        "id": "sourcing", "name": "Sourcing", "desc": "딜 발굴·자문사 마케팅 현황 파악, 지도 표시", "view": "sourcing",
        "subtypes": [
            {"id": "screening", "name": "딜 스크리닝 메모(1-pager)",
             "instructions": "투자 전략 기준(strategy.md)에 비추어 딜의 매력도·리스크·추가 확인사항·Go/No-go 의견을 1~2쪽 메모로 작성한다. 입지·임대·가격·구조·일정 항목을 표로 정리한다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "asset", "label": "자산 개요(위치·연면적·용도·준공)", "type": "area"},
                 {"key": "deal", "label": "매도인·매각자문사·입찰일정·희망가", "type": "area"},
                 {"key": "memo", "label": "기타 메모", "type": "area"}],
             "samples": ["shared/strategy.md", "shared/사내공모전/*딜Area발굴챌린지*.pdf", P + "03_IM_투자심의/IM/20260914_IM_*.pdf"]},
            {"id": "marketing", "name": "자문사 마케팅 현황 정리",
             "instructions": "자문사·브로커별로 받은 딜 정보(티저·메일·통화메모)를 딜 단위 표(자산명/유형/위치/규모/가격/자문사/마감일/진행상태/당사 대응)로 정리하고, 이번 주 대응 우선순위를 제시한다.",
             "fields": [{"key": "raw", "label": "자문사 메일·티저 내용 붙여넣기", "type": "area"}],
             "samples": ["shared/strategy.md"]},
        ]},
    {
        "id": "market", "name": "시장조사", "desc": "기초 인덱스별 리서치 및 가정 적용", "web_search": True,
        "subtypes": [
            {"id": "research", "name": "시장조사 보고서(인덱스별)",
             "instructions": "섹터·권역의 기초 인덱스(Cap rate, 공실률, 임대료, 신규공급, 거래사례, 금리)를 인덱스별 표로 정리하고, 각 수치에 출처(발행처·자료명·기준시점)를 단다. 웹 검색 결과를 쓸 때는 URL을 남긴다. 마지막에 당 딜에 대한 시사점을 쓴다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "sector", "label": "섹터·권역", "type": "text", "placeholder": "예: 수도권 남부 상온 물류"},
                 {"key": "index", "label": "필요한 인덱스", "type": "text", "placeholder": "Cap rate, 공실률, 임대료, 공급, 거래사례"},
                 {"key": "asof", "label": "기준시점", "type": "text", "placeholder": "예: 2026년 3Q"}],
             "samples": ["shared/market/*.pdf", "shared/사내공모전/*물류센터섹터현황점검*.pdf"]},
            {"id": "assumption", "name": "재무모델 가정(Assumption) 설정",
             "instructions": "임대료 상승률, 공실·Free rent, 운영비, Cap-in/Exit cap, 금리, 보유기간 등 재무모델 가정을 Base/Upside/Downside 표로 제시하고 항목별 근거(시장자료·비교사례)를 단다.",
             "fields": PROJECT_FIELDS + [{"key": "deal", "label": "딜 조건(매입가·NOI·임대현황)", "type": "area"}],
             "samples": ["shared/market/*.pdf", P + "00_원본자료/processed/20260916_렌트롤_IM기준_v01.csv"]},
        ]},
    {
        "id": "structuring", "name": "구조화", "desc": "Asset·Share·재투자·현물출자·유동화·OpCo-PropCo 구조 및 근거법규 검토",
        "subtypes": [
            {"id": "structure", "name": "투자구조 비교·근거법규 검토",
             "instructions": "검토 대상 구조(Asset deal, Share deal, 매도인 재투자, 현물출자, 유동화(PFV·ABS 등), OpCo-PropCo)를 비교표(세금·인허가·소요기간·리스크·장단점)로 정리하고, 구조별 근거법규(부동산투자회사법, 자본시장법, 법인세법, 지방세법·지방세특례제한법, 공정거래법 등) 조문을 정리한다. 법률 판단은 내부 검토이며 외부 법무법인 확인이 필요하다고 명시한다. 거래 구조도는 텍스트 도식으로 그린다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "options", "label": "검토할 구조", "type": "text", "placeholder": "Asset deal, Share deal, 재투자, 현물출자, 유동화, OpCo-PropCo"},
                 {"key": "facts", "label": "거래 사실관계(매도인·투자자·자금구조)", "type": "area"}],
             "samples": [P + "01_설립_인허가/설립서류/*태평양_v02_날인본.pdf", P + "01_설립_인허가/설립서류/*BKL_draft*.pdf",
                         P + "01_설립_인허가/정관_약정/20261008_검토메모_매도인PFV_3종재투자_적법성검토_v01.md",
                         "shared/사내공모전/*상속세*.html"]},
        ]},
    {
        "id": "docs", "name": "자료작성", "desc": "TM·IM·MOU·공문·결재·영업인가 서류",
        "subtypes": [
            {"id": "tm", "name": "TM(티저)", "instructions": "샘플 TM의 장표 구성·순서를 따라 장표별 제목(헤드메시지)과 본문 문구·표를 작성한다. 장표마다 '## [n] 장표명'으로 구분한다.",
             "fields": PROJECT_FIELDS + [{"key": "facts", "label": "자산·딜 정보", "type": "area"}],
             "samples": ["shared/강의자료/TM실습_데이터센터/*.pptx"]},
            {"id": "im", "name": "IM(투자설명서)", "instructions": "샘플 IM의 목차·장표 구성을 따라 장표별 헤드메시지·본문·표를 작성한다. 수치는 입력값만 쓰고 없으면 [데이터 대기]로 둔다.",
             "fields": PROJECT_FIELDS + [{"key": "facts", "label": "자산·임대·재무·구조 정보", "type": "area"}],
             "samples": ["shared/강의자료/IM실습_Clean/*.pptx", P + "03_IM_투자심의/IM/20261015_IM_*_v03.pptx"]},
            {"id": "mou", "name": "MOU(양해각서)", "instructions": "샘플 MOU의 조항 체계와 문체를 그대로 따르고 당사자·대상·가격·독점기간·실사·비밀유지 조항을 입력값으로 채운다. 미정 사항은 [  ]로 둔다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "당사자·주요 조건", "type": "area"}],
             "samples": [STD + "*06양해각서MOU*.docx"]},
            {"id": "letter", "name": "공문", "instructions": "샘플 공문의 머리글(회사명·주소·담당)·문서번호·수신·참조·제목·본문 번호체계(1. 2. - 아 래 - 1) 2))·별첨·서명 형식을 그대로 따른다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "sender", "label": "발신 명의", "type": "text", "placeholder": "예: 주식회사 코람코라이프로지스위탁관리부동산투자회사"},
                 {"key": "to", "label": "수신·참조", "type": "text"},
                 {"key": "body", "label": "요청·안내 내용", "type": "area"}],
             "samples": [P + "05_대출업무/약정서/20261028_대출약정_대출금입금요청공문_KLL명의_v02.docx",
                         P + "05_대출업무/약정서/20261028_참고_KLI공문서식*.docx"]},
            {"id": "approval", "name": "결재문(기안)", "instructions": "샘플 기안문 형식(제목, 승인 요청 문장, 1. 대상 2. 사유 3. 개요 4. 승인을 구하는 사항 5. 발송방법, 첨부)을 따른다.",
             "fields": PROJECT_FIELDS + [{"key": "subject", "label": "결재 대상·사유", "type": "area"}],
             "samples": [P + "05_대출업무/약정서/20261028_기안_대출금입금요청공문_날인발송_v01.docx"]},
            {"id": "license", "name": "영업인가 서류", "instructions": "부동산투자회사법상 위탁관리리츠 영업인가 신청에 필요한 서류 목록과 각 서류의 작성 초안(사업계획서 요지, 자산운용계획, 자금조달계획, 이해관계자 거래 등)을 작성한다. 근거 조문을 표시한다.",
             "fields": PROJECT_FIELDS + [{"key": "facts", "label": "리츠 개요·자산·자금조달", "type": "area"}],
             "samples": [P + "01_설립_인허가/CHECKLIST.md", P + "01_설립_인허가/인허가/*"]},
        ]},
    {
        "id": "contracts", "name": "계약서작성", "desc": "매매·임대차·주주간계약·위탁 3종·정관",
        "subtypes": [
            {"id": "spa", "name": "부동산 매매계약서", "instructions": "샘플 계약서의 조항 번호·체계·정의 조항·문체를 그대로 유지하고, 입력된 거래조건을 반영한 초안을 쓴다. 미정 사항은 [  ], 협상 포인트는 각주 형태 '※ 검토:'로 표시한다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "당사자·목적물·가격·일정·특약", "type": "area"}],
             "samples": [STD + "*01부동산매매계약서*.docx"]},
            {"id": "lease", "name": "임대차계약서", "instructions": "샘플 임대차계약서 체계를 유지하고 임대료·관리비·보증금·기간·중도해지·원상복구·임대인 변경(승계) 조항을 입력값으로 채운다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "임차인·면적·임대조건", "type": "area"}],
             "samples": [STD + "*05임대차계약서*.docx"]},
            {"id": "sha", "name": "주주간계약서", "instructions": "샘플 주주간계약서 체계(정의, 주식 인수, 의사결정, 양도제한, 매입·매도확약, 배당·워터폴, 해지)를 유지하고 입력된 지분구조·종류주식 조건을 반영한다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "주주·종류주식·조건", "type": "area"}],
             "samples": [P + "04_Equity투자자/투자자별/KLL리츠_주주간계약서(25365574.v22)_clean (1).docx"]},
            {"id": "trust3", "name": "위탁 3종(자산관리·자산보관·사무수탁)", "instructions": "자산관리위탁계약·자산보관위탁계약·일반사무수탁계약 3종을 각각 '# 계약서명'으로 구분해 작성한다. 보수 조항은 입력값을 쓰고 없으면 [  ]로 둔다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "리츠·수탁사·보수 조건", "type": "area"}],
             "samples": [P + "01_설립_인허가/정관_약정/*위탁계약서*.pdf"]},
            {"id": "aoi", "name": "정관", "instructions": "샘플 정관의 장·조 체계를 유지하고 종류주식·배당·자산운용 조항을 입력값으로 반영한다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "회사·주식·배당 조건", "type": "area"}],
             "samples": [P + "01_설립_인허가/정관_약정/KLL리츠_정관*.docx", P + "01_설립_인허가/정관_약정/20261001_검토메모_정관개정*.md"]},
            {"id": "loc", "name": "대출약정서·인수확약서", "instructions": "샘플 체계를 유지해 대출약정서 또는 인수확약서(LOC) 초안을 작성한다.",
             "fields": PROJECT_FIELDS + [{"key": "terms", "label": "대주·금액·금리·기간·선행조건", "type": "area"}],
             "samples": [STD + "*02담보대출약정서*.docx", STD + "*03인수확약서*.docx", STD + "*04인수확약서*.docx"]},
        ]},
    {
        "id": "schedule", "name": "일정관리", "desc": "캘린더(마케팅·마일스톤)·프로젝트 대시보드·이메일 요약·우선순위", "view": "schedule",
        "subtypes": [
            {"id": "timeline", "name": "마스터 일정표(역산)", "instructions": "거래종결일에서 역산해 주요 마일스톤(투심, 계약, 인가, 대출, 납입, 종결)과 선행조건(CP)을 일자별 표로 작성하고 크리티컬 패스를 표시한다. 마지막에 '캘린더 등록용' 표(일자|제목|구분)를 붙인다.",
             "fields": PROJECT_FIELDS + [{"key": "dates", "label": "확정 일자·제약 조건", "type": "area"}],
             "samples": [P + "07_일정_회의록/TIMELINE.md"]},
            {"id": "email", "name": "이메일 요약·업무 우선순위", "instructions": "붙여넣은 이메일(또는 업로드한 .eml)을 건별로 요약(보낸이·요청사항·기한)하고, 오늘 할 일을 긴급·중요도 기준 우선순위 표로 정리한다. 회신이 필요한 건은 회신 초안 한두 줄을 붙인다.",
             "fields": [{"key": "mails", "label": "이메일 본문 붙여넣기(여러 건 가능)", "type": "area"}],
             "samples": []},
            {"id": "minutes", "name": "회의록·액션아이템", "instructions": "회의 메모를 회의록(일시·참석·안건·논의·결정·액션아이템(담당·기한)) 형식으로 정리한다.",
             "fields": PROJECT_FIELDS + [{"key": "notes", "label": "회의 메모·녹취 요약", "type": "area"}],
             "samples": []},
        ]},
    {
        "id": "ic", "name": "기관별 투자심의", "desc": "기관별 Q&A 취합 및 대응 요약", "view": "ic",
        "subtypes": [
            {"id": "qa_summary", "name": "기관별 Q&A 대응 요약", "instructions": "Q&A 목록(아래 [Q&A DB])을 기관별로 묶어 질문 요지·답변 요지·근거자료·미결 사항 표로 정리하고, 공통 질문은 표준 답변으로 통합한다. 미답변 건은 답변 초안을 제시한다.",
             "fields": PROJECT_FIELDS + [{"key": "institution", "label": "기관(비우면 전체)", "type": "text"}],
             "samples": [P + "04_Equity투자자/QA_BANK.md", P + "05_대출업무/대주QA/*.md"], "include": "qa"},
            {"id": "qa_answer", "name": "개별 질문 답변 초안", "instructions": "질문에 대한 답변 초안을 근거 자료와 함께 작성한다. 수치는 출처를 달고, 확정적 표현('보장','확정')을 쓰지 않는다.",
             "fields": PROJECT_FIELDS + [{"key": "question", "label": "질문", "type": "area"}],
             "samples": [P + "04_Equity투자자/QA_BANK.md"]},
        ]},
    {
        "id": "procurement", "name": "발주관리", "desc": "PM·FM·시장·재무·세무·회계·감사·보험 발주(RFP→견적비교→품의)", "view": "procurement",
        "subtypes": [
            {"id": "rfp", "name": "RFP(제안요청서)", "instructions": "샘플 RFP 형식으로 과업 범위·일정·제출물·평가기준·견적 제출 양식을 작성한다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "category", "label": "용역 구분", "type": "text", "placeholder": "PM / FM / 시장 / 재무 / 세무 / 회계 / 감사 / 보험"},
                 {"key": "scope", "label": "과업 범위·기한", "type": "area"}],
             "samples": [P + "06_사내보고/RFP/*.docx"]},
            {"id": "compare", "name": "견적 비교표·품의문", "instructions": "아래 [견적 DB]의 업체별 견적을 비교표(금액·기간·범위·특이사항·평가점수)로 정리하고, 추천 업체와 사유를 담은 품의문(제목·목적·내용·소요예산·결론)을 작성한다.",
             "fields": PROJECT_FIELDS + [{"key": "rfp_id", "label": "RFP 번호(발주 목록의 #)", "type": "text"}],
             "samples": [P + "05_대출업무/약정서/20261028_기안_대출금입금요청공문_날인발송_v01.docx"], "include": "quotes"},
        ]},
    {
        "id": "registry", "name": "등기접수", "desc": "법인설립·증자·소유권이전·담보신탁 등기서류",
        "subtypes": [
            {"id": "checklist", "name": "등기 서류 체크리스트·신청 초안", "instructions": "등기 유형별(법인설립, 증자, 자본감소, 소유권이전, 기존 신탁말소, 담보신탁, 근저당 말소)로 신청인·필요서류·발급처·유효기간·세금(취득세 등)·접수 순서(연건 접수번호 순서)를 표로 작성하고, 대출약정상 인출 동시이행 서류(접수증 사본 등)와 연결해 표시한다.",
             "fields": PROJECT_FIELDS + [
                 {"key": "kinds", "label": "등기 유형", "type": "text", "placeholder": "소유권이전, 신탁말소, 담보신탁"},
                 {"key": "facts", "label": "당사자·부동산 표시·일자", "type": "area"}],
             "samples": [P + "01_설립_인허가/CHECKLIST.md"]},
        ]},
    {
        "id": "pmi", "name": "PMI", "desc": "운용 Package 구축(렌트롤·자산운용보고서·임대안내서)",
        "subtypes": [
            {"id": "rentroll", "name": "렌트롤 정리", "instructions": "업로드한 임대차 자료로 렌트롤 표(임차인|층|면적|보증금|월임대료|관리비|계약기간|만기|Free rent|비고)와 WALE·임대율 요약을 작성한다.",
             "fields": PROJECT_FIELDS, "samples": ["shared/강의자료/렌트롤실습_Clean/*.xlsx", P + "00_원본자료/processed/20260916_렌트롤_IM기준_v01.csv"]},
            {"id": "report", "name": "자산운용보고서", "instructions": "분기 자산운용보고서(자산개요, 임대현황, 수지, 주요 이슈, 향후 계획) 초안을 작성한다.",
             "fields": PROJECT_FIELDS + [{"key": "period", "label": "보고 기간·주요 수치", "type": "area"}],
             "samples": [P + "00_원본자료/processed/20260916_렌트롤_IM기준_v01.csv"]},
            {"id": "brochure", "name": "임대안내서", "instructions": "임차 유치용 임대안내서(입지·접근성, 건물 사양, 임대 가능 면적, 임대조건, 문의처) 문구를 작성한다.",
             "fields": PROJECT_FIELDS + [{"key": "facts", "label": "건물 사양·공실·조건", "type": "area"}],
             "samples": [P + "03_IM_투자심의/IM/20261015_IM_*_v03.pptx"]},
        ]},
    {
        "id": "distribution", "name": "배포", "desc": "반복 이메일 양식 저장·변수 채워 일괄 작성", "view": "distribution",
        "subtypes": [
            {"id": "template", "name": "새 이메일 양식 만들기", "instructions": "샘플 이메일(들)을 바탕으로 재사용 가능한 이메일 양식을 만든다. 바뀌는 값은 {{변수명}}(한글 가능)으로 치환한다. 첫 줄은 '제목: ...', 다음 줄부터 본문으로 출력한다.",
             "fields": [{"key": "purpose", "label": "양식 용도", "type": "text"}, {"key": "mails", "label": "참고할 기존 이메일 붙여넣기", "type": "area"}],
             "samples": []},
        ]},
]

MODULE_MAP = {m["id"]: m for m in MODULES}


def get_subtype(module_id: str, subtype_id: str) -> tuple[dict, dict]:
    m = MODULE_MAP[module_id]
    st = next(s for s in m["subtypes"] if s["id"] == subtype_id)
    return m, st
