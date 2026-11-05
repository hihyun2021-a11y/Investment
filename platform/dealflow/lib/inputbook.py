# -*- coding: utf-8 -*-
"""딜입력.xlsx — 신규 딜 하나에 대한 입력 양식 만들기·읽기.

시트: 안내 / 자산개요 / 사업일정 / WGL / 실사업체 / 재무가정 / 렌트롤
- 자산개요의 '자동추출' 열은 프로그램이 대장·등기부에서 읽은 값을 적는 곳이다. '입력값'이 있으면 입력값이 우선한다.
- 재무가정의 '실사반영' 열을 채우면 재무모델이 그 값을 쓴다(빈칸이면 '협의안').
"""
from __future__ import annotations

import csv
import pathlib

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .common import load_json, num, to_date

FONT = "맑은 고딕"
HEAD = PatternFill("solid", fgColor="D9D9D9")
COL1 = PatternFill("solid", fgColor="F2F2F2")
INPUT = PatternFill("solid", fgColor="FFF2CC")
AUTO = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SECTORS = ["물류센터", "오피스", "리테일", "호텔", "데이터센터", "주거", "기타"]

ASSET = [  # (key, 항목, 비고) — 자동추출 대상은 비고에 표시
    ("code", "딜 코드명", "영문 대문자-하이픈 예: ANYANG-LOGIS"),
    ("name", "자산명", "예: 안양물류센터"),
    ("sector", "자산유형", "목록 선택: " + "/".join(SECTORS)),
    ("addr", "소재지(도로명)", ""),
    ("jibun", "대지지번", "자동추출: 토지대장·건축물대장"),
    ("land_m2", "대지면적(㎡)", "자동추출"),
    ("gfa_m2", "연면적(㎡)", "자동추출: 건축물대장"),
    ("floors", "규모(지하/지상)", "자동추출 예: B2/8F"),
    ("structure", "주구조", "자동추출"),
    ("zoning", "용도지역", "자동추출"),
    ("use", "주용도", "자동추출"),
    ("approval", "사용승인일", "자동추출 YYYY-MM-DD"),
    ("owner", "소유자(매도인)", "자동추출: 등기부 갑구"),
    ("encumbrance", "제한물권(을구)", "자동추출: 근저당 등 — 말소 여부 확인 필요"),
    ("leasable_py", "임대면적(평)", "빈칸이면 렌트롤 합계"),
    ("seller_advisor", "매각자문사", ""),
    ("price", "매입(희망)가(억원)", ""),
    ("appraisal", "감정가(억원)", "없으면 빈칸"),
    ("vehicle", "매수 비히클", "예: 신설 사모 위탁관리리츠(AMC 코람코자산신탁)"),
    ("term", "운용기간(년)", ""),
    ("deal_type", "거래 방식", "예: 공개입찰 / 수의계약"),
    ("highlight1", "투자포인트 1", "Teaser 헤드라인에 쓰임"),
    ("highlight2", "투자포인트 2", ""),
    ("highlight3", "투자포인트 3", ""),
    ("location", "입지 설명", "교통·배후수요 2~3줄"),
    ("manager", "담당자 성명", "RFP·메일 서명"),
    ("manager_title", "담당자 직급", ""),
    ("team", "담당 부서", "예: 리츠투자부문/상장리츠본부/상장리츠팀"),
    ("manager_tel", "담당자 전화", ""),
    ("manager_email", "담당자 이메일", "이 주소의 도메인을 '사내'로 보고 메일을 분류"),
    ("company", "회사명", "기본 코람코자산신탁"),
]

WGL_COLS = ["구분", "기관명", "역할", "성명", "직급", "부서", "이메일", "전화", "휴대폰", "비고"]
WGL_GROUPS = ["매수인(당사)", "매도인", "매각자문사", "Equity 투자자", "대주", "법무법인", "회계법인",
              "감정평가법인", "시장실사", "물리실사", "자산보관·사무수탁", "PM·FM", "임차인", "기타"]
VENDOR_COLS = ["실사구분", "업체명", "담당자", "이메일", "전화", "견적(백만원)", "기간(주)", "제안 요지", "선정(O)", "선정 사유"]
DD_TYPES = ["시장", "재무", "물리", "감정", "법률"]
SCHED_COLS = ["키", "마일스톤", "일자", "비고"]

ASSUME = [  # (key, 항목, 단위, 기본 협의안, 비고)
    ("price", "매입가", "억원", None, "빈칸이면 자산개요 매입(희망)가"),
    ("acq_tax", "취득세 등(매입가 대비)", "%", 4.6, "중과·감면 여부 재무실사 확인"),
    ("other_cost", "기타 매입 부대비용(매입가 대비)", "%", 1.5, "실사·자문·중개 등"),
    ("amc_acq", "AMC 매입보수(매입가 대비)", "%", 1.0, ""),
    ("reserve", "운영 예비비", "억원", 10, ""),
    ("senior_amt", "선순위 대출", "억원", None, "빈칸이면 매입가 × 55%"),
    ("senior_rate", "선순위 금리(All-in)", "%", 5.5, ""),
    ("senior_fee", "선순위 취급수수료", "%", 1.0, "대출금 대비, 최초 1회"),
    ("mezz_amt", "중순위 대출", "억원", 0, ""),
    ("mezz_rate", "중순위 금리(All-in)", "%", 6.5, ""),
    ("mezz_fee", "중순위 취급수수료", "%", 1.0, ""),
    ("hold", "보유기간", "년", 5, "1~15"),
    ("rent_growth", "임대료 상승률(렌트롤 인상률 없을 때)", "%", 2.5, ""),
    ("market_rent", "만기 후 시장임대료", "원/평/월", None, "빈칸이면 현 임대료 수준 유지"),
    ("market_growth", "시장임대료 상승률", "%", 2.5, ""),
    ("downtime", "만기 시 공실기간", "개월", 3, "최초 만기 1회 적용"),
    ("vacancy", "일반 공실·대손", "%", 1.0, "임대료 대비"),
    ("opex_ratio", "운영비(관리비 수입 대비)", "%", 100, "관리비 수지 0 가정이면 100"),
    ("fixed_cost", "보유세·보험·기타 고정비", "억원/년", None, "빈칸이면 매입가 × 0.3%"),
    ("amc_fee", "AMC 운용보수(매입가 대비 연)", "%", 0.5, ""),
    ("capex", "CAPEX", "억원/년", 0, "물리실사 결과로 갱신"),
    ("exit_cap", "Exit Cap Rate", "%", None, "빈칸이면 Going-in Cap Rate"),
    ("sale_cost", "매각비용(매각가 대비)", "%", 1.0, ""),
    ("amc_sale", "AMC 매각보수(매각가 대비)", "%", 0.5, ""),
]

RR_COLS = ["층", "임차인", "용도", "임대면적(평)", "보증금(원/평)", "월임대료(원/평)", "관리비(원/평)",
           "계약개시", "임대차만료", "인상률(%)", "비고"]


def _style_head(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.fill = HEAD
        cell.font = Font(name=FONT, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _body(ws, r0, r1, ncol, input_cols=(), auto_cols=()):
    for r in range(r0, r1 + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BOX
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c in input_cols:
                cell.fill = INPUT
            elif c in auto_cols:
                cell.fill = AUTO


def create(path: pathlib.Path, sample: dict | None = None):
    """빈 딜입력.xlsx 생성. sample = 미리 채울 값 {sheet: ...} (안양 예시용)."""
    sample = sample or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "안내"
    guide = [
        "딜입력.xlsx — 신규 딜 1건의 입력 양식 (딜플로우 1단계)",
        "",
        "노란 칸 = 직접 입력, 초록 칸 = 프로그램이 자동으로 채움(대장·등기부·렌트롤에서 추출). 입력값이 있으면 입력값을 우선합니다.",
        "1) 자산개요: 자산명·유형·가격·담당자는 꼭 입력. 지번·면적·구조 등은 '01_입력자료/대장_등기'에 PDF를 넣으면 자동추출을 시도합니다.",
        "2) 사업일정: '거래종결' 날짜만 넣어도 안양 일정 샘플 간격으로 나머지 날짜를 [추정]으로 채웁니다. 아는 날짜는 직접 넣으세요.",
        "3) WGL: 참여기관 담당자 명단. 이메일 분류·주소록 팝업에 그대로 쓰입니다.",
        "4) 실사업체: 실사구분별 후보 업체(여러 줄 가능). 견적을 받고 '선정(O)'에 O를 넣으면 선정 보고장표·기안문이 만들어집니다.",
        "5) 재무가정: '협의안'으로 협의용 재무모델을 만들고, 실사 결과가 나오면 '실사반영' 열에 값과 출처를 넣어 다시 실행합니다.",
        "6) 렌트롤: 임차인별 1줄. 별도 렌트롤 파일(xlsx·csv)을 '01_입력자료'에 넣어도 됩니다(열 이름이 같으면 자동 인식).",
        "",
        "모든 산출물은 '02_산출물/1단계'에 날짜_유형_제목_v## 이름으로 저장되며, 기존 파일은 덮어쓰지 않습니다.",
        "법률·세무 관련 문구는 내부 검토 초안이며 외부 법무법인·회계법인 확인이 필요합니다.",
    ]
    for i, t in enumerate(guide, 1):
        ws.cell(i, 1, t).font = Font(name=FONT, size=11 if i > 1 else 13, bold=(i == 1))
    ws.column_dimensions["A"].width = 130

    # 자산개요
    ws = wb.create_sheet("자산개요")
    ws.append(["키", "항목", "입력값", "자동추출", "비고·출처"])
    _style_head(ws, 1, 5)
    av = sample.get("자산개요", {})
    for key, label, note in ASSET:
        ws.append([key, label, av.get(key), None, note])
    _body(ws, 2, ws.max_row, 5, input_cols=(3,), auto_cols=(4,))
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).fill = COL1
    dv = DataValidation(type="list", formula1='"' + ",".join(SECTORS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C{[k for k, *_ in ASSET].index('sector') + 2}")
    ws.column_dimensions["A"].hidden = True
    _widths(ws, [12, 24, 46, 34, 44])

    # 사업일정
    ws = wb.create_sheet("사업일정")
    ws.append(SCHED_COLS)
    _style_head(ws, 1, 4)
    sv = sample.get("사업일정", {})
    ms = sorted(load_json("milestones.json")["milestones"], key=lambda m: m["offset"])
    for m in ms:
        ws.append([m["key"], m["name"], sv.get(m["key"]), ""])
    _body(ws, 2, ws.max_row, 4, input_cols=(3,))
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).number_format = "yyyy-mm-dd"
    ws.column_dimensions["A"].hidden = True
    _widths(ws, [12, 38, 14, 40])

    # WGL
    ws = wb.create_sheet("WGL")
    ws.append(WGL_COLS)
    _style_head(ws, 1, len(WGL_COLS))
    for row in sample.get("WGL", []):
        ws.append(row)
    last = max(ws.max_row, 60)
    _body(ws, 2, last, len(WGL_COLS), input_cols=tuple(range(1, 11)))
    dv = DataValidation(type="list", formula1='"' + ",".join(WGL_GROUPS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"A2:A{last}")
    _widths(ws, [16, 22, 20, 10, 9, 16, 28, 15, 15, 20])
    ws.freeze_panes = "A2"

    # 실사업체
    ws = wb.create_sheet("실사업체")
    ws.append(VENDOR_COLS)
    _style_head(ws, 1, len(VENDOR_COLS))
    for row in sample.get("실사업체", []):
        ws.append(row)
    last = max(ws.max_row, 25)
    _body(ws, 2, last, len(VENDOR_COLS), input_cols=tuple(range(1, 11)))
    dv = DataValidation(type="list", formula1='"' + ",".join(DD_TYPES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"A2:A{last}")
    _widths(ws, [10, 20, 12, 26, 14, 12, 9, 34, 9, 30])
    ws.freeze_panes = "A2"

    # 재무가정
    ws = wb.create_sheet("재무가정")
    ws.append(["키", "항목", "단위", "협의안", "실사반영", "출처·비고"])
    _style_head(ws, 1, 6)
    fv = sample.get("재무가정", {})
    for key, label, unit, dflt, note in ASSUME:
        v = fv.get(key, dflt)
        dd = sample.get("실사반영", {}).get(key)
        ws.append([key, label, unit, v, dd[0] if dd else None, (dd[1] if dd else note)])
    _body(ws, 2, ws.max_row, 6, input_cols=(4, 5))
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).fill = COL1
    ws.column_dimensions["A"].hidden = True
    _widths(ws, [12, 34, 10, 12, 12, 44])

    # 렌트롤
    ws = wb.create_sheet("렌트롤")
    ws.append(RR_COLS)
    _style_head(ws, 1, len(RR_COLS))
    for row in sample.get("렌트롤", []):
        ws.append(row)
    last = max(ws.max_row, 40)
    _body(ws, 2, last, len(RR_COLS), input_cols=tuple(range(1, 12)))
    for r in range(2, last + 1):
        for c in (8, 9):
            ws.cell(r, c).number_format = "yyyy-mm-dd"
        for c in (4, 5, 6, 7):
            ws.cell(r, c).number_format = "#,##0.00" if c == 4 else "#,##0"
    _widths(ws, [7, 18, 12, 12, 13, 13, 12, 12, 12, 9, 30])
    ws.freeze_panes = "A2"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# ── 읽기 ───────────────────────────────────────────────────────────────
def _rows(ws, ncol):
    out = []
    for r in ws.iter_rows(min_row=2, max_col=ncol, values_only=True):
        if any(v not in (None, "") for v in r):
            out.append(list(r))
    return out


RR_ALIAS = {  # 외부 렌트롤 파일 열 이름 → 표준 열
    "층": ["층", "floor"], "임차인": ["임차인", "tenant"], "용도": ["용도"],
    "임대면적(평)": ["임대면적(평)", "임대면적_평", "면적(평)", "전용면적(평)"],
    "보증금(원/평)": ["보증금(원/평)", "보증금_원per평", "평당보증금"],
    "월임대료(원/평)": ["월임대료(원/평)", "월임대료_원per평", "평당임대료"],
    "관리비(원/평)": ["관리비(원/평)", "관리비_원per평", "평당관리비"],
    "계약개시": ["계약개시", "임대차개시", "개시일"],
    "임대차만료": ["임대차만료", "만기", "만료일"], "인상률(%)": ["인상률(%)", "인상률_pct", "인상률"],
    "비고": ["비고"],
}


def _rr_from_table(header, rows):
    idx = {}
    for std, al in RR_ALIAS.items():
        for a in al:
            if a in header:
                idx[std] = header.index(a)
                break
    out = []
    for r in rows:
        rec = {k: (r[i] if i < len(r) else None) for k, i in idx.items()}
        if rec.get("임차인") in (None, ""):
            continue
        out.append(rec)
    return out


def load_rentroll_file(p: pathlib.Path):
    if p.suffix.lower() == ".csv":
        with p.open(encoding="utf-8-sig") as f:
            rd = list(csv.reader(f))
        return _rr_from_table([h.strip() for h in rd[0]], rd[1:])
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    for i, r in enumerate(rows[:10]):           # 머리행 찾기
        if r and any(str(v).strip() == "임차인" for v in r if v is not None):
            return _rr_from_table([str(v).strip() if v is not None else "" for v in r], rows[i + 1:])
    return []


def normalize_rr(recs):
    out = []
    for r in recs:
        out.append({
            "floor": str(r.get("층") or ""), "tenant": str(r.get("임차인")).strip(), "use": str(r.get("용도") or ""),
            "area_py": num(r.get("임대면적(평)"), 0.0), "dep_py": num(r.get("보증금(원/평)"), 0.0),
            "rent_py": num(r.get("월임대료(원/평)"), 0.0), "cam_py": num(r.get("관리비(원/평)"), 0.0),
            "start": to_date(r.get("계약개시")), "expiry": to_date(r.get("임대차만료")),
            "esc": num(r.get("인상률(%)")), "note": str(r.get("비고") or ""),
        })
    return out


def load(path: pathlib.Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    d = {"asset": {}, "asset_auto": {}, "schedule_in": {}, "wgl": [], "vendors": [], "assume": {}, "assume_dd": {},
         "assume_src": {}, "rentroll": []}
    for k, _label, val, auto, _n in _rows(wb["자산개요"], 5):
        if k:
            if val not in (None, ""):
                d["asset"][k] = val
            if auto not in (None, ""):
                d["asset_auto"][k] = auto
    for k, _n, v, note in _rows(wb["사업일정"], 4):
        if k and to_date(v):
            d["schedule_in"][k] = to_date(v)
    for r in _rows(wb["WGL"], len(WGL_COLS)):
        rec = dict(zip(WGL_COLS, [("" if v is None else str(v).strip()) for v in r]))
        if rec["기관명"] or rec["성명"]:
            d["wgl"].append(rec)
    for r in _rows(wb["실사업체"], len(VENDOR_COLS)):
        rec = dict(zip(VENDOR_COLS, r))
        if rec["실사구분"] and rec["업체명"]:
            rec["선정"] = str(rec.get("선정(O)") or "").strip().upper() in ("O", "○", "Y", "1", "선정")
            rec["견적"] = num(rec.get("견적(백만원)"))
            rec["기간"] = num(rec.get("기간(주)"))
            d["vendors"].append(rec)
    for k, _l, _u, v, dd, src in _rows(wb["재무가정"], 6):
        if not k:
            continue
        d["assume"][k] = num(v)
        if dd not in (None, ""):
            d["assume_dd"][k] = num(dd)
            d["assume_src"][k] = src or ""
    rr = []
    ws = wb["렌트롤"]
    for r in _rows(ws, len(RR_COLS)):
        rr.append(dict(zip(RR_COLS, r)))
    d["rentroll"] = normalize_rr([r for r in rr if r.get("임차인")])
    return d


def write_auto(path: pathlib.Path, auto: dict):
    """자산개요 '자동추출' 열 기록(입력값 열은 건드리지 않음)."""
    wb = openpyxl.load_workbook(path)
    ws = wb["자산개요"]
    for r in range(2, ws.max_row + 1):
        k = ws.cell(r, 1).value
        if k in auto:
            v, src = auto[k]
            ws.cell(r, 4, v)
            ws.cell(r, 5, (ws.cell(r, 5).value or "").split(" ┃ ")[0] + f" ┃ {src}")
    wb.save(path)
