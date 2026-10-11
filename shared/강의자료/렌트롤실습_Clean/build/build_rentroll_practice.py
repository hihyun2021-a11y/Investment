# -*- coding: utf-8 -*-
"""렌트롤(Rent Roll) 작성 실습용 Clean 엑셀 생성기.

- 모든 임차인·면적·금액·일자는 가상의 예시값이다(특정 자산·거래와 무관).
- 계산 칸은 엑셀 수식으로 넣어 실습자가 입력값을 바꾸면 결과가 바로 바뀐다.
실행: python3 build_rentroll_practice.py
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "20261011_실습_렌트롤작성실습_강의용Clean_v01.xlsx")

FONT = "맑은 고딕"
NAVY, HEAD, SHADE, INPUT, CALC, WARN = "012353", "DAE3F3", "F2F2F2", "FFF8E5", "EEF3EA", "FDE2E1"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def f(size=10, bold=False, color="1F1F1F", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(h):
    return PatternFill("solid", fgColor=h)


BASE_DATE = dt.date(2026, 12, 31)   # 가상 기준일
PY = 0.3025                          # ㎡ → 평

# ---------------------------------------------------------------- 가상 렌트롤 데이터
# (층, 구획, 임차인, 업종, 용도, 면적㎡, 개시, 만기, 보증금/평, 임대료/평/월, 관리비/평/월, 인상률, 인상주기(개월), RF잔여(개월), 보증금담보, 비고)
D = dt.date
ROWS = [
    ("B1", "A", "임차인 A", "3PL", "상온창고", 4000, D(2025, 3, 1), D(2032, 2, 29), 200000, 30000, 8000, 0.025, 12, 0, "우선수익권", "앵커 임차인"),
    ("1F", "A", "임차인 B", "유통", "상온창고", 3500, D(2024, 7, 1), D(2029, 6, 30), 180000, 32000, 8000, 0.030, 12, 0, "예금질권", ""),
    ("1F", "B", "임차인 C", "식품", "저온창고", 1500, D(2026, 1, 1), D(2031, 12, 31), 250000, 45000, 12000, 0.030, 12, 2, "예금질권", "Rent-Free 잔여"),
    ("1F", "C", "임차인 D", "근린생활", "근생", 200, D(2025, 5, 1), D(2030, 4, 30), 300000, 50000, 10000, 0.020, 12, 0, "없음", ""),
    ("2F", "전체", "임차인 A", "3PL", "상온창고", 5000, D(2025, 3, 1), D(2032, 2, 29), 200000, 30000, 8000, 0.025, 12, 0, "우선수익권", "앵커 임차인"),
    ("3F", "A", "임차인 E", "이커머스", "상온창고", 3000, D(2023, 10, 1), D(2028, 9, 30), 190000, 31000, 8000, 0.030, 12, 0, "예금질권", ""),
    ("3F", "B", "공실", "", "상온창고", 1000, None, None, 0, 0, 0, 0, 0, 0, "", "임대 마케팅 중"),
    ("4F", "전체", "임차인 F", "제조", "상온창고", 4500, D(2024, 1, 1), D(2027, 12, 31), 180000, 29000, 8000, 0.020, 12, 0, "예금질권", "만기 1년 이내"),
    ("5F", "A", "임차인 G", "3PL", "상온창고", 2800, D(2025, 9, 1), D(2030, 8, 31), 200000, 33000, 8000, 0.030, 12, 0, "예금질권", ""),
    ("5F", "B", "임차인 H", "물류", "상온창고", 1500, D(2026, 4, 1), D(2029, 3, 31), 200000, 34000, 8000, 0.030, 12, 1, "예금질권", "Rent-Free 잔여"),
    ("6F", "A", "임차인 A", "3PL", "사무", 800, D(2025, 3, 1), D(2032, 2, 29), 200000, 30000, 8000, 0.025, 12, 0, "우선수익권", "앵커 임차인"),
    ("6F", "B", "공실", "", "사무", 200, None, None, 0, 0, 0, 0, 0, 0, "", ""),
]

HEADERS = [
    # (키, 머리글, 폭, 구분: in=입력, calc=수식)
    ("no", "No", 5, "in"), ("fl", "층", 6, "in"), ("blk", "구획", 7, "in"), ("ten", "임차인", 11, "in"),
    ("ind", "업종", 9, "in"), ("use", "용도", 9, "in"),
    ("m2", "임대면적\n(㎡)", 10, "in"), ("py", "임대면적\n(평)", 10, "calc"),
    ("start", "계약 개시일", 11, "in"), ("end", "계약 만기일", 11, "in"), ("rem", "잔여기간\n(년)", 9, "calc"),
    ("dep_py", "보증금\n(원/평)", 10, "in"), ("dep", "보증금 총액\n(원)", 15, "calc"),
    ("rent_py", "월 임대료\n(원/평)", 10, "in"), ("rent", "월 임대료\n(원)", 13, "calc"),
    ("cam_py", "월 관리비\n(원/평)", 10, "in"), ("cam", "월 관리비\n(원)", 13, "calc"),
    ("annual", "연 임대료\n(원)", 15, "calc"), ("depx", "보증금 배수\n(월 임대료 대비)", 11, "calc"),
    ("esc", "인상률\n(연)", 8, "in"), ("cyc", "인상주기\n(개월)", 8, "in"), ("rf", "Rent-Free\n잔여(개월)", 9, "in"),
    ("sec", "보증금 담보", 11, "in"), ("memo", "비고", 16, "in"),
]
COL = {h[0]: i + 1 for i, h in enumerate(HEADERS)}
L = get_column_letter
INPUT_KEYS = ["no", "fl", "blk", "ten", "ind", "use", "m2", "start", "end", "dep_py", "rent_py", "cam_py",
              "esc", "cyc", "rf", "sec", "memo"]
MONEY = ("dep_py", "rent_py", "cam_py", "dep", "rent", "cam", "annual")


def c(key):
    return L(COL[key])


def write_rentroll(ws, rows, title, first_row=6, blank_rows=0, base_ref="$C$3", note=None):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = f(14, True, NAVY)
    ws["A3"] = "기준일"; ws["A3"].font = f(10, True)
    ws.merge_cells("A3:B3")
    ws["C3"] = BASE_DATE; ws["C3"].number_format = "yyyy-mm-dd"; ws["C3"].fill = fill(INPUT)
    ws["C3"].font = f(10, True); ws["C3"].border = BORDER; ws["C3"].alignment = CENTER
    ws["E3"] = "임대가능면적(㎡)"; ws["E3"].font = f(10, True)
    ws.merge_cells("E3:F3")
    ws["G3"].fill = fill(INPUT); ws["G3"].border = BORDER; ws["G3"].alignment = RIGHT
    ws["G3"].number_format = "#,##0.00"; ws["G3"].font = f(10, True)
    ws["I3"] = note or "노란 칸 = 입력값 / 연두 칸 = 수식(수정 금지) / 단위: 원, ㎡, 평(= ㎡ × 0.3025)"
    ws["I3"].font = f(9, color="595959")
    hr = first_row - 1
    for i, (k, h, w, kind) in enumerate(HEADERS, start=1):
        cell = ws.cell(hr, i, h)
        cell.font = f(9.5, True, NAVY); cell.fill = fill(HEAD); cell.alignment = CENTER; cell.border = BORDER
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[hr].height = 36
    n = len(rows) + blank_rows
    for k in range(n):
        r = first_row + k
        R = str(r)
        data = rows[k] if k < len(rows) else None
        vals = {"no": k + 1}
        if data:
            vals.update(dict(zip(["fl", "blk", "ten", "ind", "use", "m2", "start", "end", "dep_py", "rent_py",
                                  "cam_py", "esc", "cyc", "rf", "sec", "memo"], data)))
        for key in INPUT_KEYS:
            v = vals.get(key)
            if v is not None and v != "":
                ws[c(key) + R] = v
        ws[c("py") + R] = f'=IF({c("m2")}{R}="","",ROUND({c("m2")}{R}*0.3025,2))'
        ws[c("rem") + R] = f'=IF({c("end")}{R}="","",MAX(0,ROUND(({c("end")}{R}-{base_ref})/365,2)))'
        ws[c("dep") + R] = f'=IF({c("py")}{R}="","",{c("py")}{R}*{c("dep_py")}{R})'
        ws[c("rent") + R] = f'=IF({c("py")}{R}="","",{c("py")}{R}*{c("rent_py")}{R})'
        ws[c("cam") + R] = f'=IF({c("py")}{R}="","",{c("py")}{R}*{c("cam_py")}{R})'
        ws[c("annual") + R] = f'=IF({c("rent")}{R}="","",{c("rent")}{R}*12)'
        ws[c("depx") + R] = f'=IF(OR({c("rent")}{R}="",{c("rent")}{R}=0),"",ROUND({c("dep")}{R}/{c("rent")}{R},1))'
        for i, (key, h, w, kind) in enumerate(HEADERS, start=1):
            cell = ws.cell(r, i)
            cell.border = BORDER
            cell.font = f(9.5)
            cell.fill = fill(CALC if kind == "calc" else (INPUT if not data else "FFFFFF"))
            cell.alignment = LEFT if key == "memo" else (RIGHT if key in MONEY + ("m2", "py") else CENTER)
            if key in MONEY:
                cell.number_format = "#,##0"
        ws[c("start") + R].number_format = "yyyy-mm-dd"
        ws[c("end") + R].number_format = "yyyy-mm-dd"
        ws[c("m2") + R].number_format = "#,##0.00"
        ws[c("py") + R].number_format = "#,##0.00"
        ws[c("rem") + R].number_format = "0.00"
        ws[c("depx") + R].number_format = "0.0"
        ws[c("esc") + R].number_format = "0.0%"
        ws.row_dimensions[r].height = 18
    last = first_row + n - 1
    t = last + 1
    T = str(t)
    ws.cell(t, 1, "합계 / 가중평균")
    ws.merge_cells(start_row=t, start_column=1, end_row=t, end_column=6)
    rng = lambda key: f"{c(key)}{first_row}:{c(key)}{last}"
    for key in ("m2", "py", "dep", "rent", "cam", "annual"):
        ws[c(key) + T] = f"=SUM({rng(key)})"
    ws[c("rem") + T] = f'=IFERROR(SUMPRODUCT(N(+{rng("rem")}),N(+{rng("annual")}))/SUM({rng("annual")}),"")'
    ws[c("depx") + T] = f'=IFERROR(ROUND({c("dep")}{T}/{c("rent")}{T},1),"")'
    for i, (key, h, w, kind) in enumerate(HEADERS, start=1):
        cell = ws.cell(t, i)
        cell.font = f(9.5, True); cell.fill = fill(SHADE); cell.border = BORDER
        cell.alignment = CENTER if i <= 6 else RIGHT
        if key in MONEY:
            cell.number_format = "#,##0"
    ws[c("m2") + T].number_format = "#,##0.00"; ws[c("py") + T].number_format = "#,##0.00"
    ws[c("rem") + T].number_format = "0.00"; ws[c("depx") + T].number_format = "0.0"
    ws.cell(t + 1, 1, "※ 합계 행의 잔여기간 = 연 임대료 가중 WALE(년). 공실 행은 임대료가 0이라 가중치에서 빠진다.").font = f(8.5, color="7F7F7F")
    ws.freeze_panes = ws.cell(first_row, 5)
    dv = DataValidation(type="list", formula1='"우선수익권,예금질권,보증보험,없음"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng("sec"))
    dv2 = DataValidation(type="list", formula1='"상온창고,저온창고,근생,사무,기타"', allow_blank=True)
    ws.add_data_validation(dv2); dv2.add(rng("use"))
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return first_row, last, t


wb = Workbook()

# ---------------------------------------------------------------- 1. 안내
g = wb.active
g.title = "0_안내"
g.sheet_view.showGridLines = False
g.column_dimensions["A"].width = 3; g.column_dimensions["B"].width = 22; g.column_dimensions["C"].width = 95
lines = [
    ("title", "렌트롤(Rent Roll) 작성 실습 — 강의용 Clean 양식"),
    ("note", "모든 임차인·면적·금액·일자는 가상의 예시값입니다. 특정 자산·거래·당사자와 무관합니다."),
    ("gap", ""),
    ("h", "1. 시트 구성"),
    ("kv", ("1_RentRoll_예시", "가상 물류센터의 작성 예시. 노란 칸(입력)과 연두 칸(수식)으로 구분")),
    ("kv", ("2_요약", "임대율, WALE, 임차인 비중, 만기 스케줄, 보증금 담보 구분 — 1번 시트를 수식으로 집계")),
    ("kv", ("3_실습_빈양식", "수식이 미리 들어 있는 빈 렌트롤. 입력값만 채우면 계산이 자동으로 됨")),
    ("kv", ("4_실습_오류찾기", "오류가 일부러 섞인 렌트롤. 검증 열을 직접 만들어 오류를 찾는 실습")),
    ("gap", ""),
    ("h", "2. 작성 규칙"),
    ("kv", ("단위", "면적 ㎡(평 병기, 평 = ㎡ × 0.3025), 금액 원, 비율 %. 원본 단위가 다르면 비고에 적는다")),
    ("kv", ("한 행 = 한 계약 구획", "같은 임차인이 여러 층을 쓰면 층·구획별로 행을 나눈다(만기·조건이 다를 수 있음)")),
    ("kv", ("공실", "임차인 칸에 '공실'로 적고 면적만 입력한다(임대율 계산에 필요)")),
    ("kv", ("기준일", "잔여기간·WALE는 기준일(통상 거래종결일 또는 평가일)을 기준으로 계산한다")),
    ("kv", ("출처", "실무에서는 각 행에 근거 계약서(변경합의 포함)와 페이지를 기록한다")),
    ("gap", ""),
    ("h", "3. 주요 산식"),
    ("kv", ("임대율", "임대면적 합계 ÷ 임대가능면적  (Rent-Free 구간 포함 여부를 명시)")),
    ("kv", ("WALE(임대료 기준)", "Σ(잔여기간 × 연 임대료) ÷ Σ연 임대료")),
    ("kv", ("WALE(면적 기준)", "Σ(잔여기간 × 임대면적) ÷ Σ임대면적(공실 제외)")),
    ("kv", ("보증금 배수", "보증금 총액 ÷ 월 임대료 (몇 개월분인지)")),
    ("kv", ("Effective Rent", "(월 임대료 × (12 − 연간 Rent-Free 개월) ÷ 12) + 월 관리비")),
    ("gap", ""),
    ("h", "4. 실습 과제"),
    ("kv", ("과제 1", "2_요약 시트에서 기준일을 1년 뒤로 바꾸면 WALE가 얼마나 줄어드는지 확인하라")),
    ("kv", ("과제 2", "만기 스케줄에서 특정 연도에 만기가 몰려 있는지 찾고, 재계약 리스크를 한 문장으로 적어라")),
    ("kv", ("과제 3", "앵커 임차인(최대 비중) 이탈 시 임대율과 연 임대료 감소분을 계산하라")),
    ("kv", ("과제 4", "4_실습_오류찾기 시트에서 오류 4건을 찾고, 어떤 검증식으로 찾았는지 적어라")),
    ("kv", ("과제 5", "보증금 담보 구분(우선수익권·예금질권·없음)별 금액을 집계하고, 매수인·대주 관점의 의미를 설명하라")),
    ("gap", ""),
    ("h", "5. 강사용 정답 — 오류찾기 시트 (배포 시 이 표는 지울 것)"),
    ("kv", ("오류 ①", "No 3: 계약 만기일이 개시일보다 앞섬 → 검증식 =만기일<개시일")),
    ("kv", ("오류 ②", "No 6: 월 임대료 3,100원/평(0 하나 누락, 다른 행은 3만원대) → 평균 대비 ±50% 벗어남")),
    ("kv", ("오류 ③", "No 8: 보증금 배수 0.6개월(다른 행 6~7개월) → 보증금 단위 입력 오류 의심")),
    ("kv", ("오류 ④", "No 12: 3F-A 구획 중복 입력 → 임대면적 합계가 임대가능면적을 넘어 임대율이 100% 초과")),
]
r = 1
for kind, v in lines:
    if kind == "title":
        g.cell(r, 2, v).font = f(16, True, NAVY)
    elif kind == "note":
        g.cell(r, 2, v).font = f(10, color="C00000")
    elif kind == "h":
        g.cell(r, 2, v).font = f(11.5, True, NAVY)
    elif kind == "kv":
        a, b = v
        ca, cb = g.cell(r, 2, a), g.cell(r, 3, b)
        ca.font = f(10, True); ca.fill = fill(SHADE); ca.border = BORDER; ca.alignment = CENTER
        cb.font = f(10); cb.border = BORDER; cb.alignment = LEFT
        if a.startswith("오류"):
            cb.fill = fill(WARN)
    g.row_dimensions[r].height = 20 if kind != "title" else 28
    r += 1

# ---------------------------------------------------------------- 2. 예시 렌트롤
ws = wb.create_sheet("1_RentRoll_예시")
fr, lr, tr = write_rentroll(ws, ROWS, "Rent Roll — 가상 물류센터 (작성 예시)")
ws["G3"] = 28000.00
EX = "'1_RentRoll_예시'!"

# ---------------------------------------------------------------- 3. 요약
s = wb.create_sheet("2_요약")
s.sheet_view.showGridLines = False
for col, w in zip("ABCDEFG", [3, 26, 18, 18, 18, 18, 30]):
    s.column_dimensions[col].width = w
s["B1"] = "Rent Roll 요약 (1_RentRoll_예시 시트 집계)"; s["B1"].font = f(14, True, NAVY)
s["B2"] = "모든 값은 수식으로 계산된다. 기준일은 1_RentRoll_예시!C3에서 바꾼다."; s["B2"].font = f(9, color="595959")

def rg(key):
    return f"{EX}${c(key)}${fr}:${c(key)}${lr}"

kpi = [
    ("기준일", f"={EX}C3", "yyyy-mm-dd"),
    ("임대가능면적(㎡)", f"={EX}G3", "#,##0.00"),
    ("임대면적(㎡, 공실 제외)", f'=SUMIFS({rg("m2")},{rg("ten")},"<>공실")', "#,##0.00"),
    ("임대율", "=C7/C6", "0.0%"),
    ("WALE — 임대료 기준(년)", f'=SUMPRODUCT(N(+{rg("rem")}),N(+{rg("annual")}))/SUM({rg("annual")})', "0.00"),
    ("WALE — 면적 기준(년)", f'=SUMPRODUCT(N(+{rg("rem")}),{rg("m2")}*({rg("ten")}<>"공실"))/C7', "0.00"),
    ("연 임대료 합계(원)", f'=SUM({rg("annual")})', "#,##0"),
    ("보증금 합계(원)", f'=SUM({rg("dep")})', "#,##0"),
    ("평균 월 임대료(원/평, 임대면적 가중)", f'=SUM({rg("rent")})/SUMIFS({rg("py")},{rg("ten")},"<>공실")', "#,##0"),
]
s["B4"] = "① 핵심 지표"; s["B4"].font = f(11, True, NAVY)
for i, (k, fm, nf) in enumerate(kpi):
    rr = 5 + i
    a, b = s.cell(rr, 2, k), s.cell(rr, 3, fm)
    a.font = f(10, True); a.fill = fill(SHADE); a.border = BORDER; a.alignment = CENTER
    b.font = f(10, True, NAVY); b.fill = fill(CALC); b.border = BORDER; b.alignment = RIGHT; b.number_format = nf

# ② 임차인별 비중
s["B15"] = "② 임차인별 비중"; s["B15"].font = f(11, True, NAVY)
hdr = ["임차인", "임대면적(㎡)", "면적 비중", "연 임대료(원)", "임대료 비중", "비고"]
for j, h in enumerate(hdr):
    cell = s.cell(16, 2 + j, h); cell.font = f(10, True, NAVY); cell.fill = fill(HEAD); cell.border = BORDER; cell.alignment = CENTER
tenants = []
for row in ROWS:
    if row[2] != "공실" and row[2] not in tenants:
        tenants.append(row[2])
for i, tn in enumerate(tenants):
    rr = 17 + i
    s.cell(rr, 2, tn)
    s.cell(rr, 3, f'=SUMIFS({rg("m2")},{rg("ten")},B{rr})').number_format = "#,##0.00"
    s.cell(rr, 4, f"=C{rr}/$C$7").number_format = "0.0%"
    s.cell(rr, 5, f'=SUMIFS({rg("annual")},{rg("ten")},B{rr})').number_format = "#,##0"
    s.cell(rr, 6, f"=E{rr}/$C$11").number_format = "0.0%"
    s.cell(rr, 7, f'=IF(F{rr}=MAX($F$17:$F${16+len(tenants)}),"최대 임차인(앵커)","")')
    for j in range(6):
        cell = s.cell(rr, 2 + j); cell.border = BORDER; cell.font = f(10)
        cell.alignment = CENTER if j in (0, 5) else RIGHT
        if j in (1, 2, 3, 4):
            cell.fill = fill(CALC)
end_t = 16 + len(tenants)

# ③ 만기 스케줄
ms = end_t + 2
s.cell(ms, 2, "③ 연도별 만기 스케줄").font = f(11, True, NAVY)
hdr = ["만기 연도", "만기 면적(㎡)", "면적 비중", "만기 연 임대료(원)", "임대료 비중", "누적 임대료 비중"]
for j, h in enumerate(hdr):
    cell = s.cell(ms + 1, 2 + j, h); cell.font = f(10, True, NAVY); cell.fill = fill(HEAD); cell.border = BORDER; cell.alignment = CENTER
years = list(range(2027, 2034))
for i, y in enumerate(years):
    rr = ms + 2 + i
    s.cell(rr, 2, y)
    lo, hi = f'DATE(B{rr},1,1)', f'DATE(B{rr},12,31)'
    s.cell(rr, 3, f'=SUMIFS({rg("m2")},{rg("end")},">="&{lo},{rg("end")},"<="&{hi})').number_format = "#,##0.00"
    s.cell(rr, 4, f"=C{rr}/$C$7").number_format = "0.0%"
    s.cell(rr, 5, f'=SUMIFS({rg("annual")},{rg("end")},">="&{lo},{rg("end")},"<="&{hi})').number_format = "#,##0"
    s.cell(rr, 6, f"=E{rr}/$C$11").number_format = "0.0%"
    s.cell(rr, 7, f"=SUM($F${ms+2}:F{rr})").number_format = "0.0%"
    for j in range(6):
        cell = s.cell(rr, 2 + j); cell.border = BORDER; cell.font = f(10)
        cell.alignment = CENTER if j == 0 else RIGHT
        if j:
            cell.fill = fill(CALC)
end_m = ms + 1 + len(years)
s.conditional_formatting.add(f"F{ms+2}:F{end_m}", FormulaRule(formula=[f"F{ms+2}>=0.3"], fill=fill(WARN)))
s.cell(end_m + 1, 2, "※ 임대료 비중 30% 이상인 연도는 빨간색으로 표시(만기 집중 리스크).").font = f(8.5, color="7F7F7F")

# ④ 보증금 담보 구분
ds = end_m + 3
s.cell(ds, 2, "④ 보증금 담보 구분").font = f(11, True, NAVY)
hdr = ["담보 구분", "건수", "보증금(원)", "비중", "의미(예시)"]
for j, h in enumerate(hdr):
    cell = s.cell(ds + 1, 2 + j, h); cell.font = f(10, True, NAVY); cell.fill = fill(HEAD); cell.border = BORDER; cell.alignment = CENTER
meaning = {"우선수익권": "담보신탁상 임차인이 대주보다 앞선 순위일 수 있음",
           "예금질권": "임대인 명의 예금계좌에 예치 후 임차인에게 질권 설정",
           "보증보험": "보증기관이 보증금 반환을 보증",
           "없음": "별도 담보 없음(일반채권)"}
for i, kind in enumerate(meaning):
    rr = ds + 2 + i
    s.cell(rr, 2, kind)
    s.cell(rr, 3, f'=COUNTIFS({rg("sec")},B{rr})')
    s.cell(rr, 4, f'=SUMIFS({rg("dep")},{rg("sec")},B{rr})').number_format = "#,##0"
    s.cell(rr, 5, f"=IFERROR(D{rr}/$C$12,0)").number_format = "0.0%"
    s.cell(rr, 6, meaning[kind])
    for j in range(5):
        cell = s.cell(rr, 2 + j); cell.border = BORDER; cell.font = f(10)
        cell.alignment = CENTER if j in (0, 1) else (LEFT if j == 4 else RIGHT)
        if j in (1, 2, 3):
            cell.fill = fill(CALC)

# ---------------------------------------------------------------- 4. 빈 양식
b = wb.create_sheet("3_실습_빈양식")
write_rentroll(b, [], "Rent Roll — 실습용 빈 양식", blank_rows=15,
               note="노란 칸에 계약서 기준으로 입력 / 연두 칸은 수식 / 기준일·임대가능면적(G3)을 먼저 입력")
b["C3"] = None

# ---------------------------------------------------------------- 5. 오류 찾기
ERR = [list(x) for x in ROWS]
ERR[2][7] = D(2025, 12, 31)          # ① 만기 < 개시 (No 3)
ERR[5][9] = 3100                     # ② 임대료 0 누락 (No 6)
ERR[7][8] = 18000                    # ③ 보증금 단위 오류 → 배수 0.6개월 (No 8)
dup = list(ROWS[5]); dup[15] = ""
ERR.append(tuple(dup))               # ④ 3F-A 중복 (No 13 → 정답표 No 12 표기 보정)
ERR = [tuple(x) for x in ERR]
e = wb.create_sheet("4_실습_오류찾기")
fr2, lr2, tr2 = write_rentroll(e, ERR, "Rent Roll — 오류 찾기 실습 (오류 4건 포함)",
                               note="오른쪽 '검증' 열에 직접 검증식을 넣어 오류를 찾아보세요. 예) 만기<개시, 평균 대비 ±50%, 보증금 배수, 면적 합계")
e["G3"] = 28000.00
vc = len(HEADERS) + 2
for j, h in enumerate(["검증①\n만기<개시", "검증②\n임대료 이상", "검증③\n보증금 배수", "검증④\n구획 중복"]):
    cell = e.cell(fr2 - 1, vc + j, h)
    cell.font = f(9.5, True, "7F6000"); cell.fill = fill(INPUT); cell.alignment = CENTER; cell.border = BORDER
    e.column_dimensions[L(vc + j)].width = 12
    for rr in range(fr2, lr2 + 1):
        x = e.cell(rr, vc + j); x.border = BORDER; x.fill = fill("FFFFFF")
e.cell(tr2 + 2, 1, f'임대율 점검: =SUM({c("m2")}{fr2}:{c("m2")}{lr2})/G3  → 100%를 넘으면 중복·오입력 의심').font = f(9, color="7F6000")

# 정답 표의 오류 ④ 번호를 실제 행 번호(No 13)로 맞춘다
for row in g.iter_rows(min_col=3, max_col=3):
    for cell in row:
        if isinstance(cell.value, str) and cell.value.startswith("No 12: 3F-A"):
            cell.value = cell.value.replace("No 12", "No 13")

# ---------------------------------------------------------------- 문서 속성
wb.properties.creator = ""
wb.properties.lastModifiedBy = ""
wb.properties.title = "렌트롤 작성 실습(Clean)"
wb.properties.keywords = "Rent Roll, 실습, 강의용, 가상데이터"
wb.save(OUT)
print("saved", OUT)
