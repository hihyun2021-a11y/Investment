# -*- coding: utf-8 -*-
"""부가가치세(VAT) 대출 진행경과 사내 보고자료 (2장) 생성기.

- 양식: VAT 대출 IM(20261102_IM_부가가치세대출_수협은행_v01.pptx)과 같은 사내 표준 덱(ic_deck.py) — 도구 함수 동일
- 1p 우선 참여기관 및 T/C: 운용역 제공 현황(2026-11-03)
- 2p 진행일정: IM 3p 시계열 막대 도식 재사용 + 기관별 심의·승인일, 납부일, 조기환급일
- 법령: 부가가치세법 제52조 제4항·시행령 제95조(대리납부 기한), 제59조 제2항·시행령 제107조(조기환급 15일)
실행: python3 build_사내보고_VAT대출_진행경과_261103.py
"""
import datetime as dt
import os
import sys

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "../../../.."))
sys.path.insert(0, os.path.join(ROOT, "platform/scripts"))
import ic_deck  # noqa: E402
from ic_deck import (BAND, F_BOLD, F_LIGHT, F_MEDIUM, GRAY, INK, LINE, NAVY,  # noqa: E402
                     NAVY_DARK, RED, TH_BG, WHITE, ICDeck, _run)

OUT = os.path.join(HERE, "20261103_사내보고_VAT대출_진행경과_v01.pptx")
ORANGE = RGBColor(0xF7, 0x8E, 0x3F)
GREEN = RGBColor(0x00, 0x6D, 0x52)
LGRAY = RGBColor(0xBF, 0xBF, 0xBF)
LAYOUT = "7_흰배경"
HEADER = "VAT 대출"

d = ICDeck(asset="안양물류센터", meeting="부가가치세 대출 진행경과 보고",
           date="November. 2026")
ic_deck.SECTION_LAYOUT["VAT"] = LAYOUT

# 머리말 섹션명 변경(출력본 레이아웃만 수정, 템플릿 원본 불변)
for sh in d._layouts[LAYOUT].shapes:
    if sh.has_text_frame and sh.text_frame.text.strip() == "프로젝트 추진 배경":
        for p in sh.text_frame.paragraphs:
            if p.runs:
                p.runs[0].text = HEADER
                for r in p.runs[1:]:
                    r.text = ""
# 꼬리말(마스터) 문구: 투심 설명자료 → VAT 대출 진행경과 보고
for m in d.prs.slide_masters:
    for sh in m.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "예비투자심의위원회 설명자료" in r.text:
                        r.text = r.text.replace("예비투자심의위원회 설명자료", "부가가치세 대출 진행경과 보고")


def rect(slide, x, y, w, h, *, fill=None, line=None, dash=False, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
        if dash:
            s.line.dash_style = 4  # dash
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    return s


def text(shape_or_tf, lines, *, size=9, color=INK, font=F_LIGHT, align=PP_ALIGN.CENTER):
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, "text_frame") else shape_or_tf
    for i, ln in enumerate(lines):
        if isinstance(ln, str):
            ln = (ln, {})
        s, kw = ln
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        _run(p, s, font=kw.get("font", font), size=kw.get("size", size),
             bold=kw.get("bold", False), color=kw.get("color", color))
    return tf


def label(slide, x, y, w, h, lines, **kw):
    tb, tf = d._box(slide, x, y, w, h)
    text(tf, lines, **kw)
    return tb


def arrow(slide, x1, y1, x2, y2, *, color=GRAY, dash=False, lw=1.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(lw)
    if dash:
        c.line.dash_style = 4
    ln = c.line._get_or_add_ln()
    tail = ln.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd",
                          {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return c


def bold_cells(tbl, cells, color=None):
    for (r, c) in cells:
        for p in tbl.cell(r, c).text_frame.paragraphs:
            for run in p.runs:
                run.font.name = F_BOLD
                if color is not None:
                    run.font.color.rgb = color




def split_lines(tbl):
    """셀 안 줄바꿈(\n)을 별도 문단으로 나눠 가운데 정렬이 줄마다 적용되게 한다."""
    import copy as _c
    for row in tbl.rows:
        for cell in row.cells:
            p0 = cell.text_frame.paragraphs[0]
            if not p0.runs or "\n" not in p0.runs[0].text:
                continue
            parts = p0.runs[0].text.split("\n")
            p0.runs[0].text = parts[0]
            prev = p0._p
            for part in parts[1:]:
                np_ = _c.deepcopy(p0._p)
                for r in np_.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r")[1:]:
                    np_.remove(r)
                np_.find("{http://schemas.openxmlformats.org/drawingml/2006/main}r").find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = part
                prev.addnext(np_)
                prev = np_


def cell_fill(tbl, r, color):
    for c in range(len(tbl.columns)):
        tbl.cell(r, c).fill.solid()
        tbl.cell(r, c).fill.fore_color.rgb = color


AMT = 127.0                     # 억원
D_DRAW, D_REFUND, D_MAT = dt.date(2026, 11, 25), dt.date(2026, 12, 10), dt.date(2027, 2, 25)
DAYS_MAT = (D_MAT - D_DRAW).days          # 92일
DAYS_REF = (D_REFUND - D_DRAW).days       # 15일
HC_RATE, HC_MIN = 0.065, 45
hc_min_int = AMT * HC_RATE * HC_MIN / 365  # 현대커머셜 최소 이자(45일)
D_HC_MIN = D_DRAW + dt.timedelta(days=HC_MIN)   # 2027-01-09

# =====================================================================
# 1p. 우선 참여기관 및 T/C
# =====================================================================
s = d.content_slide("VAT", "1. 우선 참여기관 및 T/C",
                    lead="VAT 대출(127억원)은 우리은행·수협은행·현대커머셜 3개 기관으로 압축하여 11월 중순 승인을 추진 중임",
                    subs=["기존 대주 중 SBI저축은행·우리카드·신협중앙회·롯데캐피탈은 Drop — 승인 일정과 부대조건을 비교하여 최종 기관 선정 예정"])
d.band(s, "기관별 진행 현황 및 주요 조건")
rows = [
    ["구 분", "기 관", "승인 일정 / 절차", "금 리", "만 기", "부대조건", "진행 상태"],
    ["우선 ①", "우리은행", "~11/18 승인\n셀다운 기관 승인 시 지점전결", "[TBD]%", "3개월", "-", "협의 중"],
    ["우선 ②", "수협은행", "~11/13 신규 승인 심의", "[TBD]%", "3개월", "지점 여유자금 예치 요건\n(환급금 예치)", "심의 준비"],
    ["우선 ③", "현대커머셜", "신규 승인 필요\n(심의 일정 [TBD])", "[6.5]%", "3개월\n(조기상환 가능)", "최소 45일 잔액 유지", "신규 검토"],
    ["Drop", "SBI저축은행 · 우리카드 · 신협중앙회 · 롯데캐피탈", "-", "-", "-", "-", "Drop"],
]
tbl = d.table(s, rows, top=2.52, col_widths=[0.95, 2.15, 2.30, 0.95, 1.30, 1.95, 1.10], row_height=0.50, size=9)
tbl.rows[0].height = Inches(0.32)
for r in range(1, 5):
    for c in (3, 4, 6):
        tbl.cell(r, c).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for p in tbl.cell(r, 2).text_frame.paragraphs + tbl.cell(r, 5).text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
bold_cells(tbl, [(r, 1) for r in range(1, 4)], color=NAVY_DARK)
split_lines(tbl)
bold_cells(tbl, [(3, 3), (3, 5), (2, 5)], color=RED)
cell_fill(tbl, 4, RGBColor(0xF2, 0xF2, 0xF2))
for c in range(7):
    for p in tbl.cell(4, c).text_frame.paragraphs:
        for run in p.runs:
            run.font.color.rgb = GRAY

# 하단: 공통 조건 / 기관별 이자비용 비교
Y2 = 5.18
d.band(s, "공통 조건", top=Y2, width=5.05)
rows = [
    ["항 목", "내 용"],
    ["한도 · 용도", "127억원 / 건물분 부가가치세 대리납부"],
    ["담보", "환급청구권 양도담보 · 환급계좌 근질권 (각 120%)"],
    ["선행조건", "기존 대주단 사전 서면동의 (공문 v02 발송 예정)"],
    ["상환재원", "조기환급금 ('26.12.10 예정)"],
]
tbl = d.table(s, rows, top=Y2 + 0.36, width=5.05, col_widths=[1.25, 3.80], row_height=0.30, size=8.5)
bold_cells(tbl, [(r, 0) for r in range(1, 5)])

XR = 5.82
bnd = d.band(s, "이자비용 비교 (127억원 기준, 단위: 억원)", top=Y2, width=5.38)
bnd.left = Inches(XR)
rows = [
    ["기 관", "12/10 환급 후 상환", "만기(92일)까지 유지", "비 고"],
    ["우리은행", "127 × [r] × 15/365", "127 × [r] × 92/365", "조기상환 조건 [TBD]"],
    ["수협은행", "-", "127 × [r] × 92/365", "환급금 예치 → 예치이자 상계"],
    ["현대커머셜", f"{hc_min_int:.2f} (최소 45일)", f"{AMT*HC_RATE*DAYS_MAT/365:.2f}", "45일 잔액 유지 ('27.01.09)"],
]
tbl = d.table(s, rows, top=Y2 + 0.36, left=XR, width=5.38, col_widths=[1.05, 1.45, 1.45, 1.43], row_height=0.30, size=8)
bold_cells(tbl, [(r, 0) for r in range(1, 4)])
bold_cells(tbl, [(3, 1), (3, 2)], color=RED)
split_lines(tbl)
for r in range(1, 4):
    for c in (1, 2):
        tbl.cell(r, c).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
d.footnote(s, "* 기관별 일정·조건은 '26.11.03 운용역 협의 현황 기준. 이자비용은 단리·실일수/365 단순 계산 [추정], 취급수수료·예치이자 미반영. "
              "[r] = 확정 전 금리.", top=7.12)

# =====================================================================
# 2p. 진행일정 (시계열 막대 — IM 3p 도식)
# =====================================================================
s = d.content_slide("VAT", "2. 진행 일정",
                    lead="11/13~11/18 기관 승인 → 11/25 인출·대리납부·조기환급 신고 → 12/10 조기환급 순으로 진행함",
                    subs=["기관 확정과 기존 대주단 동의를 11/20 전후까지 마쳐야 납부기한(11/25) 내 인출 가능"])
d.band(s, "VAT 대출 진행 일정 (시계열)")

D0, D1 = dt.date(2026, 10, 26), dt.date(2027, 3, 5)
LBL_X, LBL_W = 0.57, 3.05
CX0, CX1 = 3.70, 11.15
TOPC = 2.50


def xpos(day):
    return CX0 + (CX1 - CX0) * (day - D0).days / (D1 - D0).days


ROWH = 0.32
tasks = [
    # (라벨, 보조, 시작, 끝, 색, 표시)
    ("거래종결 · 대리납부 의무 발생", "매매대금 지급(세금계산서 수취)", dt.date(2026, 10, 29), None, NAVY, "10.29"),
    ("기존 대주단 동의 (다수대주 2/3)", "공문 발송 → 동의서 회신", dt.date(2026, 11, 4), dt.date(2026, 11, 20), BAND, "11.04 ~ 11.20 [TBD]"),
    ("수협은행 신규 승인 심의", "지점 여유자금 예치 요건", dt.date(2026, 11, 4), dt.date(2026, 11, 13), BAND, "심의 ~ 11.13"),
    ("우리은행 승인", "셀다운 기관 승인 시 지점전결", dt.date(2026, 11, 4), dt.date(2026, 11, 18), BAND, "승인 ~ 11.18"),
    ("현대커머셜 신규 승인", "심의 일정 [TBD]", dt.date(2026, 11, 4), dt.date(2026, 11, 20), LGRAY, "[TBD]"),
    ("기관 확정 · 약정 체결 · 담보 설정", "계좌 개설, 양도요구서 확정일자", dt.date(2026, 11, 19), dt.date(2026, 11, 24), ORANGE, "11.19 ~ 11.24 [TBD]"),
    ("VAT 납부 (대출 인출 · 대리납부)", "10월분 조기환급 신고 동시 진행", D_DRAW, None, RED, "11.25 (납부기한)"),
    ("조기환급 심사 (15일)", "신고기한 경과 후 15일 이내", D_DRAW, D_REFUND, ORANGE, "~ 12.10"),
    ("조기환급일 · 대출 상환/예치", "현대커머셜은 '27.01.09까지 잔액 유지", D_REFUND, None, GREEN, "12.10 조기환급"),
    ("VAT 대출 기간 (3개월)", "만기 '27.02.25 / 일반환급 기한 '27.02.24", D_DRAW, D_MAT, NAVY, "11.25 ~ '27.02.25"),
]
NR = len(tasks)
for m in [dt.date(2026, 11, 1), dt.date(2026, 12, 1), dt.date(2027, 1, 1), dt.date(2027, 2, 1), dt.date(2027, 3, 1)]:
    x = xpos(m)
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(TOPC + 0.28), Inches(x), Inches(TOPC + 0.30 + NR * ROWH))
    ln.line.color.rgb = LINE
    ln.line.width = Pt(0.5)
    label(s, x - 0.45, TOPC, 0.9, 0.24, [(f"'{m:%y}.{m.month:02d}", {"font": F_MEDIUM, "size": 8.5, "color": NAVY_DARK})])

for i, (name, sub, a, b_, col, tag) in enumerate(tasks):
    y = TOPC + 0.32 + i * ROWH
    if i % 2 == 0:
        rect(s, LBL_X, y - 0.02, CX1 - LBL_X, ROWH - 0.02, fill=RGBColor(0xF7, 0xF7, 0xF7))
    label(s, LBL_X + 0.05, y, LBL_W, ROWH - 0.06,
          [(name, {"font": F_MEDIUM, "size": 8.5, "color": INK}), (sub, {"size": 7, "color": GRAY})],
          align=PP_ALIGN.LEFT)
    yc = y + (ROWH - 0.06) / 2
    if b_ is None:
        x = xpos(a)
        rect(s, x - 0.09, yc - 0.09, 0.18, 0.18, fill=col, shape=MSO_SHAPE.DIAMOND)
        label(s, x + 0.12, yc - 0.12, 1.6, 0.24, [(tag, {"font": F_BOLD, "size": 8, "color": col})], align=PP_ALIGN.LEFT)
    else:
        x1, x2 = xpos(a), xpos(b_)
        gray = (col == LGRAY)
        bar = rect(s, x1, yc - 0.09, x2 - x1, 0.18, fill=None if gray else col,
                   line=GRAY if gray else None, dash=gray)
        if x2 - x1 > 1.6:
            text(bar, [(tag, {"font": F_BOLD, "size": 7.5, "color": GRAY if gray else WHITE})])
        else:
            label(s, x2 + 0.06, yc - 0.12, 1.9, 0.24, [(tag, {"font": F_BOLD, "size": 7.5, "color": GRAY if gray else col})],
                  align=PP_ALIGN.LEFT)

# 기준선: 납부기한(11/25), 현대 45일(1/9), 만기(2/25)
for day, col, tag in ((D_DRAW, RED, "납부기한"), (D_HC_MIN, GRAY, "현대 45일"), (D_MAT, RED, "만기")):
    x = xpos(day)
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(TOPC + 0.28), Inches(x), Inches(TOPC + 0.30 + NR * ROWH))
    ln.line.color.rgb = col
    ln.line.width = Pt(0.75)
    ln.line.dash_style = 4

# 하단: 핵심 일정 요약 표
YB = TOPC + 0.42 + NR * ROWH
d.band(s, "핵심 일정 요약", top=YB)
rows = [
    ["대출 심의", "대출 승인", "VAT 납부", "조기환급", "만기"],
    ["수협 ~11/13 · 우리 ~11/18 · 현대 [TBD]", "11/13 ~ 11/18 (기관 확정 11/19 [TBD])", "'26.11.25 (수)", "'26.12.10 (목) 예정", "'27.02.25 (목)"],
]
tbl = d.table(s, rows, top=YB + 0.32, col_widths=[3.0, 2.6, 1.6, 1.9, 1.6], row_height=0.28, size=8.5)
for c in range(5):
    tbl.cell(1, c).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
bold_cells(tbl, [(1, 2), (1, 3)], color=RED)
d.footnote(s, "* 대리납부: 대가 지급일이 속하는 달의 다음 달 25일까지(부가가치세법 제52조 제4항·시행령 제95조). 조기환급: 신고기한 경과 후 15일 이내"
              "(제59조 제2항·시행령 제107조). 세무서 현장확인 등으로 지연 가능 — 세무 판단은 외부 확인 필요.", top=7.12)

p = d.save(OUT)
print("생성:", p, len(d.prs.slides._sldIdLst), "장")
