# -*- coding: utf-8 -*-
"""데이터센터 실물투자 실습사례 — TM(Teaser Memorandum) 자산개요 1페이지.

수치 출처: PART2_데이터센터_실습파일_vf_2회차 개시.xlsx(실습).xlsx_완료.xlsx
  6.사업개요 및 사업일정계획 / 7.총 투자비의 산정 / 8.순영업소득의 산정(1) 임대조건
실행: python3 build_tm_asset_overview.py
"""
import os

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "20261011_실습_TM데이터센터_자산개요_강의용_v01.pptx")

F_L, F_M, F_B = "KoPubWorld돋움체 Light", "KoPubWorld돋움체 Medium", "KoPubWorld돋움체 Bold"
NAVY, NAVY2, BAND = "012353", "002060", "345B86"
INK, GREY, LGREY = "1F1F1F", "595959", "8C8C8C"
HEAD, SHADE, LINE = "DAE3F3", "F2F2F2", "BFBFBF"
W, H = 13.333, 7.5
L, R = 0.45, 13.333 - 0.45

# ---------------------------------------------------------------- 입력값 (실습파일 기준)
PY = 0.3025
SITE, GFA, WHITE, SUPPORT = 6000.0, 43000.0, 30000.0, 13000.0      # ㎡
MW_IN, EFF, IT_LOAD, PUE = 40.0, 0.90, 26.0, 1.385                  # MW, -, MW, -
PRICE_RE, PRICE_ME = 5200.0, 2000.0                                  # 억원 (부동산, 기계장치)
LAND, BLDG = 1200.0, 4000.0                                          # 억원
TENANTS = [("임차인 A (하이퍼스케일, 1~5차)", 22000), ("임차인 B (통신·금융)", 800),
           ("임차인 C (통신, 소버린)", 2000), ("공실 (Ramp-up)", 1200)]  # kW
RATE = 220  # 천원/kW·월

MW_EFF = MW_IN * EFF
PRICE = PRICE_RE + PRICE_ME
LEASED = sum(k for n, k in TENANTS if not n.startswith("공실"))
assert abs(LEASED + 1200 - IT_LOAD * 1000) < 1e-6


def rgb(h):
    return RGBColor.from_string(h)


def _font(run, size, font=F_L, color=INK, bold=False):
    run.font.size = Pt(size); run.font.name = font; run.font.bold = bold
    run.font.color.rgb = rgb(color)
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", font)


def _no_style(shp):
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)


def text(slide, x, y, w, h, lines, size=11, font=F_L, color=INK, align=PP_ALIGN.LEFT,
         fill=None, line=None, shape=None, anchor=MSO_ANCHOR.MIDDLE, margin=0.05):
    if shape is None:
        shp = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        _no_style(shp)
        if fill:
            shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line); shp.line.width = Pt(0.75)
        else:
            shp.line.fill.background()
    tf = shp.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin + 0.03)
    tf.margin_top = tf.margin_bottom = Inches(margin)
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        opt = {}
        if isinstance(ln, tuple):
            ln, opt = ln
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opt.get("align", align)
        p.space_after = Pt(opt.get("after", 1))
        r = p.add_run(); r.text = ln
        _font(r, opt.get("size", size), opt.get("font", font), opt.get("color", color))
    return shp


def _border(cell):
    tcpr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcpr, qn(tag)); ln.set("w", str(int(Pt(0.5))))
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", LINE)


def table(slide, x, y, w, rows, col_w, row_h=0.3, size=9.5, head=True, key_col=True,
          total_last=False, aligns=None):
    nr, nc = len(rows), len(rows[0])
    shp = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr))
    t = shp.table
    t._tbl.tblPr.set("firstRow", "0"); t._tbl.tblPr.set("bandRow", "0")
    tot = sum(col_w)
    for j, cw in enumerate(col_w):
        t.columns[j].width = Inches(w * cw / tot)
    for i in range(nr):
        t.rows[i].height = Inches(row_h)
        for j in range(nc):
            c = t.cell(i, j)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.01)
            is_h = head and i == 0
            is_k = key_col and j == 0 and not is_h
            is_t = total_last and i == nr - 1
            c.fill.solid()
            c.fill.fore_color.rgb = rgb(HEAD if is_h else (SHADE if (is_k or is_t) else "FFFFFF"))
            p = c.text_frame.paragraphs[0]
            c.text_frame.word_wrap = True
            if is_h or is_k:
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = (aligns[j] if aligns else PP_ALIGN.LEFT)
            r = p.add_run(); r.text = str(rows[i][j])
            _font(r, size, F_B if (is_h or is_t) else (F_M if is_k else F_L), NAVY2 if is_h else INK)
            _border(c)
    return shp


prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
s = prs.slides.add_slide(prs.slide_layouts[6])

# ---------------------------------------------------------------- 머리말
text(s, L, 0.28, 0.55, 0.42, [("01", {"font": F_B, "size": 16, "color": "FFFFFF", "align": PP_ALIGN.CENTER})],
     shape=MSO_SHAPE.RECTANGLE, fill=NAVY)
text(s, L + 0.65, 0.22, 7.5, 0.28, [("Teaser Memorandum", {"size": 9.5, "color": LGREY})])
text(s, L + 0.65, 0.46, 9.0, 0.36, [("자산개요", {"font": F_B, "size": 18, "color": NAVY})])
text(s, R - 3.6, 0.28, 3.6, 0.3, [("수도권 데이터센터 실물투자 (강의 실습사례)", {"size": 9, "color": LGREY, "align": PP_ALIGN.RIGHT})])
ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(L), Inches(0.92), Inches(R), Inches(0.92))
_no_style(ln); ln.line.color.rgb = rgb(NAVY); ln.line.width = Pt(1.5)

text(s, L, 1.0, R - L, 0.62, [
    (f"수도권 소재 연면적 {GFA:,.0f}㎡, 수전용량 {MW_IN:.0f}MW 규모의 운영형 데이터센터 — "
     f"IT-Load {IT_LOAD:.0f}MW 중 {LEASED/1000:.1f}MW({LEASED/(IT_LOAD*1000):.1%}) 임대 계약",
     {"font": F_M, "size": 14}),
    (f"• 하이퍼스케일 임차인 비중 {22000/(IT_LOAD*1000):.1%}, 매입가 {PRICE:,.0f}억원(부동산 {PRICE_RE:,.0f}억 + 기계장치 {PRICE_ME:,.0f}억), "
     f"수전 MW당 {PRICE/MW_IN:,.0f}억원", {"size": 11, "color": GREY}),
])

# ---------------------------------------------------------------- 좌: 자산개요 표
rows = [
    ["구분", "내용"],
    ["자산명", "수도권 데이터센터 (가칭)"],
    ["위치", "수도권"],
    ["대지면적", f"{SITE:,.0f}㎡ ({SITE*PY:,.0f}평)"],
    ["연면적", f"{GFA:,.0f}㎡ ({GFA*PY:,.0f}평)"],
    ["전산실 / 지원면적", f"{WHITE:,.0f}㎡ ({WHITE*PY:,.0f}평) / {SUPPORT:,.0f}㎡ ({SUPPORT*PY:,.0f}평)"],
    ["수전용량 (계약전력)", f"{MW_IN:.0f}MW (효율 {EFF:.0%} 적용 시 유효 {MW_EFF:.0f}MW)"],
    ["확약 IT-Load", f"{IT_LOAD:.0f}MW (임대 {LEASED/1000:.1f}MW · 공실 1.2MW)"],
    ["PUE", f"{PUE:.3f}"],
    ["임대 방식", f"kW 기반 상면임대, {RATE}천원/kW·월 · 전력요금 Pass-Through"],
    ["주요 임차인", "하이퍼스케일 1개사(5개 계약) 외 통신·금융 2개사"],
    ["매입가", f"{PRICE:,.0f}억원 (토지 {LAND:,.0f} · 건물 {BLDG:,.0f} · 기계장치 {PRICE_ME:,.0f})"],
    ["주요 일정", "매매계약 2025.12 → 소유권이전 2026.03 → 운영개시 2026.05"],
    ["운영기간", "60개월 (Exit 2031.05 가정)"],
]
table(s, L, 1.75, 6.75, rows, [1.6, 4.4], row_h=0.355, size=9.5)

# ---------------------------------------------------------------- 우: 핵심 스펙 타일
RX, RW = 7.4, R - 7.4
tiles = [("수전용량", f"{MW_IN:.0f}MW", f"유효 {MW_EFF:.0f}MW"),
         ("IT-Load", f"{IT_LOAD:.0f}MW", f"임대율 {LEASED/(IT_LOAD*1000):.1%}"),
         ("연면적", f"{GFA:,.0f}㎡", f"{GFA*PY:,.0f}평"),
         ("PUE", f"{PUE:.2f}", "전력 효율 지표")]
tw = (RW - 0.3) / 4
for i, (k, v, d) in enumerate(tiles):
    x = RX + i * (tw + 0.1)
    text(s, x, 1.75, tw, 0.28, [(k, {"font": F_B, "size": 10, "color": "FFFFFF", "align": PP_ALIGN.CENTER})],
         shape=MSO_SHAPE.RECTANGLE, fill=BAND)
    text(s, x, 2.03, tw, 0.82, [(v, {"font": F_B, "size": 18, "color": NAVY, "align": PP_ALIGN.CENTER}),
                               (d, {"size": 8.5, "color": GREY, "align": PP_ALIGN.CENTER})],
         shape=MSO_SHAPE.RECTANGLE, fill="FFFFFF", line=LINE)

# 전력 용량 구성
text(s, RX, 3.0, RW, 0.28, [("전력 용량 구성", {"font": F_B, "size": 10, "color": "FFFFFF"})],
     shape=MSO_SHAPE.RECTANGLE, fill=BAND)
table(s, RX, 3.3, RW, [
    ["구분", "용량(MW)", "비고"],
    ["계약전력(수전)", f"{MW_IN:.1f}", "계약전력(수전) 기준"],
    ["유효 수전용량", f"{MW_EFF:.1f}", f"효율 {EFF:.0%} 적용"],
    ["시설 총 전력수요", f"{IT_LOAD*PUE:.1f}", f"IT-Load × PUE {PUE:.3f} — 유효 수전용량과 거의 같음 [추정]"],
    ["확약 IT-Load", f"{IT_LOAD:.1f}", "임차인 계약전력 합계"],
], [1.6, 1.0, 2.9], row_h=0.3, size=9, aligns=[PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.LEFT])

# 임차 구성
text(s, RX, 4.92, RW, 0.28, [("임차 구성 (kW 기준)", {"font": F_B, "size": 10, "color": "FFFFFF"})],
     shape=MSO_SHAPE.RECTANGLE, fill=BAND)
trows = [["임차인", "계약전력(kW)", "비중"]]
for n, k in TENANTS:
    trows.append([n, f"{k:,}", f"{k/(IT_LOAD*1000):.1%}"])
trows.append(["합계", f"{IT_LOAD*1000:,.0f}", "100.0%"])
table(s, RX, 5.22, RW, trows, [3.0, 1.4, 1.1], row_h=0.285, size=9, total_last=True,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER])

# ---------------------------------------------------------------- 꼬리말
ln2 = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(L), Inches(7.05), Inches(R), Inches(7.05))
_no_style(ln2); ln2.line.color.rgb = rgb(LINE); ln2.line.width = Pt(0.5)
text(s, L, 7.07, R - L - 0.6, 0.32, [
    ("출처: PART2 데이터센터 실습파일(6.사업개요 · 7.총 투자비 · 8.임대조건), 작성기준일 2025.12.31 · 강의 실습용 가상사례로 특정 자산과 무관 · "
     "[추정] 시설 총 전력수요는 IT-Load × PUE 단순 산식", {"size": 7.5, "color": LGREY})])
text(s, R - 0.5, 7.07, 0.5, 0.32, [("1", {"size": 8, "color": LGREY, "align": PP_ALIGN.RIGHT})])

cp = prs.core_properties
cp.author = ""; cp.last_modified_by = ""; cp.title = "TM 자산개요 — 데이터센터 실습사례"
prs.save(OUT)
print("saved", OUT)
