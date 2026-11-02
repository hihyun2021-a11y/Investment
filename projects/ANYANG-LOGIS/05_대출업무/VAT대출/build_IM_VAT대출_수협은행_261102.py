# -*- coding: utf-8 -*-
"""수협은행 앞 부가가치세(VAT) 담보대출 심의 요청용 IM (3장) 생성기.

- 양식: 사내 표준 덱(platform/scripts/ic_deck.py, 빈 템플릿) — 머리말 섹션명만 'VAT 대출 IM'으로 변경
- 근거:
  · 대주 구성·금리: 05_대출업무/약정서/20261019_대출약정_담보대출약정서_BKL_v38_clean.docx (별첨 대주별 약정금·이자율)
  · 추가차입·담보제공 제한: 담보대출약정서 제11조 제7항 (대주 사전 서면동의)
  · Equity·재원: 03_IM_투자심의/IM/20261015_IM_안양물류센터_InformationMemorandum_v03.pptx (투자구조도)
  · 담보 구조: 참고계약서 3종(코크렙제72호 VAT 대출, 하나은행 2026-04-27) 제6조 — 환급청구권 양도담보 + 환급계좌 예금근질권(각 120%)
  · VAT 일정: 부가가치세법 제52조 제4항·시행령 제95조(대리납부), 제59조 제2항·시행령 제107조(조기환급)
- T/C(한도 127억·금리 TBD·3개월·환급금 수협 예치)는 운용역 지시(2026-11-02)
실행: python3 build_IM_VAT대출_수협은행_261102.py
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

OUT = os.path.join(HERE, "20261102_IM_부가가치세대출_수협은행_v01.pptx")
ORANGE = RGBColor(0xF7, 0x8E, 0x3F)
GREEN = RGBColor(0x00, 0x6D, 0x52)
LGRAY = RGBColor(0xBF, 0xBF, 0xBF)
LAYOUT = "7_흰배경"
HEADER = "VAT 대출"

d = ICDeck(asset="안양물류센터", meeting="부가가치세 대출 Information Memorandum",
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
# 꼬리말(마스터) 문구: 투심 설명자료 → VAT 대출 IM
for m in d.prs.slide_masters:
    for sh in m.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "예비투자심의위원회 설명자료" in r.text:
                        r.text = r.text.replace("예비투자심의위원회 설명자료", "부가가치세 대출 IM (수협은행)")


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


# =====================================================================
# 1p. Term Sheet
# =====================================================================
s = d.content_slide("VAT", "1. 대출 개요 (Term Sheet)",
                    lead="안양물류센터 매입에 따른 건물분 부가가치세 대리납부 재원 127억원을 3개월 단기로 조달하고자 함",
                    subs=["상환재원은 부가가치세 조기환급금이며, 환급청구권 양도담보와 환급계좌 예금근질권(각 1순위, 120%)으로 회수 구조 확보"])
d.band(s, "주요 대출조건 (안)")
rows = [
    ["구 분", "주요 조건", "비 고"],
    ["차주", "주식회사 코람코라이프로지스위탁관리부동산투자회사 (KLL리츠)", "AMC: ㈜코람코자산신탁"],
    ["대주", "수협은행", "기존 담보대출 Tranche A-1 대주 (100억원)"],
    ["대출한도", "금 127억원 (일시 인출)", "본건 건물분 부가가치세 상당액"],
    ["자금용도", "본건 부동산 매입에 따른 건물분 부가가치세 대리납부", "부가가치세법 제52조 제4항, 매매계약 제1.5조"],
    ["대출금리", "[TBD]  (고정/변동 [TBD])", "이자지급: [매월 후취 / 만기 일시] [TBD]"],
    ["대출기간", "인출일로부터 3개월 (2026.11.25 ~ 2027.02.25 예정)", "만기 일시상환"],
    ["상환재원", "부가가치세 조기환급금 (2026.12.10 수령 예정)", "조기상환 시 수수료 면제 [협의]"],
    ["담보 ①", "부가가치세환급청구권 채권양도담보 제1순위 (담보한도 152.4억원, 대출금의 120%)",
     "국세환급금양도요구서·양도담보설정통지서 확정일자"],
    ["담보 ②", "부가세환급계좌(수협은행 개설) 예금채권근질권 제1순위 (채권최고액 152.4억원, 120%)",
     "근질권설정 통지(인출일) 및 승낙서(10영업일 내)"],
    ["예치 조건", "환급금 전액을 수협은행 부가세환급계좌에 여유자금운용 예치금으로 예치 → 만기 상환재원으로 사용",
     "예치금은 담보② 근질권 목적물"],
    ["인출선행조건", "기존 담보대출 대주단 사전 서면동의, 담보계약 체결, 세금계산서 사본 제출 등",
     "담보대출약정서 제11조 제7항 (추가차입·담보제공 제한)"],
    ["비용", "법률자문비용, 담보권 설정·해지 비용 등 차주 부담", "기존 담보대출 담보목적물과 중복 없음"],
]
tbl = d.table(s, rows, top=2.52, col_widths=[1.45, 5.65, 3.60], row_height=0.355, size=9)
bold_cells(tbl, [(r, 0) for r in range(1, len(rows))])
bold_cells(tbl, [(3, 1), (8, 1), (9, 1), (10, 1)], color=NAVY_DARK)
bold_cells(tbl, [(5, 1)], color=RED)
d.footnote(s, "* 담보 구조는 ㈜코크렙제72호위탁관리부동산투자회사 부가가치세 대출약정서(대주 하나은행, 2026.04.27) 제6조 및 부속 담보계약 2종을 준용함. "
              "[TBD]는 수협은행 여신심의 결과에 따라 확정.", top=7.12)

# =====================================================================
# 2p. 차주 정보 및 투자구조
# =====================================================================
s = d.content_slide("VAT", "2. 차주 정보 및 투자구조",
                    lead="차주(KLL리츠)는 '26.10.29 안양물류센터 매입을 완료했으며, 담보대출 1,980억원은 25개 대주로 구성됨",
                    subs=["수협은행은 Tranche A-1 대주(100억원)로 참여 중이며, 본건 VAT 대출은 기존 담보와 분리된 별도 담보로 제공"])
d.band(s, "차주 개요")
rows = [
    ["구 분", "내 용", "구 분", "내 용"],
    ["차주", "주식회사 코람코라이프로지스위탁관리부동산투자회사", "보유자산", "안양물류센터 (안양시 동안구 관양동 934 외, 연면적 28,881평)"],
    ["형태·AMC", "사모 위탁관리리츠 / ㈜코람코자산신탁", "매입가 / 감정가", "2,700억원 (평당 935만원) / 3,345억원 (미래새한, '26.09)"],
    ["모리츠", "㈜코람코라이프인프라리츠 (KLI, 상장리츠)", "임대 현황", "임대율 99.3%, 쿠팡 77.2% (WALE 6.7년)"],
]
tbl = d.table(s, rows, top=2.50, col_widths=[1.15, 3.95, 1.35, 4.25], row_height=0.27, size=9)
bold_cells(tbl, [(r, c) for r in range(1, 4) for c in (0, 2)])

# 좌: 투자구조도
TOP2 = 3.72
d.band(s, "투자구조도", top=TOP2, width=5.05)
X0 = 0.57
# Equity
b = rect(s, X0 + 1.40, TOP2 + 0.40, 2.40, 0.66, fill=RGBColor(0xEE, 0xF3, 0xFA), line=BAND)
text(b, [("Equity 988억원", {"font": F_BOLD, "size": 9, "color": NAVY_DARK}),
         ("1종 240 · 2종 500 · 3종 100 · 보통주 148", {"size": 7}),
         ("KLI 보통주·2종 / FI 1·2종 / 매도인 측 3종", {"size": 6.5, "color": GRAY})])
# 차주
b = rect(s, X0 + 1.55, TOP2 + 1.45, 2.10, 0.62, fill=NAVY)
text(b, [("차주", {"size": 8, "color": WHITE}),
         ("코람코라이프로지스리츠 (KLL)", {"font": F_BOLD, "size": 9.5, "color": WHITE})])
arrow(s, X0 + 2.60, TOP2 + 1.06, X0 + 2.60, TOP2 + 1.45, color=BAND)
# 담보대출
b = rect(s, X0 + 0.00, TOP2 + 1.30, 1.35, 0.92, fill=RGBColor(0xF2, 0xF2, 0xF2), line=GRAY)
text(b, [("담보대출 1,980억원", {"font": F_BOLD, "size": 8.5, "color": NAVY_DARK}),
         ("A-1  900억 (21개사)", {"size": 7.5}),
         ("A-2  700억 (우리은행)", {"size": 7.5}),
         ("B     380억 (3개사)", {"size": 7.5})])
arrow(s, X0 + 1.35, TOP2 + 1.76, X0 + 1.55, TOP2 + 1.76, color=GRAY)
# VAT 대출(신규)
b = rect(s, X0 + 3.85, TOP2 + 1.30, 1.20, 0.92, fill=RGBColor(0xFE, 0xF4, 0xEA), line=ORANGE, dash=True, lw=1.25)
text(b, [("[신규] VAT 대출", {"font": F_BOLD, "size": 8.5, "color": ORANGE}),
         ("수협은행", {"font": F_MEDIUM, "size": 8}),
         ("127억원 / 3개월", {"size": 7.5})])
arrow(s, X0 + 3.85, TOP2 + 1.76, X0 + 3.65, TOP2 + 1.76, color=ORANGE, dash=True)
# 자산
b = rect(s, X0 + 1.55, TOP2 + 2.50, 2.10, 0.55, line=NAVY, lw=1.0)
text(b, [("안양물류센터", {"font": F_BOLD, "size": 9, "color": NAVY_DARK}),
         ("매입가 2,700억원 · 보증금 21억원", {"size": 7.5})])
arrow(s, X0 + 2.60, TOP2 + 2.07, X0 + 2.60, TOP2 + 2.50, color=NAVY)
label(s, X0, TOP2 + 3.15, 5.05, 0.30,
      [("총 투자비 2,989억원 = Equity 988 + 담보대출 1,980 + 보증금 21 (LTV 59.8%, 감정가 기준)", {"size": 7.5, "color": GRAY})],
      align=PP_ALIGN.LEFT)

# 우: 대주 구성
XR = 5.82
bnd = d.band(s, "대주 구성 (담보대출약정서 기준, 단위: 억원)", top=TOP2, width=5.38)
bnd.left = Inches(XR)
rows = [
    ["Tranche", "대 주", "약정금", "금리"],
    ["A-1\n(21개사)", "롯데캐피탈 300 · 신협(중앙회 포함 17개) 300\n우리카드 100 · 수협은행 100 · SBI저축은행 100", "900", "연 5.00% 고정"],
    ["A-2", "우리은행", "700", "CD+1.43% 변동"],
    ["B\n(3개사)", "키움캐피탈 200 · MG캐피탈 100 · DB저축은행 80", "380", "연 6.00% 고정"],
    ["합계", "25개사 (만기 24개월, 2026.10.29 인출)", "1,980", "-"],
    ["[신규] VAT", "수협은행 (본건)", "127", "[TBD]"],
]
tbl = d.table(s, rows, top=TOP2 + 0.40, left=XR, width=5.38, col_widths=[0.85, 2.85, 0.65, 1.03],
              row_height=0.42, size=8.5)
tbl.rows[1].height = Inches(0.62)
for r in range(1, len(rows)):
    for c in (2,):
        tbl.cell(r, c).text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
bold_cells(tbl, [(4, c) for c in range(4)])
bold_cells(tbl, [(5, c) for c in range(4)], color=ORANGE)
for c in range(4):
    tbl.cell(5, c).fill.solid()
    tbl.cell(5, c).fill.fore_color.rgb = RGBColor(0xFE, 0xF4, 0xEA)
label(s, XR, TOP2 + 3.15, 5.38, 0.30,
      [("담보: 담보신탁 우선수익권(A 공동 2순위, B 3순위), 보험·예금 근질권 — VAT 대출 담보(환급청구권·환급계좌)와 목적물 분리",
        {"size": 7.5, "color": GRAY})], align=PP_ALIGN.LEFT)
d.footnote(s, "* 출처: 담보대출약정서(BKL v38, 대주별 약정금·이자율), IM v03(투자구조도·Equity 구성), 감정평가서(미래새한, '26.09). "
              "Equity 보유자는 클로징 기준 [확인 필요].", top=7.12)

# =====================================================================
# 3p. 추진 일정 (시계열 막대)
# =====================================================================
s = d.content_slide("VAT", "3. 추진 일정 및 상환 구조",
                    lead="'26.11.25 대리납부·조기환급 신고 → '26.12.10 환급금 수협은행 예치 → '27.02.25 만기 상환",
                    subs=["조기환급이 지연돼 일반환급으로 넘어가도 환급기한('27.02.24)이 대출 만기 이내에 있음"])
d.band(s, "VAT 대출 일정 (시계열)")

D0, D1 = dt.date(2026, 10, 26), dt.date(2027, 3, 5)
LBL_X, LBL_W = 0.57, 3.05
CX0, CX1 = 3.70, 11.15
TOPC = 2.62


def xpos(day):
    return CX0 + (CX1 - CX0) * (day - D0).days / (D1 - D0).days


# 월 눈금
for m in [dt.date(2026, 11, 1), dt.date(2026, 12, 1), dt.date(2027, 1, 1), dt.date(2027, 2, 1), dt.date(2027, 3, 1)]:
    x = xpos(m)
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(TOPC + 0.28), Inches(x), Inches(TOPC + 3.72))
    ln.line.color.rgb = LINE
    ln.line.width = Pt(0.5)
    label(s, x - 0.45, TOPC, 0.9, 0.24, [(f"'{m:%y}.{m.month:02d}", {"font": F_MEDIUM, "size": 8.5, "color": NAVY_DARK})])

ROWH = 0.42
tasks = [
    # (라벨, 보조, 시작, 끝, 색, 표시)
    ("거래종결 · 대리납부 의무 발생", "매매대금 지급(세금계산서 수취)", dt.date(2026, 10, 29), None, NAVY, "10.29"),
    ("수협은행 심의 · 대주단 동의 · 약정", "기존 대주단 사전 서면동의 포함", dt.date(2026, 11, 3), dt.date(2026, 11, 20), BAND, "11.03 ~ 11.20 [TBD]"),
    ("대출 인출 · 대리납부 · 조기환급 신고", "10월분 조기환급 신고(사업설비 취득)", dt.date(2026, 11, 25), None, RED, "11.25"),
    ("VAT 대출 기간 (3개월)", "127억원, 금리 [TBD]", dt.date(2026, 11, 25), dt.date(2027, 2, 25), NAVY, "11.25 ~ '27.02.25"),
    ("조기환급 심사 (15일)", "신고기한 경과 후 15일 이내 환급", dt.date(2026, 11, 25), dt.date(2026, 12, 10), ORANGE, "~ 12.10"),
    ("환급금 수협은행 예치", "여유자금운용 예치금 (근질권 설정)", dt.date(2026, 12, 10), dt.date(2027, 2, 25), GREEN, "12.10 ~ '27.02.25"),
    ("만기 일시상환", "예치금으로 대출원리금 상환", dt.date(2027, 2, 25), None, RED, "'27.02.25"),
    ("(참고) 일반환급 전환 시", "2기 확정신고('27.01.25) 후 30일 이내", dt.date(2026, 11, 25), dt.date(2027, 2, 24), LGRAY, "환급기한 '27.02.24"),
]
for i, (name, sub, a, b_, col, tag) in enumerate(tasks):
    y = TOPC + 0.34 + i * ROWH
    if i % 2 == 0:
        bg = rect(s, LBL_X, y - 0.02, CX1 - LBL_X, ROWH - 0.02, fill=RGBColor(0xF7, 0xF7, 0xF7))
    label(s, LBL_X + 0.05, y, LBL_W, ROWH - 0.06,
          [(name, {"font": F_MEDIUM, "size": 8.5, "color": INK}), (sub, {"size": 7, "color": GRAY})],
          align=PP_ALIGN.LEFT)
    yc = y + (ROWH - 0.06) / 2
    if b_ is None:  # 마일스톤
        x = xpos(a)
        m = rect(s, x - 0.09, yc - 0.09, 0.18, 0.18, fill=col, shape=MSO_SHAPE.DIAMOND)
        label(s, x + 0.12, yc - 0.12, 1.3, 0.24, [(tag, {"font": F_BOLD, "size": 8, "color": col})], align=PP_ALIGN.LEFT)
    else:
        x1, x2 = xpos(a), xpos(b_)
        bar = rect(s, x1, yc - 0.10, x2 - x1, 0.20, fill=None if col == LGRAY else col,
                   line=GRAY if col == LGRAY else None, dash=(col == LGRAY))
        if x2 - x1 > 1.6:
            text(bar, [(tag, {"font": F_BOLD, "size": 7.5, "color": GRAY if col == LGRAY else WHITE})])
        else:
            label(s, x2 + 0.06, yc - 0.12, 1.6, 0.24, [(tag, {"font": F_BOLD, "size": 7.5, "color": col})],
                  align=PP_ALIGN.LEFT)

# 만기선
xm = xpos(dt.date(2027, 2, 25))
ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(xm), Inches(TOPC + 0.28), Inches(xm), Inches(TOPC + 3.72))
ln.line.color.rgb = RED
ln.line.width = Pt(1.0)
ln.line.dash_style = 4

# 하단 요약
d.band(s, "일정 근거 및 상환 구조", top=6.42)
label(s, 0.62, 6.72, 10.6, 0.36,
      [("대리납부: 대가 지급일이 속하는 달의 다음 달 25일까지(부가가치세법 제52조 제4항·시행령 제95조)  │  조기환급: 신고기한 경과 후 15일 이내(제59조 제2항·시행령 제107조)  │  "
        "환급금은 수협은행 부가세환급계좌로 수령 → 만기까지 예치 → 대출 상환", {"size": 8, "color": INK})],
      align=PP_ALIGN.LEFT)
d.footnote(s, "* 세무서 현장확인 등으로 환급이 지연될 수 있음. 세무 판단은 내부 검토이며 외부 회계법인·세무대리인 확인 필요. "
              "일정은 '26.11.02 기준 계획으로 변동 가능.", top=7.12)

p = d.save(OUT)
print("생성:", p, len(d.prs.slides._sldIdLst), "장")
