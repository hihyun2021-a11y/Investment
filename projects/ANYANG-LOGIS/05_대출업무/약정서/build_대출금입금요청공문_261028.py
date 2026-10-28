# -*- coding: utf-8 -*-
"""대출금 입금(인출) 요청 공문 — 차주(코람코라이프로지스) 명의 예시.
양식 참고: 대주가 보내준 신한은행 금융결제부 「DDI REITs 신촌2 담보대출 인출요청」(2025.07.21) 공문 서식
근거: 대출약정서 최종 체결본(2026-10-27) 별지 대주별 대출약정금, 날인된 인출요청서 25부(2026-10-28, 인출일 10-29)
출력: 05_대출업무/약정서/20261028_대출약정_대출금입금요청공문_KLL명의_v01.docx
"""
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "20261028_대출약정_대출금입금요청공문_KLL명의_v01.docx")
FONT = "맑은 고딕"
BORROWER = "주식회사 코람코라이프로지스위탁관리부동산투자회사"

ROWS = [  # (구분, 표 표기 대주명, 수신처 표기(별지 명칭), 약정액(백만원))
    ("Tr.A-1", "롯데캐피탈", "롯데캐피탈(주)", 30000), ("Tr.A-1", "우리카드", "(주)우리카드", 10000),
    ("Tr.A-1", "신용협동조합중앙회", "신용협동조합중앙회", 100), ("Tr.A-1", "서로신용협동조합", "서로신용협동조합", 4000),
    ("Tr.A-1", "수지신용협동조합", "수지신용협동조합", 2800), ("Tr.A-1", "정선신용협동조합", "정선신용협동조합", 2400),
    ("Tr.A-1", "연희신용협동조합", "연희신용협동조합", 1400), ("Tr.A-1", "금모래신용협동조합", "금모래신용협동조합", 1800),
    ("Tr.A-1", "남인천신용협동조합", "남인천신용협동조합", 1000), ("Tr.A-1", "미호신용협동조합", "미호신용협동조합", 1300),
    ("Tr.A-1", "대구대서신용협동조합", "대구대서신용협동조합", 1300), ("Tr.A-1", "한빛신용협동조합", "한빛신용협동조합", 1700),
    ("Tr.A-1", "양동행복 신용협동조합", "양동행복 신용협동조합", 1900), ("Tr.A-1", "서귀포동부신용협동조합", "서귀포동부신용협동조합", 1200),
    ("Tr.A-1", "광주어룡신용협동조합", "광주어룡신용협동조합", 2700), ("Tr.A-1", "광주미래신용협동조합", "광주미래신용협동조합", 1400),
    ("Tr.A-1", "새안양 신용협동조합", "새안양 신용협동조합", 4000), ("Tr.A-1", "남수원 신용협동조합", "남수원 신용협동조합", 300),
    ("Tr.A-1", "북인천신용협동조합", "북인천신용협동조합", 700), ("Tr.A-1", "수협은행", "수협은행", 10000),
    ("Tr.A-1", "에스비아이저축은행", "(주)에스비아이저축은행", 10000),
    ("Tr.A-2", "우리은행", "(주)우리은행", 70000),
    ("Tr.B", "키움캐피탈", "키움캐피탈(주)", 20000), ("Tr.B", "엠지캐피탈", "엠지캐피탈(주)", 10000),
    ("Tr.B", "디비저축은행", "(주)디비저축은행", 8000),
]
assert sum(r[3] for r in ROWS) == 198000


def set_font(run, size=10, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    if color:
        run.font.color.rgb = RGBColor(*color)


def para(doc_or_cell, text="", size=10, bold=False, align=None, before=0, after=0, indent=None, color=None, highlight=False):
    p = doc_or_cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), 1.15
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Cm(indent)
    if text:
        r = p.add_run(text)
        set_font(r, size, bold, color)
    return p


def bottom_border(p, sz=12, color="000000"):
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", str(sz)), ("w:space", "1"), ("w:color", color)):
        b.set(qn(k), v)
    bdr.append(b)
    ppr.append(bdr)


def shade(cell, hexcolor):
    tcpr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexcolor)
    tcpr.append(sh)


def cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    set_font(r, size, bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.top_margin, sec.bottom_margin = Cm(1.4), Cm(1.2)
sec.left_margin = sec.right_margin = Cm(2.0)
doc.styles["Normal"].font.name = FONT
doc.styles["Normal"]._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)

# 머리: 발신 회사
para(doc, "코람코라이프로지스위탁관리부동산투자회사", 17, True, WD_ALIGN_PARAGRAPH.CENTER, after=2, color=(0x1F, 0x38, 0x64))
p = para(doc, "06158 서울시 강남구 삼성로 511, 4층    전화: 02-787-0049    문의: 안용현 차장 / 염진경 과장(02-787-0139)",
         8.5, False, WD_ALIGN_PARAGRAPH.CENTER, after=4)
bottom_border(p, 18, "808080")

head = [("문서번호", "KLL-2026-[    ]"), ("발신일자", "2026. 10. 28."), ("수    신", "수신처 참조"),
        ("참    조", "업무 담당자"), ("제    목", "코람코라이프로지스리츠 담보대출 인출(대출금 입금) 요청")]
for i, (k, v) in enumerate(head):
    p = para(doc, before=4 if i == 0 else 0)
    r = p.add_run(f"{k} :  ")
    set_font(r, 10)
    r = p.add_run(v)
    set_font(r, 10, bold=(k.startswith("제")))
    if k.startswith("제"):
        bottom_border(p, 12)

para(doc, "1.  귀사의 무궁한 발전을 기원합니다.", 10, before=6, indent=0.6)
para(doc, ("2.  2026년 10월 27일 체결한 『코람코라이프로지스리츠 담보대출』 대출약정서(차주: " + BORROWER +
           ")와 관련됩니다. 당사는 차주로서 대출약정서 제4조 제1항에 따라 2026년 10월 28일 인출요청서를 제출하였으며, "
           "귀사가 약정하신 대출금을 아래와 같이 입금해 주실 것을 요청합니다."), 10, before=2, after=8, indent=0.6)

for lab, val in [("입금일시", "2026년 10월 29일 목요일"),
                 ("입금계좌", "우리은행 1005-504-949583"),
                 ("", "(예금주: " + BORROWER + ")")]:
    p = para(doc, indent=1.6, after=1)
    if lab:
        r = p.add_run(f"■  {lab} :  ")
        set_font(r, 10)
    else:
        r = p.add_run("　　　　　　　 ")
        set_font(r, 10)
    r = p.add_run(val)
    set_font(r, 10, bold=True)
p = para(doc, indent=1.6, before=4, after=0)
r = p.add_run("■  대주별 입금금액")
set_font(r, 10)
para(doc, "(단위 : 백만원)", 8.5, align=WD_ALIGN_PARAGRAPH.RIGHT, after=1)

# 표
hdr = ["구분", "대    주", "약정액", "인출 전 잔액", "금번 인출액", "인출 후 미인출한도"]
groups = [("Tr.A-1", "Tr.A-1 소계"), ("Tr.A-2", "Tr.A-2 소계"), ("Tr.B", "Tr.B 소계")]
nrows = 1 + len(ROWS) + len(groups) + 1
tbl = doc.add_table(rows=nrows, cols=6)
tbl.style = "Table Grid"
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [Cm(1.8), Cm(4.4), Cm(2.4), Cm(2.4), Cm(2.4), Cm(3.4)]
for j, h in enumerate(hdr):
    c = tbl.cell(0, j)
    cell_text(c, h, bold=True)
    shade(c, "BFBFBF")
r = 1
for g, sub in groups:
    items = [x for x in ROWS if x[0] == g]
    start = r
    for _, name, _, amt in items:
        cell_text(tbl.cell(r, 1), name)
        cell_text(tbl.cell(r, 2), f"{amt:,}", align=WD_ALIGN_PARAGRAPH.RIGHT)
        cell_text(tbl.cell(r, 3), "0", align=WD_ALIGN_PARAGRAPH.RIGHT)
        cell_text(tbl.cell(r, 4), f"{amt:,}", align=WD_ALIGN_PARAGRAPH.RIGHT)
        cell_text(tbl.cell(r, 5), "0", align=WD_ALIGN_PARAGRAPH.RIGHT)
        r += 1
    m = tbl.cell(start, 0).merge(tbl.cell(r - 1, 0))
    cell_text(m, g)
    tot = sum(x[3] for x in items)
    s = tbl.cell(r, 0).merge(tbl.cell(r, 1))
    cell_text(s, "소    계", bold=True)
    for j, v in ((2, f"{tot:,}"), (3, "-"), (4, f"{tot:,}"), (5, "0")):
        cell_text(tbl.cell(r, j), v, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT if v != "-" else WD_ALIGN_PARAGRAPH.CENTER)
    for j in range(6):
        shade(tbl.cell(r, j), "E7E6E6" if g != "Tr.A-2" else "DDEBF7")
    r += 1
s = tbl.cell(r, 0).merge(tbl.cell(r, 1))
cell_text(s, "합    계", bold=True)
for j, v in ((2, "198,000"), (3, "-"), (4, "198,000"), (5, "0")):
    cell_text(tbl.cell(r, j), v, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT if v != "-" else WD_ALIGN_PARAGRAPH.CENTER)
for j in range(6):
    shade(tbl.cell(r, j), "E2EFDA")
# 열 너비 고정(LibreOffice·Word 공통): 표 레이아웃 fixed + tblGrid + 셀 너비
tblpr = tbl._tbl.tblPr
lay = OxmlElement("w:tblLayout")
lay.set(qn("w:type"), "fixed")
tblpr.append(lay)
grid = tbl._tbl.tblGrid
for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
    gc.set(qn("w:w"), str(int(widths[j].twips)))
for row in tbl.rows:
    row.height = Cm(0.42)
    for tc in row._tr.findall(qn("w:tc")):
        span = tc.tcPr.find(qn("w:gridSpan")) if tc.tcPr is not None else None
        n = int(span.get(qn("w:val"))) if span is not None else 1
        idx = list(row._tr.findall(qn("w:tc"))).index(tc)

para(doc, "붙임 : 차주 인출요청서 25부(대주별 각 1부).  끝.", 10, before=6, indent=0.6)
para(doc, BORROWER, 13, True, WD_ALIGN_PARAGRAPH.CENTER, before=10, after=0)
para(doc, "대표이사   원  재  이   (인)", 11, False, WD_ALIGN_PARAGRAPH.CENTER, before=2)

p = para(doc, before=4)
bottom_border(p, 12)
recv = ", ".join(x[2] for x in ROWS if x[0] == "Tr.A-2") + ", " + ", ".join(x[2] for x in ROWS if x[0] != "Tr.A-2")
p = para(doc, after=0)
r = p.add_run("수 신 처 :  ")
set_font(r, 8)
r = p.add_run(recv)
set_font(r, 8)

doc.save(OUT)
print("saved", OUT)
