# -*- coding: utf-8 -*-
"""KLI 투자심의자료(안양 지분투자, 라이프인프라 관점) draft V2 → v03 보완 빌드.

사용자 지시(2026-10-01, 사용자 페이지 = 파일 슬라이드 번호 - 1):
  2p~4p(슬3~5) 계약 조건·일정 / 5p(슬6) 지분구성표 표 / 8p(슬9) 예상수익률·시총·자본총계
  14p(슬15) 표 형식 통일 / 15p(슬16) 법률검토·수권절차 / 16p(슬17) 점도표→차트, '27~'28 확장, 5년 채권금리 표시
  공통: 모든 글씨 세로 '중간' 정렬
"""
import copy, json, pathlib, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

HERE = pathlib.Path(__file__).parent
SRC = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else HERE.parent / "20261001_투심자료_KLI_안양지분투자_draftV2.pptx"
OUT = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "out.pptx"

F_B, F_M, F_L = "KoPubWorld돋움체 Bold", "KoPubWorld돋움체 Medium", "KoPubWorld돋움체 Light"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

prs = Presentation(str(SRC))
S = lambda n: prs.slides[n - 1]          # 파일 슬라이드 번호(1부터)


# ── 공통 도구 ───────────────────────────────────────────────────────────
def set_font(run, name, size=None, bold=None, color=None):
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", name)
    # a:latin/ea/cs 순서는 solidFill 뒤여야 한다 → 재정렬
    fills = rPr.findall(qn("a:solidFill"))
    for f in fills:
        rPr.remove(f); rPr.insert(0, f)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
        sf = rPr.find(qn("a:solidFill")); rPr.remove(sf); rPr.insert(0, sf)


def find(slide, sid):
    def walk(shapes):
        for sh in shapes:
            if sh.shape_id == sid:
                return sh
            if sh.shape_type == 6:
                r = walk(sh.shapes)
                if r is not None:
                    return r
    return walk(slide.shapes)


def set_text_keep(shape_or_tf, lines):
    """기존 서식을 유지한 채 문단별 텍스트 교체. lines: 문단 목록."""
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, "text_frame") else shape_or_tf
    paras = tf.paragraphs
    txBody = tf._txBody
    # 필요한 만큼 문단 복제
    while len(tf.paragraphs) < len(lines):
        txBody.append(copy.deepcopy(tf.paragraphs[-1]._p))
    while len(tf.paragraphs) > len(lines):
        txBody.remove(tf.paragraphs[-1]._p)
    for p, line in zip(tf.paragraphs, lines):
        runs = p.runs
        if not runs:
            r = p.add_run()
            end = p._p.find(qn("a:endParaRPr"))
            if end is not None:
                rPr = copy.deepcopy(end); rPr.tag = qn("a:rPr")
                old = r._r.find(qn("a:rPr"))
                if old is not None: r._r.remove(old)
                r._r.insert(0, rPr)
            runs = [r]
        runs[0].text = line
        for extra in runs[1:]:
            p._p.remove(extra._r)


def fill_cell_keep(cell, lines):
    set_text_keep(cell.text_frame, lines if isinstance(lines, list) else [lines])


def ln(tag, w=3175, color="BFBFBF"):
    el = etree.Element(qn(f"a:{tag}"), w=str(w), cap="flat", cmpd="sng", algn="ctr")
    if color is None:
        etree.SubElement(el, qn("a:noFill"))
    else:
        sf = etree.SubElement(el, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=color)
    etree.SubElement(el, qn("a:prstDash"), val="solid")
    return el


def style_cell(cell, lines, kind="body", size=10, align=None, first_col=False, last_col=False,
               top_dark=False, bold=False, color="000000"):
    """앞 장표(정관·주주간계약 표)와 같은 표 디자인.
    kind: head(회색 D9D9D9·Bold·가운데) / label(F2F2F2·Medium·가운데) / body(무채·Light)"""
    lines = lines if isinstance(lines, list) else [lines]
    tf = cell.text_frame
    txBody = tf._txBody
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    font = {"head": F_B, "label": F_M, "body": F_L}[kind]
    if bold and kind == "body":
        font = F_M
    if align is None:
        align = PP_ALIGN.LEFT if kind == "body" else PP_ALIGN.CENTER
    for line in lines:
        p = etree.SubElement(txBody, qn("a:p"))
        from pptx.text.text import _Paragraph
        para = _Paragraph(p, tf)
        para.alignment = align
        r = para.add_run(); r.text = line
        set_font(r, font, size=size, color=color)
    tcPr = cell._tc.get_or_add_tcPr()
    for ch in list(tcPr):
        tcPr.remove(ch)
    tcPr.set("anchor", "ctr")
    tcPr.set("marL", "72000" if (kind == "body" and align == PP_ALIGN.LEFT) else "9525")
    tcPr.set("marR", "9525"); tcPr.set("marT", "9525"); tcPr.set("marB", "0")
    tcPr.append(ln("lnL", 9525, None) if first_col else ln("lnL"))
    tcPr.append(ln("lnR", 3175, None) if last_col else ln("lnR"))
    tcPr.append(ln("lnT", 6350, "000000") if top_dark else ln("lnT"))
    tcPr.append(ln("lnB"))
    fill = {"head": "D9D9D9", "label": "F2F2F2"}.get(kind)
    if fill:
        sf = etree.SubElement(tcPr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=fill)
    else:
        etree.SubElement(tcPr, qn("a:noFill"))


def restyle_table(tbl, size=10, head_rows=1, label_cols=1, bold_rows=(), bold_cols=(), center_body=True):
    nr, nc = len(tbl.rows), len(tbl.columns)
    for ri in range(nr):
        for ci in range(nc):
            cell = tbl.cell(ri, ci)
            if cell.is_spanned:
                continue
            lines = [p.text for p in cell.text_frame.paragraphs] or [""]
            kind = "head" if ri < head_rows else ("label" if ci < label_cols else "body")
            style_cell(cell, lines, kind, size=size,
                       align=PP_ALIGN.CENTER if (kind != "body" or center_body) else PP_ALIGN.LEFT,
                       first_col=(ci == 0), last_col=(ci == nc - 1), top_dark=(ri == 0),
                       bold=(ri in bold_rows or ci in bold_cols))


def textbox(slide, x, y, w, h, lines, size=8, color="404040", font=F_L, align=PP_ALIGN.LEFT, bold_first=False):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        set_font(r, F_B if (bold_first and i == 0) else font, size=size, color=color)
    return tb


# ── 슬3 (사용자 2p): 심의안건 및 추진일정 ─────────────────────────────────
s = S(3)
set_text_keep(find(s, 3), [
    "본 투자심의 안건은 KLI의 자리츠 보통주 출자 및 1·2종 우선주 매입확약서(LOC) 날인 건임",
    "10/16 주주간계약 체결·매입확약서 제출 후 10/21 자리츠(KLL리츠) 보통주 주금 납입 예정",
])
tb = [sh for sh in s.shapes if sh.has_table][0].table
fill_cell_keep(tb.cell(1, 2), [
    "주주간계약 체결 및 보통주 148억·제2-1종 160억원 출자",
    "이사 2인·대표이사 지명, 의결권 지분 71.4% (납입 10/21)",
])
fill_cell_keep(tb.cell(2, 2), [
    "제1종 240억(24개월)·제2종 340억(12개월) 매입확약서 제출",
    "미이행 시 1종 연 5.0%, 2종 연 3.0% + 위약벌 20%",
])
# 타임라인 (KLI 추진일정 17p 기준으로 변경)
set_text_keep(find(s, 26), ["발기설립 완료"]); set_text_keep(find(s, 27), ["9월 11일"])
set_text_keep(find(s, 107), ["9월 23일"])
set_text_keep(find(s, 106), ["영업인가 접수", "선순위 대출심의 완료"])
set_text_keep(find(s, 99), ["10월 7일"])
set_text_keep(find(s, 36), ["당 투자심의위원회 [10/7]"])
set_text_keep(find(s, 20), ["10월 16일"])
set_text_keep(find(s, 19), ["영업인가 10/12", "매매계약 10/15", "주주간계약 체결", "대출약정서 체결"])
set_text_keep(find(s, 118), ["10월 28일"])
set_text_keep(find(s, 109), [
    "코람코라이프인프라리츠 설립자본금 3억원 출자 및 발기설립 완료(9/11)",
    "사업일정은 협의에 따라 변동될 수 있음. (KLI 추진일정 기준, 2026-10-01 갱신)",
])

# ── 슬4 (사용자 3p): 주요 조건(정관) ─────────────────────────────────────
AOI = [
    ("상호·목적", ["주식회사 코람코라이프로지스위탁관리부동산투자회사(KLL REIT)",
                "부동산 취득·관리·임대·처분 등 부투법상 자산 투자·운용 (제1·2조)"]),
    ("존립기간·상장", ["영속형 (제5조) / 상장요건 충족 시 지체 없이 상장 (제13조)"]),
    ("발행주식", ["액면가 1,000원, 발행예정주식 10억주 (제7·8조)",
              "보통주, 제1-1·1-2·2-1·2-2·3종 종류주식 (종류별 5천만주 한도, 제10·11조①)"]),
    ("의결권", ["보통주·제1-1종·제2-1종 의결권 있음 / 제1-2종·제2-2종·제3종 무의결권",
             "무의결권 종류주식 합계는 발행주식총수의 1/4 이내 (제11조①⑤)"]),
    ("배당률", ["제1종 연 7.0%·제2종 연 7.5% 누적배당(발행가액 기준), 제3종 누적배당 없음",
             "보통주 확정배당 연 7.5% (납입일·소유권취득일 중 늦은 날부터 1년) (제11조②)"]),
    ("평상시 배당순서", ["① 1종 누적미배당 → ② 1종 당기 → ③ 2종 누적미배당 → ④ 2종 당기",
                   "→ ⑤⑥ 보통주 확정배당(미배당분 포함) → ⑦ 잔여 이익 보통주 (제11조②)"]),
    ("매각 결산기 배당", ["제2항 ①~⑥ 배당 → 1종 원본 → 2종 원본 → 3종 원본 → 보통주 원본",
                    "잔여 이익 20% 제2종, 80% 보통주 (제11조③)"]),
    ("청산 시 분배", ["해산 결산기 미배당액 → 1종 원본 → 2종 원본 → 3종 원본 → 보통주 원본",
                 "잔여재산 20% 제2종, 80% 보통주 (제11조④)"]),
    ("신주발행", ["이사회 결의로 발행, 영업인가 전 제3자 배정 불가",
              "인가 후 3년 내 신주발행계획·매매대금 지급 목적 등은 제3자 배정 허용 (제14조)"]),
    ("주식 양도", ["제3자에게 양도 시 이사회 승인 필요 (제19조)"]),
    ("주주총회 결의", ["보통결의: 총자산 30% 초과 자산 취득·처분, 차입·사채발행계획, 자산관리·보관계약 등",
                  "특별결의: 정관 변경 등 (제26·27조)"]),
]
tb = [sh for sh in S(4).shapes if sh.has_table][0].table
for i, (k, v) in enumerate(AOI, start=1):
    fill_cell_keep(tb.cell(i, 0), [k]); fill_cell_keep(tb.cell(i, 1), v)
textbox(S(4), 521371, 6420000, 9756000, 330000, [
    "※ 출처: KLL리츠 정관(BKL 25359401.v8). 제3종 유상감자 미이행 시 가산배당(연 6%→8%)·매각청구 및 배당순서"
    "(1종 배당→2종 배당→1종 원본→2종 원본→3종 배당) 반영을 위한 정관 개정 예정(협의 중)",
], size=8)

# ── 슬5 (사용자 4p): 주요 조건(주주간계약서) ─────────────────────────────
SHA = [
    ("당사자", ["KLI(보통주·제2종), 이지스리츠3호·키움캐피탈·애큐온캐피탈·엠지캐피탈(제1-1종), 기계설비건설공제조합·",
             "삼성증권·이지스K리츠(제2종), 엘에프(제3종), 대상회사 (별첨1)"]),
    ("출자 구조", ["총 988억원: 보통주 148억원(@5,000원, KLI) / 제1-1종 240억원 / 제2종 500억원(KLI 160억원 포함)",
                "제3종 100억원(엘에프) — 종류주식 발행가 @25,000원 (제2.1조)"]),
    ("주금 납입", ["거래종결일 2영업일 전까지 이사회가 정한 납입예정일에 회사 지정 계좌로 납입",
               "(14:00 납입 노력, 15:00까지 완료) — 기준 일정 10/21 납입 (제2.1·2.2조)"]),
    ("발기주식 감자", ["유상증자 완료 후 KLI 발기설립 보통주 300,000주(3억원) 100% 유상감자 (제2.3조)"]),
    ("배당", ["제1종 연 7.0%·제2종 연 7.5% 누적배당, 제3종 누적배당 미적용 (제1.3조)",
            "보통주 초기 1년 연 7.5% 배당 협조 — 배당 재원·결의가 없는 경우 지급·보전 의무 없음"]),
    ("제1종 매입확약", ["KLI가 발행일부터 24개월 응당일까지 제1종 전부를 발행가액(25,000원/주)으로 매입",
                    "미이행 시 미매입잔액에 연 5.0% 지연손해금, 매입의무 존속 (제3조, 매입확약서)"]),
    ("제2종 매입확약", ["KLI가 발행일부터 12개월 응당일까지 제2종 전부(KLI 보유분 제외)를 발행가액으로 매입",
                    "미이행 시 연 3.0% 지연이자 + 미매입잔액 20% 위약벌, 매입의무 존속 (제4조, 매입확약서)"]),
    ("매도확약", ["종류주주는 체결일(셀다운 양수인은 양도 완결 전)에 KLI 앞 매도확약서 제출",
              "KLI 요청일부터 10일 내 동일 조건 매도, 미이행 시 1종 연 5.0% / 2종 연 3.0%+위약벌 20% (제3.5·4.5조)"]),
    ("제3종 유상감자", ["발행일부터 24개월(법정절차 지연 시 +[3]개월) 내 엘에프에 발행가액으로 감자",
                    "미완료 시 엘에프 감자요청 → 회사는 리파이낸싱 등으로 감자재원 조달 (제5조)"]),
    ("지배구조", ["이사 3인 중 2인 및 대표이사 KLI 지명, 감사 1인 주주총회 선임 (제6조)"]),
    ("양도 제한", ["원칙: 다른 주주 전원 서면동의 / KLI 매입거래 양도는 동의 불요 (제7.1·7.5조)",
                "셀다운: 매도확약서 제출 + KLI 사전동의(우리투자증권·삼성증권 최초 셀다운은 면제) (제7.6조)"]),
]
tb = [sh for sh in S(5).shapes if sh.has_table][0].table
for i, (k, v) in enumerate(SHA, start=1):
    fill_cell_keep(tb.cell(i, 0), [k]); fill_cell_keep(tb.cell(i, 1), v)
textbox(S(5), 521371, 6420000, 9756000, 330000, [
    "※ 출처: 주주간계약서 BKL v23 final(2026-09-30), KLI 매입확약서(제1종 투자자용 v1 / 제2종 BKL v1_4), 매도확약서(투자자용 v1)."
    " 체결 예정일 2026-10-16. 제2종 KLI 인수분 160억원은 지분구성표 기준(계약서 표는 [*])",
], size=8)

# ── 슬6 (사용자 5p): 지분 구성표 → 표 ─────────────────────────────────────
s = S(6)
pic = find(s, 11); pic._element.getparent().remove(pic._element)
CAP = [  # (종류, 주주, 의결권, 투자금 백만원, 발행가)
    ("제1-1종 종류주식", "이지스자산운용(3호펀드)", True, 10000, 25000),
    ("제1-1종 종류주식", "키움캐피탈", True, 6000, 25000),
    ("제1-1종 종류주식", "애큐온캐피탈", True, 5000, 25000),
    ("제1-1종 종류주식", "MG캐피탈", True, 3000, 25000),
    ("제2-1종 종류주식", "코람코라이프인프라리츠", True, 16000, 25000),
    ("제2-1종 종류주식", "삼성증권", True, 12000, 25000),
    ("제2-2종 종류주식\n(무의결)", "이지스자산운용(K리츠펀드)", False, 10000, 25000),
    ("제2-2종 종류주식\n(무의결)", "기계설비건설공제조합", False, 12000, 25000),
    ("제3종 종류주식\n(무의결)", "매도인(코크렙안양PFV)", False, 10000, 23000),
    ("보통주식", "코람코라이프인프라리츠", True, 14800, 5000),
]
shares = [round(m * 1_000_000 / px) for *_, m, px in CAP]
tot_sh = sum(shares); vote_sh = sum(sh for (c, n, v, m, px), sh in zip(CAP, shares) if v)
hdr = ["종류", "주주명", "의결권", "투자금(백만원)", "발행가(원)", "주식수(주)", "지분율", "의결권 주식수", "의결권 지분율"]
widths = [1350000, 2050000, 620000, 1000000, 880000, 1050000, 800000, 1156000, 850000]
nr = 1 + len(CAP) + 1
gf = s.shapes.add_table(nr, len(hdr), Emu(521371), Emu(1450179), Emu(sum(widths)), Emu(274320 * nr))
gf.table.first_row = False; gf.table.horz_banding = False
tbl_el = gf._element.graphic.graphicData.tbl
sid = tbl_el.tblPr.find(qn("a:tableStyleId"))
if sid is not None: sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
t = gf.table
for i, w in enumerate(widths): t.columns[i].width = Emu(w)
for r in t.rows: r.height = Emu(274320)
last = len(hdr) - 1
for ci, h in enumerate(hdr):
    style_cell(t.cell(0, ci), [h], "head", first_col=ci == 0, last_col=ci == last, top_dark=True)
for ri, ((cls, nm, v, m, px), sh) in enumerate(zip(CAP, shares), start=1):
    row = [cls.split("\n"), nm, "O" if v else "-", f"{m:,}", f"{px:,}", f"{sh:,}", f"{sh / tot_sh:.1%}",
           f"{sh:,}" if v else "-", f"{sh / vote_sh:.1%}" if v else "-"]
    for ci, val in enumerate(row):
        kl = "label" if ci == 0 else "body"
        style_cell(t.cell(ri, ci), val, kl, align=PP_ALIGN.LEFT if ci == 1 else PP_ALIGN.CENTER,
                   first_col=ci == 0, last_col=ci == last,
                   bold=("코람코라이프인프라리츠" in nm and ci in (1, 8)))
tr = nr - 1
tot = ["합계", "", "", f"{sum(m for *_, m, px in CAP):,}", "", f"{tot_sh:,}", "100.0%", f"{vote_sh:,}", "100.0%"]
for ci, val in enumerate(tot):
    style_cell(t.cell(tr, ci), [val], "head" if ci == 0 else "body", first_col=ci == 0, last_col=ci == last,
               bold=True, top_dark=True)
# 종류 셀 병합 (같은 종류 연속행)
ri = 1
while ri <= len(CAP):
    rj = ri
    while rj + 1 <= len(CAP) and CAP[rj][0] == CAP[ri - 1][0]:
        rj += 1
    if rj > ri:
        t.cell(ri, 0).merge(t.cell(rj, 0))
    ri = rj + 1
for ri in range(1, len(CAP) + 1):
    c = t.cell(ri, 0)
    if c.is_merge_origin or not c.is_spanned:
        style_cell(c, CAP[ri - 1][0].split("\n"), "label", first_col=True)
t.cell(tr, 0).merge(t.cell(tr, 2))
style_cell(t.cell(tr, 0), ["합계"], "head", first_col=True, top_dark=True)
kli_vote = (shares[4] + shares[9]) / vote_sh
nonvote = (shares[6] + shares[7] + shares[8]) / tot_sh
textbox(s, 521371, 1450179 + 274320 * nr + 120000, 9756000, 560000, [
    f"* 불균등발행가 / 유상증자 방식. 액면가 1,000원, 보통주 5,000원·우선주 25,000원·제3종 23,000원(상증세법 순자산가치 기준 검토가액)",
    f"* KLI 의결권 지분 {kli_vote:.1%} (보통주 {shares[9] / vote_sh:.1%} + 제2-1종 {shares[4] / vote_sh:.1%}), "
    f"무의결권 주식 비율 {nonvote:.1%} (상법 제344조의3 25% 한도 이내)",
    "* 주주간계약 v23 final 표는 기계설비건설공제조합을 제2-1종, 삼성증권을 제2-2종으로, 제3종 주주를 엘에프(25,000원·400,000주)로 기재 — 확정 시 정합 필요",
], size=8)

# ── 슬9 (사용자 8p): 투자개요 — 시총·자본총계·예상수익률 ──────────────────
s = S(9)
SHARES_KLI, PX = 97_335_354, 4_215
mcap = SHARES_KLI * PX / 1e8
t9 = [sh for sh in s.shapes if sh.has_table and sh.name == "표 2"][0].table
fill_cell_keep(t9.cell(3, 1), [f"{mcap:,.0f}억원 (97,335,354주 × @4,215원, ’26.9월말 종가)"])
fill_cell_keep(t9.cell(3, 3), ["4,720억원 (제12기 결산 기준)"])
# Sell-down 탭(기준 재무모델 v02) 결과
RET = {  # CoC(C.G제외), CoC(C.G포함), IRR(C.G포함)
    "1종 우선주": (0.06937781798839748, (13440000.000000002 - 215232.6018808782) / 24215232.60188088 * 12 / 96, 0.06881138014446986),
    "2종 우선주": (0.07432423259998445, (33750000 + 11660499.233842118) / 50454607.721046075 * 12 / 108, 0.09283432211768972),
    "보통주": (0.0075, (1110000 + 48460427.81955277) / 14800000 * 12 / 120, 0.16438586395039967),
    "KLI 수익률": (0.053984672181898334, (48300000 + 59905694.45151401) / 89469840.32292695 * 12 / 120, 0.10979473504796933),
}
t5 = [sh for sh in s.shapes if sh.has_table and sh.name == "표 5"][0].table
for ri in range(1, 5):
    key = t5.cell(ri, 0).text.strip()
    a, b, c = RET[key]
    for ci, v in zip((1, 2, 3), (a, b, c)):
        fill_cell_keep(t5.cell(ri, ci), [f"{v:.2%}"])
textbox(s, 1889525, 5700000, 8208908, 175000, [
    "※ 기준 재무모델 v02 Sell-down 탭 (보통주 ’26.10 출자, 2종 ’27.10·1종 ’28.10 매입, ’36.10 매각) / "
    "CoC(매각차익 포함) = (운영배당+매각차익배분)÷투자원금÷보유기간"], size=7, align=PP_ALIGN.RIGHT)

# ── 슬15 (사용자 14p): 표 형식 통일 ──────────────────────────────────────
s = S(15)
for sh in s.shapes:
    if not sh.has_table:
        continue
    if sh.name == "표 170":
        restyle_table(sh.table, size=10, bold_rows=(7,))
    elif sh.name == "표 13":
        restyle_table(sh.table, size=10, bold_cols=(1,))
    elif sh.name == "표 15":
        restyle_table(sh.table, size=10)

# ── 슬16 (사용자 15p): 법률검토·수권절차 ─────────────────────────────────
LEGAL = HERE / "15p_법률검토_수권절차.json"
if LEGAL.exists():
    L = json.loads(LEGAL.read_text(encoding="utf-8"))
    s = S(16)
    cover = find(s, 14); cover._element.getparent().remove(cover._element)
    set_text_keep(find(s, 70), L["headline"])
    for sh in s.shapes:
        if not sh.has_table:
            continue
        rows = {"표 6": L["securities"], "표 1": L["bonds"], "표 66": L["issues"]}.get(sh.name)
        if rows is None:
            continue
        t = sh.table
        for ri, row in enumerate(rows, start=1):
            for ci, val in enumerate(row):
                fill_cell_keep(t.cell(ri, ci), [val])
        restyle_table(t, size=9 if sh.name == "표 66" else 9, center_body=(sh.name != "표 66"))
    set_text_keep(find(s, 8), L["footnote"])

# ── 슬17 (사용자 16p): 점도표 → 차트 ─────────────────────────────────────
s = S(17)
D = json.loads((HERE / "차트판독값_원장표점도표.json").read_text(encoding="utf-8"))
grp = find(s, 8)
lbl_xml = copy.deepcopy(find(s, 69)._element)      # ’25~’26.1Q Cap.Rate Avg. 5.42%
box_xml = copy.deepcopy(find(s, 67)._element)      # 점선 박스
grp._element.getparent().remove(grp._element)
# 원 차트 판독값 보정(겹친 점에 가려진 구간)
bond = D["bond"][:]; office = D["office"][:]
bond[27] = 3.92
for i, v in {15: 4.6, 19: 4.2, 22: 4.0, 23: 4.0, 26: 3.9, 39: 4.69}.items():
    office[i] = v
office[40] = None
N = 52                                         # 2016Q1 ~ 2028Q4
pad = lambda a: a + [None] * (N - len(a))
bond = pad(bond); office = pad(office)
bond[42] = 4.133                               # 2026Q3: 재무모델 채권금리 탭 국고채 5년(’26.9.9)
dots = [pad([(d[k] if k < len(d) else None) for d in D["logi"]]) for k in range(4)]
pts = [(i, v) for i, d in enumerate(D["logi"]) for v in d]
n = len(pts); mx = sum(i for i, _ in pts) / n; my = sum(v for _, v in pts) / n
b1 = sum((i - mx) * (v - my) for i, v in pts) / sum((i - mx) ** 2 for i, _ in pts); b0 = my - b1 * mx
trend = [b0 + b1 * i for i in range(N)]

cd = CategoryChartData(number_format="0.0%")
for y in range(2016, 2029):
    yc = cd.categories.add_category(str(y))
    for q in range(1, 5):
        yc.add_sub_category(f"Q{q}")
pct = lambda a: [None if v is None else round(v / 100, 5) for v in a]
cd.add_series("5년 채권금리", pct(bond))
cd.add_series("한국 오피스 Cap.Rate", pct(office))
for k in range(4):
    cd.add_series("물류 Cap.Rate" if k == 0 else f"물류 Cap.Rate{k + 1}", pct(dots[k]))
cd.add_series("물류 Cap.Rate 추세(’27~’28 연장)", pct(trend))
CX, CY, CW, CH = 4535926, 2691299, 5778531, 1673926
gfc = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Emu(CX), Emu(CY), Emu(CW), Emu(CH), cd)
ch = gfc.chart
ch.font.size = Pt(7); ch.font.name = F_L
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(7)
NAVY, TEAL, PLUM = "25273A", "6499A2", "6B2E50"
for idx, ser in enumerate(ch.plots[0].series):
    ser.smooth = False
    if idx in (0, 1, 6):
        ser.marker.style = XL_MARKER_STYLE.NONE
        ser.format.line.color.rgb = RGBColor.from_string((NAVY, TEAL, None, None, None, None, PLUM)[idx])
        ser.format.line.width = Pt(1.5 if idx == 0 else 1.25 if idx == 1 else 1.25)
        if idx == 6:
            ser.format.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    else:
        ser.format.line.fill.background()
        ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 4
        ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = RGBColor.from_string(PLUM)
        ser.marker.format.line.fill.background()
# 5년 채권금리 데이터 레이블: 매년 Q4 + 최근값(’26.1Q, ’26.3Q)
bser = ch.plots[0].series[0]
for i in [3 + 4 * k for k in range(10)] + [40, 42]:
    if bond[i] is None:
        continue
    dl = bser.points[i].data_label
    dl.has_text_frame = True
    dl.text_frame.text = f"{bond[i]:.1f}%"
    dl.position = XL_LABEL_POSITION.BELOW
    r = dl.text_frame.paragraphs[0].runs[0]
    set_font(r, F_M, size=6.5, color=NAVY)
va = ch.value_axis
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 0.09, 0.01
va.has_major_gridlines = False
va.tick_labels.number_format = "0%"; va.tick_labels.number_format_is_linked = False
va.tick_labels.font.size = Pt(7)
va.format.line.fill.background()
ca = ch.category_axis
ca.tick_labels.font.size = Pt(5.5)
_tx = ca._element.get_or_add_txPr(); _tx.find(qn("a:bodyPr")).set("rot", "0"); _tx.find(qn("a:bodyPr")).set("vert", "horz")
ca.format.line.color.rgb = RGBColor.from_string("BFBFBF")
cs = ch._chartSpace
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
# 겹치는 물류 점 계열 2~4는 범례에서 숨김
leg = cs.find(f".//{C}legend")
for i in (3, 4, 5):
    le = etree.SubElement(leg, f"{C}legendEntry")
    etree.SubElement(le, f"{C}idx", val=str(i)); etree.SubElement(le, f"{C}delete", val="1")
    leg.remove(le); leg.insert(1, le)
dba = cs.find(f".//{C}dispBlanksAs")
if dba is None:
    dba = etree.SubElement(cs.find(f"{C}chart"), f"{C}dispBlanksAs")
dba.set("val", "span")
# 플롯영역 수동 배치 → 점선 박스 위치 계산용
PX0, PY0, PW, PH = 0.055, 0.13, 0.935, 0.63
pa = cs.find(f".//{C}plotArea")
lay = pa.find(f"{C}layout")
if lay is None:
    lay = etree.Element(f"{C}layout"); pa.insert(0, lay)
for ch_ in list(lay): lay.remove(ch_)
ml = etree.SubElement(lay, f"{C}manualLayout")
for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                 ("x", PX0), ("y", PY0), ("w", PW), ("h", PH)):
    etree.SubElement(ml, f"{C}{tag}", val=str(val))
# 점선 박스·평균 라벨 재배치 (’25.1Q~’26.1Q, 4.9~6.0%)
qx = lambda i: CX + CW * (PX0 + PW * i / N)
vy = lambda v: CY + CH * (PY0 + PH * (1 - v / 9.0))
bx0, bx1 = qx(36) + 4000, qx(41) - 4000
by0, by1 = vy(6.05), vy(4.9)
spTree = s.shapes._spTree
for _r in lbl_xml.iter(qn("a:rPr")):
    _r.set("sz", "800")
for _p in lbl_xml.iter(qn("a:pPr")):
    _p.set("algn", "ctr")
for el, (x, y, w, h) in ((box_xml, (bx0, by0, bx1 - bx0, by1 - by0)),
                         (lbl_xml, (bx1 + 20000, vy(7.7), CX + CW - bx1 - 20000, 300000))):
    off = el.find(f".//{{{A}}}off"); ext = el.find(f".//{{{A}}}ext")
    off.set("x", str(int(x))); off.set("y", str(int(y)))
    ext.set("cx", str(int(w))); ext.set("cy", str(int(h)))
    spTree.append(el)
# 차트 테두리(원 그룹의 사각 틀 대체)
gfc_el = gfc._element; spTree.remove(gfc_el); spTree.append(gfc_el)
spTree.remove(box_xml); spTree.append(box_xml); spTree.remove(lbl_xml); spTree.append(lbl_xml)
textbox(s, 283367, 6813808, 10009067, 230000, [
    "* 5년 채권금리·한국 오피스 및 물류 Cap.Rate: 원 장표 점도표 판독값(분기, ±0.05%p) / ’26.3Q 채권금리는 기준 재무모델 채권금리 탭 국고채 5년 4.133%(’26.9.9) / "
    f"물류 추세선은 판독 점 {n}개 선형회귀를 ’28.4Q까지 연장({trend[-1]:.2f}%) [추정]",
], size=7)

# ── 공통: 모든 글씨 세로 '중간' 정렬 ─────────────────────────────────────
for sl in prs.slides:
    for el in sl.shapes._spTree.iter(qn("a:bodyPr")):
        el.set("anchor", "ctr")
    for el in sl.shapes._spTree.iter(qn("a:tcPr")):
        el.set("anchor", "ctr")

prs.save(str(OUT))
print("saved", OUT, "trend", round(b0, 3), round(b1, 4), "mcap", round(mcap, 1))
