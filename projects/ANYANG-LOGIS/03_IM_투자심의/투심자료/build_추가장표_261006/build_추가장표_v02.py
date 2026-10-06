# -*- coding: utf-8 -*-
"""ANYANG-LOGIS 추가 장표 v02 (운용역 제공 양식 261006): ① 투자구조 및 당사 손익(모리츠·자리츠1(72호)·자리츠2(본건)) ② 제반 실사 업체 선정 현황.

기준: 투심자료 v12(2026-10-02) — 당사 예상수익·재원조달·지분구성·대출조건, 기준 재무모델 v03(260930) A&R·Sell-down 탭,
대출약정서 v16 중순위 의견본(2026-10-06, 중순위 대주 구성). 양식 서식(나눔스퀘어·색상·배치)은 그대로 두고 내용만 교체.
사용법: python3 build_form.py <양식.pptx> <출력.pptx>
"""
import copy, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
IN = lambda v: Emu(int(round(v * 914400)))
S1, S2 = prs.slides[0], prs.slides[1]
NB, NBB, NXB = "나눔스퀘어", "나눔스퀘어 Bold", "나눔스퀘어 ExtraBold"
WHITE, INK, RED = "FFFFFF", "262626", "C00000"
HEAD, HEAD2, LAB, TOT, SEL, GRID = "A39382", "6F6255", "F5F3F1", "EDE6DF", "FBEEEE", "BFBFBF"


def find(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)


def drop(slide, *sids):
    for sid in sids:
        el = find(slide, sid)._element
        el.getparent().remove(el)


def fill(shape_or_tf, paras):
    """paras: [[(text, k), ...], ...] — k는 원본 문단의 k번째 run 서식을 복제. 문단 서식은 원본 i번째(부족하면 마지막) 복제."""
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, "text_frame") else shape_or_tf
    body = tf._txBody
    olds = list(body.findall(qn("a:p")))
    tmpl = [(copy.deepcopy(p.find(qn("a:pPr"))), [copy.deepcopy(r.find(qn("a:rPr"))) for r in p.findall(qn("a:r"))],
             copy.deepcopy(p.find(qn("a:endParaRPr")))) for p in olds]
    for p in olds:
        body.remove(p)
    for i, runs in enumerate(paras):
        ppr, rprs, end = tmpl[min(i, len(tmpl) - 1)]
        if not rprs:                                 # run이 없는 문단은 endParaRPr를 run 서식으로 사용
            rprs = [copy.deepcopy(end) if end is not None else etree.Element(qn("a:rPr"))]
            for r in rprs:
                r.tag = qn("a:rPr")
        p = etree.SubElement(body, qn("a:p"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        for text, k in runs:
            r = etree.SubElement(p, qn("a:r"))
            r.append(copy.deepcopy(rprs[min(k, len(rprs) - 1)]))
            t = etree.SubElement(r, qn("a:t")); t.text = text


def simple(shape, *lines):
    fill(shape, [[(ln, 0)] for ln in lines])


def cell_set(cell, *lines, k=0):
    fill(cell.text_frame, [[(ln, k)] for ln in lines])


def add_row(table, src_idx=-1):
    tr = table._tbl.tr_lst[src_idx]
    new = copy.deepcopy(tr)
    tr.addnext(new)
    return new


def del_rows(table, keep):
    for i, tr in reversed(list(enumerate(table._tbl.tr_lst))):
        if i >= keep:
            table._tbl.remove(tr)


def rpr(font, size, color, bold_none=True):
    r = etree.Element(qn("a:rPr"), lang="ko-KR", sz=str(int(size * 100)), dirty="0")
    sf = etree.SubElement(r, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=color)
    for tag in ("latin", "ea", "cs"):
        etree.SubElement(r, qn(f"a:{tag}"), typeface=font)
    return r


def style_cell(cell, text, font=NB, size=8, color=INK, fill_=None, align="r", bot=(6350, GRID)):
    tc = cell._tc
    body = cell.text_frame._txBody
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for ln in (text.split("\n") if text else [""]):
        p = etree.SubElement(body, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"), marL="0", indent="0", algn=align); etree.SubElement(ppr, qn("a:buNone"))
        if ln:
            r = etree.SubElement(p, qn("a:r")); r.append(rpr(font, size, color))
            etree.SubElement(r, qn("a:t")).text = ln
        else:
            e = rpr(font, size, color); e.tag = qn("a:endParaRPr"); p.append(e)
    pr = tc.get_or_add_tcPr()
    for ch in list(pr):
        pr.remove(ch)
    for a, v in (("marL", "73152"), ("marR", "73152"), ("marT", "18288"), ("marB", "18288"), ("anchor", "ctr")):
        pr.set(a, v)
    for tag, (w, c) in (("lnL", (6350, GRID)), ("lnR", (6350, GRID)), ("lnT", (6350, GRID)), ("lnB", bot)):
        ln = etree.SubElement(pr, qn(f"a:{tag}"), w=str(w), cap="flat", cmpd="sng", algn="ctr")
        if c:
            sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=c)
        else:
            etree.SubElement(ln, qn("a:noFill"))
    if fill_:
        sf = etree.SubElement(pr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=fill_)
    else:
        etree.SubElement(pr, qn("a:noFill"))


# ══════════════════════════════════════════════════════════════════════
# 장표 ① 투자구조 및 당사 손익 — 모리츠(KLI) · 자리츠1(코크렙제72호, 현대차) · 자리츠2(본건)
#   자리츠1 수치: 현대자동차 실물자산 유동화 프로젝트 투심자료 v1.0('26.03.26) p.3·4·15·31
#   자리츠2 수치: 투심자료 v12 · 기준 재무모델 v04(=v03 계산) / 모리츠 본건 매입보수는 투자금액×0.5% 산식 [추정]
# ══════════════════════════════════════════════════════════════════════
s = S1
simple(find(s, 2), "투자구조 및 당사 손익")
fill(find(s, 8), [[("코람코라이프인프라리츠", 0)], [("(모리츠 · 상장리츠, 포트폴리오 약 1.68조)", 0)]])
# 기존 단일 자리츠용 보수 띠·자2~4·대출표 등 삭제
drop(s, 13, 18, 19, 23, 24, 25, 26, 33, 35, 37, 38, 39, *range(51, 65))


def clone(sid):
    el = copy.deepcopy(find(s, sid)._element)
    s.shapes._spTree.append(el)
    return [sh for sh in s.shapes if sh._element is el][0]


def place(sh, x, y, w=None, h=None):
    sh.left, sh.top = IN(x), IN(y)
    if w is not None:
        sh.width = IN(w)
    if h is not None:
        sh.height = IN(h)


def new_table(rows, x, y, widths, heights, fmt):
    gf = s.shapes.add_table(len(rows), len(widths), IN(x), IN(y), IN(sum(widths)), IN(sum(heights)))
    tb = gf.table
    tp = gf._element.graphic.graphicData.tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tp.set(a, "0")
    for i, w in enumerate(widths):
        tb.columns[i].width = IN(w)
    for i, h in enumerate(heights):
        tb.rows[i].height = IN(h)
    for ri, row in enumerate(rows):
        for ci, v in enumerate(row):
            style_cell(tb.cell(ri, ci), v, *fmt(ri, ci, v))
    return gf, tb


# ── 보수 구조 표 (자리츠1 · 자리츠2 · 모리츠) ─────────────────────────────
fill(find(s, 65), [[("· ", 0), ("보수 구조 (자리츠1 · 자리츠2 · 모리츠)", 1)]])
place(find(s, 65), 3.75, 1.02, 3.6)
fee_rows = [["구 분", "Equity IRR", "매입보수", "운용보수(연)", "매각기본보수", "매각성과보수"],
            ["자리츠1 (72호)", "8.89%", "매입가 0.5%", "AUM 0.223%", "매각가 0.5%*", "매각차익 15%"],
            ["자리츠2 (본건)", "10.21%", "매입가 0.5%", "AUM 0.223%", "매각가 0.5%", "매각이익 15%"],
            ["모리츠 (KLI)", "-", "투자금 0.5%", "출자분 미청구", "-", "-"]]
def fee_fmt(ri, ci, v):
    if ri == 0:
        return (NBB, 8, WHITE, HEAD, "ctr")
    if ci == 0:
        return (NBB, 8, INK, LAB, "ctr")
    return (NXB if ci == 1 else NB, 8, RED if ci in (1, 5) and ri < 3 else INK, None, "ctr")
new_table(fee_rows, 3.75, 1.24, [1.0, 0.78, 0.92, 1.0, 0.98, 1.03], [0.25, 0.23, 0.23, 0.23], fee_fmt)

# ── 연결선 · 출자 라벨 ───────────────────────────────────────────────────
X1, X2 = 1.45, 7.15                                   # 자리츠1·2 연결 x
place(find(s, 16), X1, 2.38, X2 - X1)                  # 가로선
place(find(s, 17), X1, 2.38, None, 0.24)
place(find(s, 20), X2, 2.38, None, 0.24)
simple(find(s, 15), "KLI 출자"); place(find(s, 15), 1.55, 2.06, 1.5)
lb1 = clone(15); simple(lb1, "1종 1,193억 (50.13%)"); place(lb1, X1 + 0.08, 2.38, 2.2, 0.24)
lb2 = clone(15); simple(lb2, "보통주·2-1종 308억 → 확약 후 888억"); place(lb2, X2 + 0.08, 2.38, 2.25, 0.24)
for lb, sz in ((lb1, 9), (lb2, 8.5)):
    lb.text_frame.word_wrap = False
    for r_ in lb.text_frame.paragraphs[0].runs:
        r_.font.size = Pt(sz)

# ── 자리츠 박스 2개 (자리츠1 = 72호, 자리츠2 = 본건) ─────────────────────
BY, BH = 2.62, 2.26
box1 = find(s, 21); place(box1, 0.4, BY, 4.25, BH)
box2 = clone(21); place(box2, 4.85, BY, 4.61, BH)
h1 = find(s, 22); place(h1, 0.55, BY + 0.08, 3.1, 0.42)
fill(h1, [[("자리츠1 ", 0), ("코크렙제72호 (5,760억)", 1)], [("현대차 실물자산 11개 · 거래종결 ’26.5", 0)]])
h2 = clone(22); place(h2, 5.0, BY + 0.08, 3.6, 0.42)
fill(h2, [[("자리츠2 ", 0), ("코람코라이프로지스 (2,989억)", 1)], [("본건 안양물류센터 · 거래종결 ’26.10", 0)]])

def bars(x0, uses, srcs, top=BY + 0.62, total=1.52, w=1.0):
    """uses/srcs: [(라벨, 값, 원본 shape id)] — 높이는 금액 비례(최소 0.26)."""
    for col, items in ((0, uses), (1, srcs)):
        tot = sum(v for _, v, _ in items)
        hs = [max(0.26, total * v / tot) for _, v, _ in items]
        k = total / sum(hs); y = top
        for (lab, v, sid), h in zip(items, hs):
            sh = clone(sid)
            place(sh, x0 + col * w, y, w, h * k)
            vtxt = f"{v:,.1f}".rstrip("0").rstrip(".") if v % 1 else f"{v:,.0f}"
            if sid in (27, 30, 31):
                fill(sh, [[(lab, 0)], [(vtxt, 0)]])
            else:
                fill(sh, [[(lab, 0), (" ", 1), (vtxt, 2)]])
            y += h * k

bars(0.55, [("매매대금", 5230, 27), ("부대비용 등", 529.9, 28)],
     [("Equity", 2379.9, 30), ("Loan", 3140, 31), ("보증금", 240, 32)])
bars(5.0, [("매매대금", 2700, 27), ("부대비용", 189, 28), ("예비비", 100, 29)],
     [("Equity", 988, 30), ("Loan", 1980, 31), ("보증금", 21, 32)])
drop(s, 27, 28, 29, 30, 31, 32)

t1 = find(s, 34)
fill(t1, [[("KLI(1종) ", 0), ("1,193억 ", 1), ("(50.13%)", 2)], [("한투(2종) ", 0), ("475.2억 ", 1), ("(19.97%)", 2)],
          [("현대차(보통주) ", 0), ("711.7억 ", 1), ("(29.90%)", 2)]])
place(t1, 2.65, BY + 0.62, 1.95, 0.62)
l1 = find(s, 36)
fill(l1, [[("LTV 60% (매입가 기준)", 0)], [("선순위 담보대출 3,140억", 0)], [("한국투자증권 조건부 총액인수", 0)]])
place(l1, 2.65, BY + 1.32, 1.95, 0.70)
t2 = clone(34)
fill(t2, [[("KLI ", 0), ("308억 ", 1), ("(31.2%)", 2)], [("FI(1·2종) ", 0), ("580억 ", 1), ("(58.7%)", 2)],
          [("매도인(3종) ", 0), ("100억 ", 1), ("(10.1%)", 2)]])
place(t2, 7.12, BY + 0.62, 2.3, 0.62)
l2 = clone(36)
fill(l2, [[("LTV 59.8% (감정가 3,345억)", 0)], [("선순위 1,600 · 중순위 380", 0)], [("만기 24개월 (리파이 가정)", 0)]])
place(l2, 7.12, BY + 1.32, 2.3, 0.70)
for b_ in (t1, l1, t2, l2):
    for p_ in b_.text_frame.paragraphs:
        for r_ in p_.runs:
            r_.font.size = Pt(8.5)

# ── 당사 손익 (연도별 보수 스케줄) ────────────────────────────────────────
fill(find(s, 41), [[("당사 손익 ", 0), ("(자리츠1 · 자리츠2 · 모리츠)", 4)]])
place(find(s, 41), 0.6, 4.94); place(find(s, 40), 0.4, 5.00); place(find(s, 42), 7.9, 5.00)
old = find(s, 43); x0, w0 = old.left, old.width
old._element.getparent().remove(old._element)
yrs = ["’26"] + [f"’{y_}" for y_ in range(27, 37)]
N = 11
j1 = {"매입": [26.2] + [None] * 10, "운용": [None] + [11.81] * 10, "매각": [None] * 10 + [34.6 + 175.9]}
j2 = {"매입": [13.5] + [None] * 10, "운용": [None] + [6.66547] * 10, "매각": [None] * 10 + [18.793038 + 124.803226]}
mo = {"매입": [5.9 + 1.54, 1.70, 1.20] + [None] * 8, "운용": [None] * N}
tot_y = [sum((d[k][i] or 0) for d in (j1, j2, mo) for k in d) for i in range(N)]
f = lambda v: "" if v is None else f"{v:,.1f}"
sm = lambda a: sum(v or 0 for v in a)
R = [["구 분", ""] + yrs + ["계"]]
for veh, d in (("자리츠1\n72호", j1), ("자리츠2\n본건", j2), ("모리츠\nKLI", mo)):
    for i, (k, arr) in enumerate(d.items()):
        R.append([veh if i == 0 else "", k] + ([f(v) for v in arr] if k != "운용" or veh[:3] != "모리츠" else ["-"] * N)
                 + [f(sm(arr)) if sm(arr) else "-"])
R.append(["합 계", ""] + [f(v) for v in tot_y] + [f(sum(tot_y))])
assert abs(sm(j1["매입"]) + sm(j1["운용"]) + sm(j1["매각"]) - 354.8) < 0.11
assert abs(sum(tot_y) - (354.8 + 223.75 + 10.34)) < 0.2, sum(tot_y)
nc = len(R[0])
ws_ = [0.62, 0.40] + [0.60] * 11 + [0.66]
kk = (w0 / 914400) / sum(ws_)
hs = [0.27] + [0.175] * 8 + [0.24]
def pl_fmt(ri, ci, v):
    if ri == 0:
        return (NBB, 8.5, WHITE, HEAD2 if ci == nc - 1 else HEAD, "ctr")
    if ri == len(R) - 1:
        return (NXB, 8, INK, TOT, "ctr" if ci < 2 else "r")
    if ci < 2:
        return (NBB if ci == 0 else NB, 7.5 if ci else 7, INK, LAB, "ctr")
    return (NBB if ci == nc - 1 else NB, 7.5, INK, None, "r")
gf, tb = new_table(R, x0 / 914400, 5.22, [v * kk for v in ws_], hs, pl_fmt)
for r1, r2 in ((1, 3), (4, 6), (7, 8)):
    tb.cell(r1, 0).merge(tb.cell(r2, 0))
tb.cell(0, 0).merge(tb.cell(0, 1)); tb.cell(len(R) - 1, 0).merge(tb.cell(len(R) - 1, 1))

# ── 우측: 보수 합계 (자리츠1 · 자리츠2 · 모리츠) ──────────────────────────
fill(find(s, 46), [[("보수 ", 0), ("(자리츠1 · 자리츠2 · 모리츠)", 3)]])
old = find(s, 47); xr, wr = old.left / 914400, old.width / 914400
old._element.getparent().remove(old._element)
FR = [["구 분", "자리츠1\n(72호)", "자리츠2\n(본건)", "모리츠\n(KLI)", "계"],
      ["매입", "26.2", "13.5", "10.3", "50.0"],
      ["운용", "118.1", "66.7", "-", "184.8"],
      ["매각기본", "34.6", "18.8", "-", "53.4"],
      ["매각성과", "175.9", "124.8", "-", "300.7"],
      ["합 계", "354.8", "223.8", "10.3", "588.9"]]
for r_ in FR[1:-1]:
    vals_ = [float(v) for v in r_[1:4] if v != "-"]
    assert abs(sum(vals_) - float(r_[4])) < 0.11, r_
def fr_fmt(ri, ci, v):
    if ri == 0:
        return (NBB, 9, WHITE, HEAD2 if ci == 4 else HEAD, "ctr")
    if ri == len(FR) - 1:
        return (NXB, 10.5 if ci == 4 else 9.5, RED if ci == 4 else INK, TOT, "ctr" if ci == 0 else "r")
    if ci == 0:
        return (NBB, 9, INK, LAB, "ctr")
    return (NBB if ci == 4 else NB, 9, INK, None, "r")
new_table(FR, xr, 1.32, [0.68, 0.66, 0.66, 0.63, 0.70], [0.42, 0.34, 0.34, 0.34, 0.34, 0.40], fr_fmt)
note = s.shapes.add_textbox(IN(xr), IN(3.55), IN(3.33), IN(1.05))
note.text_frame.word_wrap = True
lines = ["· 자리츠1: 매입가 5,230억 × 0.5% / AUM 5,759.9억 × 0.223% × 약 10년(C자산 중도매각 반영) / T.Cap 4.93%, *임차인 우선매수권 행사 시 매각기본 0.3%·성과보수 제외",
         "· 자리츠2: 매입가 2,700억 × 0.5% / AUM 2,989억 × 0.223% × 10년 / T.Cap 4.89%",
         "· 모리츠: 투자금액 × 0.5% (72호 1,193억 5.9 + 본건 888억 4.4[추정]), 운용보수 출자분 미청구(감사팀 의견)"]
note.text_frame.text = lines[0]
for ln in lines[1:]:
    note.text_frame.add_paragraph().text = ln
for p_ in note.text_frame.paragraphs:
    for r_ in p_.runs:
        r_.font.size = Pt(7.5); r_.font.name = NB
        rp_ = r_._r.get_or_add_rPr()
        for tag in ("ea", "cs"):
            etree.SubElement(rp_, qn(f"a:{tag}"), typeface=NB)

# ── 우측 하단: 모리츠(KLI) 투자수익 (참고) ────────────────────────────────
fill(find(s, 49), [[("모리츠(KLI) 투자수익 ", 0), ("(참고)", 3)]])
t = find(s, 50).table
cells = {(0, 0): "구 분", (0, 1): "자리츠1 (72호)", (0, 2): "자리츠2 (본건)",
         (1, 0): "투자금액", (1, 1): "1,193억 (1종)", (1, 2): "888억\n(보통주·1·2종)",
         (2, 0): "수익률", (2, 1): "CoC 7.00%\nIRR 8.32%", (2, 2): "CoC 5.44%\nIRR 10.99%",
         (3, 0): "투자수익\n(10년)", (3, 1): "[미기재]", (3, 2): "1,071.4억\n(I.G 483 · C.G 588)"}
for (ri, ci), v in cells.items():
    head = ri == 0
    lab = ci == 0 and not head
    style_cell(t.cell(ri, ci), v, NBB if (head or lab) else NB, 9 if head else 8.5,
               WHITE if head else INK, HEAD if head else (LAB if lab else (TOT if ri == 3 else None)), "ctr")
for i, w in enumerate((0.85, 1.24, 1.24)):
    t.columns[i].width = IN(w)

# 출처 각주
fn = s.shapes.add_textbox(IN(0.4), IN(7.18), IN(12.6), IN(0.30))
fn.text_frame.word_wrap = True
fn.text_frame.text = ("※ 출처: 자리츠1 = 현대자동차 실물자산 유동화 프로젝트 투심자료 v1.0(’26.03.26) p.3·4·15·31 / 자리츠2 = 투심자료 v12(’26.10.02)·기준 재무모델 v04 "
                      "(A&R·Sell-down 탭, 중순위 대주 10/06 의견 기준). 자리츠1 운용보수 연도 배분은 10년 균등 가정[추정], 연도는 각 리츠 사업연도 기준. "
                      "모리츠 본건 매입보수 = KLI 취득금액 × 0.5% 산식 적용[추정] — 투심자료 v12 KLI 투자개요(597백만원)와 차이 확인 필요. 당사 직접 출자(PI) 없음. 수익·보수는 예상치이며 확정된 것이 아님.")
r = fn.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(7); r.font.name = NB
rp = r._r.get_or_add_rPr()
for tag in ("ea", "cs"):
    etree.SubElement(rp, qn(f"a:{tag}"), typeface=NB)

# ══════════════════════════════════════════════════════════════════════
# 장표 ② 제반 실사 업체 선정 현황
# ══════════════════════════════════════════════════════════════════════
s = S2
fill(find(s, 7), [[("총  ", 0), ("270백만원", 1), ("  (재무모델 견적 기준)", 2)]])
cols = [  # (제목, 범위, 배지, 업체 박스(이름, 금액, 단위), 표, 각주) — 수수료: 기준 재무모델 v03 실사/자문비용
    (11, 12, 14, 15, 16, 45, "법률실사",
     ["법률실사 진행 및 보고서 작성", "매매·주주간·대출 계약서 작성·검토", "신설리츠 투자구조 법률의견"],
     ("태평양(BKL)", "80", "백만원"), ("법무법인 태평양", "80 백만원"), "※ 법률실사보고서 ’26.09.09"),
    (18, 19, 21, 22, 23, 47, "감정평가",
     ["감정평가서(담보·시가) 작성", "감정평가액 3,345억원", "기준시점 ’26.08.31"],
     ("미래새한", "85", "백만원"), ("미래새한감정평가", "85.3백만원"), "※ 담보 감정평가서(우리은행 외)"),
    (25, 26, 28, 29, 30, 49, "재무실사",
     ["10년 현금흐름·투자자 수익률 분석", "DSCR 등 재무지표 검토", "재무분석보고서 작성"],
     ("우리회계법인", "40", "백만원"), ("우리회계법인", "40 백만원"), "※ 재무분석보고서 ’26.08.14"),
    (32, 33, 35, 36, 37, 51, "물리실사",
     ["건축·구조·기계·전기·소방 실사", "CAPEX 추정", "물리실사보고서 작성"],
     ("CBRE", "25", "백만원"), ("CBRE", "25 백만원"), "※ 물리실사보고서 ’26.08.31"),
    (39, 40, 42, 43, 44, None, "시장실사",
     ["시장조사보고서 작성", "매매·임대 거래사례 조사", "적정 임대료(E.NOC) 분석"],
     ("Savills", "40", "백만원"), ("세빌스(Savills)", "40 백만원"), "※ 시장조사보고서 ’26.08.10"),
]
note_tmpl = find(s, 51)._element
for ti, bi, gi, fi, tbi, ni, title, scope, firm, row, note in cols:
    simple(find(s, ti), title)
    fill(find(s, bi), [[(sc, 1)] for sc in scope])
    simple(find(s, gi), "견적 기준")
    fill(find(s, fi), [[(firm[0], 0)], [(firm[1], 0), (firm[2], 1)]])
    t = find(s, tbi).table
    del_rows(t, 3)
    cell_set(t.cell(0, 1), "제안가격")
    cell_set(t.cell(1, 0), row[0]); cell_set(t.cell(1, 1), row[1])
    cell_set(t.cell(2, 0), "타 제안사"); cell_set(t.cell(2, 1), "[미접수]")
    gf = find(s, tbi); gf.height = Emu(sum(r.height for r in t.rows))
    if ni is None:                                   # 시장실사 각주 신설 (물리실사 각주 복제)
        el = copy.deepcopy(note_tmpl); s.shapes._spTree.append(el)
        nsh = [sh for sh in s.shapes if sh._element is el][0]
        nsh.left = find(s, tbi).left
    else:
        nsh = find(s, ni)
    fill(nsh, [[(note, 0)]])
    nsh.top = gf.top + gf.height + IN(0.04); nsh.width = IN(2.05)
fn = s.shapes.add_textbox(IN(0.49), IN(7.18), IN(12.4), IN(0.26))
fn.text_frame.text = ("※ 수수료: 기준 재무모델 v03(’26.09.30) 실사/자문비용 견적 반영액(감정 85.3·법률 80·재무 40·시장 40·물리 25, 합계 270.3백만원, "
                      "소유권이전비용 76백만원 별도). VAT 포함 여부·성공불 여부·경쟁 제안가격 [확인 필요].")
r = fn.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(7.5); r.font.name = NB
rp = r._r.get_or_add_rPr()
for tag in ("ea", "cs"):
    etree.SubElement(rp, qn(f"a:{tag}"), typeface=NB)
fn.text_frame.word_wrap = True

# 세로 중간 정렬 (운용역 지시 2026-10-01)
for sl in (S1, S2):
    for el in sl.shapes._spTree.iter(qn("a:bodyPr")):
        el.set("anchor", "ctr")
    for el in sl.shapes._spTree.iter(qn("a:tcPr")):
        el.set("anchor", "ctr")

prs.save(OUT)
print("saved", OUT)
