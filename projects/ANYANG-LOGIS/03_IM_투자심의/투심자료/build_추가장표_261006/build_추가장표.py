# -*- coding: utf-8 -*-
"""ANYANG-LOGIS 추가 장표 (운용역 제공 양식 261006): ① 투자구조 및 당사 손익 ② 제반 실사 업체 선정 현황.

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
# 장표 ① 투자구조 및 당사 손익
# ══════════════════════════════════════════════════════════════════════
s = S1
simple(find(s, 2), "투자구조 및 당사 손익")
fill(find(s, 8), [[("코람코라이프인프라리츠", 0)], [("(모리츠 · 상장리츠, 포트폴리오 약 1.68조)", 0)]])
fill(find(s, 65), [[("· ", 0), ("자리츠(코람코라이프로지스리츠) 기준", 1)]])
simple(find(s, 51), "예상수익률")
simple(find(s, 59), "IRR 10.21%")
simple(find(s, 60), "매입가 0.5%")
simple(find(s, 61), "AUM × 0.223%")
simple(find(s, 63), "매각가 0.5%")
for sid, sz in ((59, 9), (60, 8), (63, 8.5)):
    for r_ in find(s, sid).text_frame.paragraphs[0].runs:
        r_.font.size = Pt(sz)
fill(find(s, 64), [[("기본보수 차감 후", 0)], [("매각이익의 15%", 0)]])
fill(find(s, 13), [[("※ 매입확약 이행 후 : ", 0), ("KLI 888억 (보통주 148 · 2종 500 · 1종 240)", 1)]])
find(s, 13).width = IN(4.6)
simple(find(s, 15), "+ KLI 308억")
# 단일 자리츠 구조 → 자2~4 및 분기선 삭제
drop(s, 16, 18, 19, 20, 23, 24, 25)
fill(find(s, 22), [[("코람코라이프로지스리츠", 0)], [("(2,700억) ", 1), ("안양물류센터", 1)]])
g = find(s, 22); g.width = IN(2.5)
fill(find(s, 27), [[("매매대금", 0)], [("2,700", 0)]])
fill(find(s, 28), [[("부대비용", 0), (" ", 1), ("189", 2)]])
fill(find(s, 29), [[("예비비 ", 0), ("100", 1)]])
fill(find(s, 30), [[("Equity", 0)], [("988", 0)]])
fill(find(s, 31), [[("Loan", 0)], [("1,980", 0)]])
fill(find(s, 32), [[("보증금", 0), (" ", 1), ("21", 2)]])
# Sources 비율 (Equity 988 : Loan 1,980) 반영 — 높이 합 1.25" 유지
e, l = find(s, 30), find(s, 31)
top, tot = e.top, e.height + l.height
e.height = Emu(int(tot * 0.36)); l.top = e.top + e.height; l.height = tot - e.height
find(s, 33).top = e.top + e.height // 2
find(s, 35).top = l.top + l.height // 2
fill(find(s, 34), [[("KLI ", 0), ("308억 ", 1), ("(31.2%)", 2)], [("FI ", 0), ("580억 ", 1), ("(58.7%)", 2)],
                   [("매도인 3종 ", 0), ("100억 ", 1), ("(10.1%)", 2)]])
b34 = find(s, 34); b34.top = e.top - IN(0.10); b34.height = IN(0.58); b34.width = IN(2.0)
fill(find(s, 36), [[("LTV 59.8%", 0)], [("(감정가 3,345억 기준)", 0)], [("선순위 1,600 · 중순위 380", 0)],
                   [("만기 24개월", 0)]])
b36 = find(s, 36); b36.top = l.top + IN(0.10); b36.height = IN(0.80); b36.width = IN(2.0)
for b_ in (b34, b36):
    for p_ in b_.text_frame.paragraphs:
        for r_ in p_.runs:
            r_.font.size = Pt(min(r_.font.size.pt if r_.font.size else 9, 9))
fill(find(s, 38), [[("[ ", 0), ("대주 구성 ]", 1)]])
t = find(s, 39).table
add_row(t)
rows = [("기관", "금액", "비고"),
        ("우리은행", "1,600억", "선순위 4.78% (All-in 5.50%)"),
        ("키움캐피탈", "200억", "중순위 6.00% (All-in 6.50%)"),
        ("MG캐피탈", "100억", "중순위 (상동)"),
        ("DB저축은행", "80억", "중순위 (상동) · LOC 미접수")]
for ri, row in enumerate(rows):
    for ci, v in enumerate(row):
        cell_set(t.cell(ri, ci), v)
for i_, w_ in enumerate((0.95, 0.72, 1.73)):
    t.columns[i_].width = IN(w_)
gf = find(s, 39); gf.height = Emu(sum(r.height for r in t.rows)); gf.left = IN(5.95); gf.width = IN(3.4)
gf.top = IN(3.42)
find(s, 38).top = IN(3.10); find(s, 38).left = IN(5.95)
find(s, 37).top = IN(3.95); find(s, 37).left = IN(5.45)

# 당사 손익 표 (사업연도 기준: ’26 = 거래종결, ’27~’36 = 매년 10월말 종료 사업연도)
fill(find(s, 41), [[("당사 손익 ", 0), ("(자리츠 기준)", 4)]])
old = find(s, 43)
x, y, w = old.left, old.top, old.width
old._element.getparent().remove(old._element)
yrs = ["’26"] + [f"’{y_}" for y_ in range(27, 37)]
fee_m = [None] + [6.66547] * 10                     # 운용보수 666.547백만/년 (모델 A&R, AUM 2,989억 × 0.223%)
sell = [None] * 10 + [18.793038 + 124.803226]       # 매각기본 18.8 + 매각성과 124.8 ('36.10 매각, 모델 A&R)
buy = [13.5] + [None] * 10
tot_y = [(b or 0) + (m or 0) + (sl or 0) for b, m, sl in zip(buy, fee_m, sell)]
f = lambda v: "" if v is None else f"{v:,.1f}"
R = [["구 분", ""] + yrs + ["계"],
     ["보\n수", "매입"] + [f(v) for v in buy] + ["13.5"],
     ["", "운용"] + [f(v) for v in fee_m] + ["66.7"],
     ["", "매각"] + [f(v) for v in sell] + ["143.6"],
     ["계", ""] + [f(v) for v in tot_y] + ["223.8"],
     ["P\nI", "I.G · C.G"] + ["-"] * 11 + ["-"],
     ["합 계", ""] + [f(v) for v in tot_y] + ["223.8"]]
assert abs(sum(tot_y) - 223.8) < 0.06, sum(tot_y)
nc = len(R[0])
gf = s.shapes.add_table(len(R), nc, x, y, w, IN(1.84))
tb = gf.table
tblPr = gf._element.graphic.graphicData.tbl.tblPr
for a in ("firstRow", "bandRow"):
    tblPr.set(a, "0")
ws = [0.19, 0.62] + [0.61] * 11 + [0.68]
k = (w / 914400) / sum(ws)
for i, v in enumerate(ws):
    tb.columns[i].width = IN(v * k)
for ri, h in enumerate([0.38, 0.22, 0.22, 0.22, 0.24, 0.24, 0.30]):
    tb.rows[ri].height = IN(h)
for ri, row in enumerate(R):
    for ci, v in enumerate(row):
        c = tb.cell(ri, ci)
        if ri == 0:
            style_cell(c, v, NBB, 9.5, WHITE, HEAD2 if ci == nc - 1 else HEAD, "ctr")
        elif ri in (4, 6):
            fz = TOT if ri == 6 else LAB
            fn = NXB if ri == 6 else NBB
            style_cell(c, v, fn, 8, INK, fz, "ctr" if ci < 2 else "r")
        elif ci == 0:
            style_cell(c, v, NBB, 8, INK, LAB, "ctr")
        elif ci == 1:
            style_cell(c, v, NB, 8, INK, LAB, "ctr")
        else:
            style_cell(c, v, NBB if ci == nc - 1 else NB, 8, INK, "F2F2F2" if ci == 2 and ri < 4 else None,
                       "ctr" if ri == 5 else "r")
for (r1, c1, r2, c2) in ((0, 0, 0, 1), (1, 0, 3, 0), (4, 0, 4, 1), (6, 0, 6, 1)):
    tb.cell(r1, c1).merge(tb.cell(r2, c2))
fill(find(s, 42), [[("(단위 : 억원)", 0)]])

# 보수 표 (자리츠 기준, 투심자료 v12 '참고. 당사 예상수익')
fill(find(s, 46), [[("보수 ", 0), ("(자리츠 기준)", 3)]])
t = find(s, 47).table
vals = {(1, 2): "13.5억", (1, 3): "· 2,700억 × 0.5%",
        (2, 0): "운용\n(10년)", (2, 2): "66.7억", (2, 3): "· 2,989억 × 0.223%\n  × 약 10년",
        (3, 2): "18.8억", (3, 3): "· 매각금액 × 0.5%\n  (T.Cap 4.89%)",
        (4, 2): "124.8억", (4, 3): "· 기본보수 차감 후\n  매각이익 × 15%",
        (5, 2): "143.6억", (6, 2): "223.8억"}
tmpl_body = copy.deepcopy(t.cell(2, 3).text_frame._txBody)
for (ri, ci), v in vals.items():
    cell = t.cell(ri, ci)
    if not cell.text_frame.text.strip() and ci == 3:
        body = cell.text_frame._txBody
        for p in body.findall(qn("a:p")):
            body.remove(p)
        for p in tmpl_body.findall(qn("a:p")):
            body.append(copy.deepcopy(p))
    cell_set(cell, *v.split("\n"))
for ri in range(1, 7):                              # 산출근거 열 좌측 정렬
    for p in t.cell(ri, 3).text_frame.paragraphs:
        p.alignment = PP_ALIGN.LEFT

# PI 투자수익 → (참고) 모리츠 KLI 투자수익 (당사 직접 출자 없음)
fill(find(s, 49), [[("모리츠(KLI) 투자수익 ", 0), ("(참고)", 3)]])
t = find(s, 50).table
for (ri, ci), v in {(1, 1): "483.0억", (1, 2): "· 운영배당 누적\n  (CoC 5.44%)",
                    (2, 1): "588.4억", (2, 2): "· ’36.10 매각 기준\n  매각차익 배분",
                    (3, 1): "1,071.4억", (3, 2): "· KLI 888억 투자\n  IRR 10.99%"}.items():
    cell_set(t.cell(ri, ci), *v.split("\n"))
for ri in range(1, 4):
    for p in t.cell(ri, 2).text_frame.paragraphs:
        p.alignment = PP_ALIGN.LEFT
for r_ in t.cell(3, 1).text_frame.paragraphs[0].runs:
    r_.font.size = Pt(12)
# 출처 각주
fn = s.shapes.add_textbox(IN(0.4), IN(7.20), IN(12.6), IN(0.26))
fn.text_frame.text = "x"
fill_src = fn.text_frame
p = fill_src.paragraphs[0]
p.runs[0].text = ("※ 출처: 투심자료 v12(’26.10.02) 재원조달·지분구성·당사 예상수익, 기준 재무모델 v03(’26.09.30) A&R·Sell-down 탭(KLI는 보통주 ’26.10 출자, "
                  "2종 ’27.10·1종 ’28.10 매입, ’36.10 매각 가정). 중순위 대주 구성은 대출약정서 v16 중순위 의견본(’26.10.06) 기준. "
                  "연도는 사업연도(매년 10월말) 기준, 단수 차이 있음. 당사 직접 출자(PI)는 없음. 수익·보수는 예상치이며 확정된 것이 아님.")
r = p.runs[0]; r.font.size = Pt(7.5); r.font.name = NB
rp = r._r.get_or_add_rPr()
for tag in ("ea", "cs"):
    etree.SubElement(rp, qn(f"a:{tag}"), typeface=NB)
fn.text_frame.word_wrap = True

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
