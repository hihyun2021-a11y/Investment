# -*- coding: utf-8 -*-
"""부동산담보신탁계약서 및 특약 v05 — 신탁사(우리은행 신탁부) 의견 반영 + 대출약정서 v27 기준 우선수익자(대주 9개사) 기재 + 서식 통일.

입력(원본 불변): 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v03_당사자정리_변경추적.docx
대주 자료: 05_대출업무/약정서/20261013_대출약정_담보대출약정서_BKL_v27_markup.docx (INDEX #82) — 별지(대주목록·대출약정금)·부록 5(통지처)
법인등록번호: 사업자등록증 #83 롯데캐피탈·#84 디비저축은행·#85 수협은행, #81 신협통지처(최종)(신협중앙회), 기존 기재값(우리은행)
출력: 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v05_신탁사의견반영_변경추적.docx

- v04(대주통지처 엑셀 기준 22개사)를 대체한다. 대출약정서 v27은 신협을 신용협동조합중앙회 1개 대주(300억원)로 두므로 우선수익자도 9개사로 맞춘다.
- 내용 변경은 Word 변경추적(작성자 "코람코 상장리츠팀") + 메모. v03까지의 변경추적·메모는 유지(특약 말미 v03 삽입 문구 1건만 신탁사 문안으로 대체).
- 서식 통일(줄간격·글꼴·크기·여백)은 v04와 같이 변경추적 없이 직접 적용.
실행: python3 build_담보신탁_v05_신탁사의견_대주9개사_261013.py
"""
import copy
import os
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v03_당사자정리_변경추적.docx")
LOAN = os.path.join(HERE, "20261013_대출약정_담보대출약정서_BKL_v27_markup.docx")
OUT = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v05_신탁사의견반영_변경추적.docx")
if len(sys.argv) > 1:
    OUT = sys.argv[1]

AUTHOR, INIT, DATE = "코람코 상장리츠팀", "KR", "2026-10-13T15:00:00Z"
_id = [9500]
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


# ================================================================ 공통 도우미 (v03 빌더와 동일 방식)
def nid():
    _id[0] += 1
    return str(_id[0])


def mark(tag):
    el = OxmlElement(tag)
    el.set(qn("w:id"), nid())
    el.set(qn("w:author"), AUTHOR)
    el.set(qn("w:date"), DATE)
    return el


def to_del(r):
    for t in r.findall(qn("w:t")):
        dt = OxmlElement("w:delText")
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        dt.text = t.text
        t.addprevious(dt)
        r.remove(t)
    d = mark("w:del")
    r.addprevious(d)
    d.append(r)
    return d


def new_run(text, rpr_src=None):
    r = OxmlElement("w:r")
    if rpr_src is not None and rpr_src.find(qn("w:rPr")) is not None:
        r.append(copy.deepcopy(rpr_src.find(qn("w:rPr"))))
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def direct_runs(p):
    return [c for c in p._p if c.tag == qn("w:r")]


def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def split_run(r, pos):
    txt = run_text(r)
    if pos <= 0 or pos >= len(txt):
        return r
    r2 = copy.deepcopy(r)
    ts = r.findall(qn("w:t"))
    for t in ts[1:]:
        r.remove(t)
    ts[0].text = txt[:pos]
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t2s = r2.findall(qn("w:t"))
    for t in t2s[1:]:
        r2.remove(t)
    t2s[0].text = txt[pos:]
    t2s[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.addnext(r2)
    return r


def track_replace(p, old, new, ctx=None, after=False):
    runs = direct_runs(p)
    full = "".join(run_text(r) for r in runs)
    if ctx is not None:
        c0 = full.find(ctx)
        assert c0 >= 0, f"문맥 미발견: {ctx[:30]}"
        i = c0 + ctx.index(old)
    else:
        i = full.find(old)
    assert i >= 0, f"'{old[:30]}' 미발견: {full[:80]}"
    if after:
        i, old = i + len(old), ""
    j = i + len(old)

    def cut(pos):
        acc = 0
        for r in direct_runs(p):
            L = len(run_text(r))
            if acc < pos < acc + L:
                split_run(r, pos - acc)
                return
            acc += L
    cut(i)
    cut(j)
    acc, covered, before = 0, [], None
    for r in direct_runs(p):
        L = len(run_text(r))
        if acc >= i and acc + L <= j and L > 0:
            covered.append(r)
        if acc + L <= i:
            before = r
        acc += L
    tmpl = covered[0] if covered else before
    anchor = None
    for r in covered:
        anchor = to_del(r)
    if new == "":
        return anchor
    ins = mark("w:ins")
    ins.append(new_run(new, tmpl))
    if anchor is not None:
        anchor.addnext(ins)
    elif before is not None:
        before.addnext(ins)
    else:
        ppr = p._p.pPr
        (ppr.addnext(ins) if ppr is not None else p._p.insert(0, ins))
    return ins


def track_set(p, text, rpr_tmpl):
    """칸 문단의 기존 글자를 모두 삭제 표시하고 text를 삽입 표시한다(빈 칸이면 삽입만)."""
    for r in direct_runs(p):
        if run_text(r):
            to_del(r)
        else:
            p._p.remove(r)
    ins = mark("w:ins")
    ins.append(new_run(text, rpr_tmpl))
    p._p.append(ins)
    return ins


def comment(doc, p_el, text):
    para = Paragraph(p_el, doc._body)
    a = OxmlElement("w:r")
    b = OxmlElement("w:r")
    ppr = p_el.find(qn("w:pPr"))
    (ppr.addnext(a) if ppr is not None else p_el.insert(0, a))
    p_el.append(b)
    doc.add_comment([Run(a, para), Run(b, para)], text=text, author=AUTHOR, initials=INIT)


def find(ps, start, prefix):
    for k in range(start, len(ps)):
        if ps[k].text.strip().startswith(prefix):
            return k
    raise ValueError(prefix)


def strip_ids(el):
    for e in el.iter():
        for a in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
            if a in e.attrib:
                del e.attrib[a]


def make_rpr(sz=22, b=False, u=False):
    rpr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts")
    for k in ("ascii", "hAnsi", "cs"):
        f.set(qn(f"w:{k}"), "Times New Roman")
    f.set(qn("w:eastAsia"), "바탕체")
    rpr.append(f)
    if b:
        rpr.append(OxmlElement("w:b"))
        rpr.append(OxmlElement("w:bCs"))
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "000000")
    rpr.append(c)
    k = OxmlElement("w:kern")
    k.set(qn("w:val"), "0")
    rpr.append(k)
    for tag in ("w:sz", "w:szCs"):
        s = OxmlElement(tag)
        s.set(qn("w:val"), str(sz))
        rpr.append(s)
    if u:
        uu = OxmlElement("w:u")
        uu.set(qn("w:val"), "single")
        rpr.append(uu)
    return rpr


def run_tmpl(sz=22, b=False, u=False):
    r = OxmlElement("w:r")
    r.append(make_rpr(sz, b, u))
    return r


def ins_para(ref_el, text, sz=22, b=False, u=False, jc=None, line=360):
    """ref_el 다음에 삽입 표시된 새 문단을 만든다(본계약 첫 페이지 서식)."""
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    sg = OxmlElement("w:snapToGrid")
    sg.set(qn("w:val"), "0")
    ppr.append(sg)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), str(line))
    sp.set(qn("w:lineRule"), "auto")
    ppr.append(sp)
    if jc:
        j = OxmlElement("w:jc")
        j.set(qn("w:val"), jc)
        ppr.append(j)
    mrpr = make_rpr(sz, b, u)
    mrpr.insert(0, mark("w:ins"))
    ppr.append(mrpr)
    p.append(ppr)
    if text:
        ins = mark("w:ins")
        ins.append(new_run(text, run_tmpl(sz, b, u)))
        p.append(ins)
    ref_el.addnext(p)
    return p


# ---------------------------------------------------------------- 새 표(삽입 표시)
def _cell(w, text, sz, b=False, jc="center", vm=None, span=1, fill=None):
    tc = OxmlElement("w:tc")
    pr = OxmlElement("w:tcPr")
    tw = OxmlElement("w:tcW")
    tw.set(qn("w:w"), str(w))
    tw.set(qn("w:type"), "dxa")
    pr.append(tw)
    if span > 1:
        gs = OxmlElement("w:gridSpan")
        gs.set(qn("w:val"), str(span))
        pr.append(gs)
    if vm:
        v = OxmlElement("w:vMerge")
        if vm == "restart":
            v.set(qn("w:val"), "restart")
        pr.append(v)
    if fill:
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear")
        sh.set(qn("w:color"), "auto")
        sh.set(qn("w:fill"), fill)
        pr.append(sh)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    pr.append(va)
    tc.append(pr)
    lines = (text or "").split("\n")
    for ln in lines:
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        sg = OxmlElement("w:snapToGrid")
        sg.set(qn("w:val"), "0")
        ppr.append(sg)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:after"), "0")
        sp.set(qn("w:line"), "240")
        sp.set(qn("w:lineRule"), "auto")
        ppr.append(sp)
        j = OxmlElement("w:jc")
        j.set(qn("w:val"), jc)
        ppr.append(j)
        mrpr = make_rpr(sz, b)
        mrpr.insert(0, mark("w:ins"))
        ppr.append(mrpr)
        p.append(ppr)
        if ln and vm != "cont":
            ins = mark("w:ins")
            ins.append(new_run(ln, run_tmpl(sz, b)))
            p.append(ins)
        tc.append(p)
    return tc


def make_table(widths, rows, sz):
    """rows: [[(text, opts), ...], ...]  opts: b, jc, vm('restart'/'cont'), span, fill"""
    tbl = OxmlElement("w:tbl")
    pr = OxmlElement("w:tblPr")
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(sum(widths)))
    tw.set(qn("w:type"), "dxa")
    pr.append(tw)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pr.append(jc)
    bd = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        bd.append(e)
    pr.append(bd)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    pr.append(lay)
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", 28), ("left", 57), ("bottom", 28), ("right", 57)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    pr.append(mar)
    tbl.append(pr)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        g = OxmlElement("w:gridCol")
        g.set(qn("w:w"), str(w))
        grid.append(g)
    tbl.append(grid)
    for ri, row in enumerate(rows):
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        if ri == 0:
            trpr.append(OxmlElement("w:tblHeader"))
        cs = OxmlElement("w:cantSplit")
        trpr.append(cs)
        trpr.append(mark("w:ins"))
        tr.append(trpr)
        col = 0
        for text, o in row:
            span = o.get("span", 1)
            w = sum(widths[col:col + span])
            tr.append(_cell(w, text, sz, b=o.get("b", ri == 0), jc=o.get("jc", "center"),
                            vm=o.get("vm"), span=span, fill=o.get("fill", "F2F2F2" if ri == 0 else None)))
            col += span
        tbl.append(tr)
    return tbl


def mark_row_inserted(tr, texts, rpr_tmpl):
    """복제한 행을 '삽입된 행'으로 만들고 칸별 글자를 채운다(texts[i] None이면 비움)."""
    strip_ids(tr)
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = OxmlElement("w:trPr")
        tr.insert(0, trpr)
    trpr.append(mark("w:ins"))
    for tc, txt in zip(tr.findall(qn("w:tc")), texts):
        for p in tc.findall(qn("w:p")):
            for c in list(p):
                if c.tag != qn("w:pPr"):
                    p.remove(c)
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                p.insert(0, ppr)
            mrpr = ppr.find(qn("w:rPr"))
            if mrpr is None:
                mrpr = OxmlElement("w:rPr")
                ppr.append(mrpr)
            mrpr.insert(0, mark("w:ins"))
        ps_ = tc.findall(qn("w:p"))
        for extra in ps_[1:]:
            tc.remove(extra)
        if txt:
            ins = mark("w:ins")
            ins.append(new_run(txt, rpr_tmpl))
            ps_[0].append(ins)
    return tr


# ================================================================ 대주 자료 — 대출약정서 v27(INDEX #82) 별지·부록 5 기준
REGNO = {  # 법인등록번호 근거
    "주식회사 우리은행": ("110111-0023393", "기존 신탁계약서 기재값"),
    "롯데캐피탈 주식회사": ("110111-1217416", "사업자등록증 #83"),
    "신용협동조합중앙회": ("115041-0000361", "대주통지처 vf 신협통지처(최종) #81"),
    "수협은행": ("244235-0009895", "사업자등록증(사당역지점) #85"),
    "주식회사 디비저축은행": ("110111-0135627", "사업자등록증 #84"),
}


def parse_notice(txt):
    """'주소: …\n전화번호: …' 형식 → dict. 키 없는 줄은 앞 항목에 이어 붙인다."""
    out, last = {}, None
    keymap = {"주소": "addr", "전화번호": "tel", "팩스번호": "fax", "담당자": "person",
              "Email": "email", "전자우편": "email", "수취계좌": "acct"}
    for ln in [x.strip() for x in str(txt).split("\n") if x.strip()]:
        m = re.match(r"^(주소|전화번호|팩스번호|담당자|Email|전자우편|수취계좌)\s*[:/]\s*(.*)$", ln)
        if m:
            last = keymap[m.group(1)]
            out[last] = m.group(2).strip()
        elif last:
            out[last] += " / " + ln
    return out


def office_phones(s):
    """휴대전화(010) 번호는 빼고 사무실 번호만 남긴다(신탁원부 공시 고려)."""
    parts = [x.strip() for x in re.split(r"[/,]", s or "") if x.strip()]
    keep = [x.replace(".", "-") for x in parts if not re.sub(r"\D", "", x).startswith("010")]
    return " / ".join(keep)


def clean_person(s):
    return re.sub(r"\s*/\s*", ", ", (s or "").strip(" /"))


def _t(el):
    """삭제 표시(delText) 제외 = 변경 반영본. 줄바꿈(w:br)은 개행으로."""
    return "".join("\n" if e.tag in (qn("w:br"), qn("w:cr")) else (e.text or "")
                   for e in el.iter(qn("w:t"), qn("w:br"), qn("w:cr")))


loan = Document(LOAN)
amt_by_name, tranche_by_name = {}, {}
for tbl in loan.tables:
    head = " ".join(_t(tc) for tc in tbl._tbl.findall(qn("w:tr"))[0].findall(qn("w:tc")))
    if "대출약정금" in head and "대주명" in head:
        cur = None
        for tr in tbl._tbl.findall(qn("w:tr"))[1:]:
            cs = [_t(tc).strip() for tc in tr.findall(qn("w:tc"))]
            if any("Tranche A-1 대주" == c for c in cs):
                cur = "Tranche A-1 대출"
            if any("Tranche A-2 대주" == c for c in cs):
                cur = "Tranche A-2 대출"
            if any(c == "Tranche B 대주" for c in cs):
                cur = "Tranche B 대출"
            for k, c in enumerate(cs):
                m = re.search(r"￦([\d,]+)\)", c)
                if m and k > 0 and "합계" not in cs[k - 1] and cs[k - 1] and "총" not in cs[0] + cs[k - 1]:
                    amt_by_name[cs[k - 1]] = int(m.group(1).replace(",", ""))
                    tranche_by_name[cs[k - 1]] = cur
        break
notice_rows = []
for tbl in loan.tables:
    rows = tbl._tbl.findall(qn("w:tr"))
    head = [_t(tc).strip() for tc in rows[0].findall(qn("w:tc"))]
    if head[:3] == ["순번", "대주명", "통지처"]:
        for tr in rows[1:]:
            tcs = tr.findall(qn("w:tc"))
            nm = _t(tcs[1]).strip()
            body_txt = "\n".join(_t(p) for p in tcs[2].findall(qn("w:p")))
            notice_rows.append((nm, parse_notice(body_txt)))
        break
assert len(notice_rows) == 9, notice_rows


def formal(short):
    return next(n for n in amt_by_name if n.startswith(short) or short.startswith(n))


TA, TB = [], []
for short, n in notice_rows:
    name = formal(short)
    x = dict(name=name, tranche=tranche_by_name[name], amt=amt_by_name[name],
             regno=REGNO.get(name, ("[*]", ""))[0], **n)
    (TB if x["tranche"] == "Tranche B 대출" else TA).append(x)
assert TA[0]["name"] == "주식회사 우리은행" and len(TA) == 6 and len(TB) == 3
assert sum(x["amt"] for x in TA) == 160_000_000_000, "Tranche A 합계 불일치"
assert sum(x["amt"] for x in TB) == 38_000_000_000, "Tranche B 합계 불일치"
for x in TA + TB:
    x["cap"] = x["amt"] * 12 // 10
    x["addr"] = x.get("addr", "").replace(" ,", ",").strip()
    x["tel"] = office_phones(x.get("tel"))
    x["fax"] = (x.get("fax") or "").replace(".", "-").strip()
    x["person"] = clean_person(x.get("person"))
    x["email"] = re.sub(r"\s*[/,]\s*", "\n", (x.get("email") or "").strip())
CAP_A = sum(x["cap"] for x in TA)
CAP_B = sum(x["cap"] for x in TB)
COUPANG_CAP = 2_132_586_000
assert CAP_A == 192_000_000_000 and CAP_B == 45_600_000_000
assert COUPANG_CAP + CAP_A + CAP_B == 239_732_586_000
N_OTHERS = len(TA) - 1   # '주식회사 우리은행 외 N개사' (대출약정서 부록 5 Tranche A 통지처 수 − 1)
MISSING = [x["name"] for x in TA + TB if x["regno"] == "[*]"]


def won(v):
    return f"{v:,}"


# ================================================================ 문서 열기
doc = Document(SRC)
ps = doc.paragraphs
tables = doc.tables
k_tk = find(ps, 0, "부동산담보신탁계약에 대한 특약")
tk_title = ps[k_tk]._p


def roster_rows():
    """우선수익자 세부 목록(별첨 1) 표 행."""
    rows = [[("순위", {}), ("연번", {}), ("우선수익자(상호)", {}), ("법인등록번호", {}),
             ("대출약정금 등(원)", {}), ("우선수익한도금액(원)", {})]]
    rows.append([("제1순위", {}), ("1", {}), ("쿠팡 주식회사", {"jc": "left"}), ("[*]", {}),
                 ("2,132,586,000\n(임대차보증금)", {"jc": "right"}), (won(COUPANG_CAP), {"jc": "right"})])
    for i, x in enumerate(TA):
        rows.append([("(공동)\n제2순위\n(Tranche A)", {"vm": "restart"}) if i == 0 else ("", {"vm": "cont"}),
                     (str(i + 1), {}), (x["name"], {"jc": "left"}), (x["regno"], {}),
                     (won(x["amt"]), {"jc": "right"}), (won(x["cap"]), {"jc": "right"})])
    rows.append([(f"Tranche A 소계 ({len(TA)}개사)", {"span": 4, "b": True}),
                 (won(sum(x['amt'] for x in TA)), {"jc": "right", "b": True}), (won(CAP_A), {"jc": "right", "b": True})])
    for i, x in enumerate(TB):
        rows.append([("(공동)\n제3순위\n(Tranche B)", {"vm": "restart"}) if i == 0 else ("", {"vm": "cont"}),
                     (str(i + 1), {}), (x["name"], {"jc": "left"}), (x["regno"], {}),
                     (won(x["amt"]), {"jc": "right"}), (won(x["cap"]), {"jc": "right"})])
    rows.append([(f"Tranche B 소계 ({len(TB)}개사)", {"span": 4, "b": True}),
                 (won(sum(x['amt'] for x in TB)), {"jc": "right", "b": True}), (won(CAP_B), {"jc": "right", "b": True})])
    rows.append([("합   계", {"span": 4, "b": True}), ("", {}),
                 (won(COUPANG_CAP + CAP_A + CAP_B), {"jc": "right", "b": True})])
    return rows


ROSTER_W = [1100, 500, 2726, 1500, 1600, 1600]   # 합 9,026 = 본문 폭(A4, 좌우 여백 1,440)
NOTICE_W = [800, 1500, 2300, 1000, 1300, 2126]

ROSTER_NOTE = ("[우선수익자 목록 기재] 담보대출약정서 v27(BKL, INDEX #82) 별지 '대주목록 및 대주별 대출약정금' 기준입니다. "
               f"Tranche A {len(TA)}개사 1,600억원(Tranche A-2 변동: 우리은행 700억원 / Tranche A-1 고정: 롯데캐피탈 300·우리카드 100·신용협동조합중앙회 300·수협은행 100·에스비아이저축은행 100), "
               f"Tranche B {len(TB)}개사 380억원(키움캐피탈 200·엠지캐피탈 100·디비저축은행 80). 우선수익한도금액 = 각 대출약정금의 120%"
               "(Tranche A 1,920억원, Tranche B 456억원, 1순위 포함 총 금239,732,586,000원 = 별지 2 제4항 총액과 일치). "
               "※ v04는 대주통지처 엑셀(#81)대로 신협 조합 17곳을 따로 적었으나, 대출약정서는 신용협동조합중앙회를 단일 대주(300억원)로 두므로 약정서에 맞췄습니다(조합별 배분은 중앙회 내부 사항 [추정]). "
               "법인등록번호: 우리은행(기존 기재값), 롯데캐피탈·디비저축은행·수협은행(10/13 접수 사업자등록증), 신협중앙회(#81). "
               f"미기재 [*]: {', '.join(MISSING)}, 쿠팡 주식회사 — 법인등기사항증명서 또는 사업자등록증 수령 후 기재합니다.")

# ================================================================ 1) 본계약 별첨 1 (표 2 정정 + 세부 목록 표 추가)
t2 = tables[2]
track_replace(t2.cell(1, 1).paragraphs[0], "외 Tranche A 대주 [확정 시 기재]", f"외 {N_OTHERS}개사")
db_p = next(p for p in t2.cell(3, 0).paragraphs if p.text.startswith("DB저축은행"))
track_replace(db_p, "DB저축은행 [정식 상호 확인 필요]", "주식회사 디비저축은행")
cap1 = ins_para(t2._tbl, "※ 우선수익자별 세부 내역", b=True)
tbl_r1 = make_table(ROSTER_W, roster_rows(), sz=18)
cap1.addnext(tbl_r1)
comment(doc, cap1, ROSTER_NOTE)

# ================================================================ 2) 별지 2 수익자 표(표 3): Tranche A 22개사·Tranche B 상호
t3 = tables[3]
rows3 = t3._tbl.findall(qn("w:tr"))
tmpl_a, tmpl_b = copy.deepcopy(rows3[11]), copy.deepcopy(rows3[12])   # 상호 / 법인등록번호 한 쌍
rpr3 = direct_runs(t3.cell(13, 2).paragraphs[0])[0]                   # '키움캐피탈 주식회사' 글자 서식
for i, x in enumerate(TA[:5]):
    r_name, r_reg = 3 + 2 * i, 4 + 2 * i
    track_set(t3.cell(r_name, 2).paragraphs[0], x["name"], rpr3)
    track_set(t3.cell(r_reg, 2).paragraphs[0], x["regno"], rpr3)
    track_set(t3.cell(r_name, 3).paragraphs[0], f"금{won(x['cap'])}원", rpr3)
anchor = rows3[12]
for x in TA[5:]:
    ra = mark_row_inserted(copy.deepcopy(tmpl_a), [None, "상 호", x["name"], f"금{won(x['cap'])}원"], rpr3)
    rb = mark_row_inserted(copy.deepcopy(tmpl_b), [None, "법인등록번호", x["regno"], None], rpr3)
    anchor.addnext(ra)
    ra.addnext(rb)
    anchor = rb
track_replace(t3.cell(len(t3.rows) - 2, 2).paragraphs[0], "DB저축은행 [정식 상호 확인 필요]", "주식회사 디비저축은행")
nr3 = len(t3.rows)
for i, x in enumerate(TB):                       # Tranche B 법인등록번호(접수분만)
    if x["regno"] != "[*]":
        track_set(t3.cell(nr3 - 2 * len(TB) + 2 * i + 1, 2).paragraphs[0], x["regno"], rpr3)
comment(doc, t3.cell(3, 2).paragraphs[0]._p,
        f"[별지 2 수익자 기재] (공동) 제2순위 Tranche A 대주 {len(TA)}개사를 상호·법인등록번호·수익권증서 발행금액(대출약정금의 120%)으로 채웠습니다(기존 5칸 + 1쌍 행 추가). "
        "[발행금액 검증] 제2순위 84,000,000,000 + 36,000,000,000 + 12,000,000,000 + 36,000,000,000 + 12,000,000,000 + 12,000,000,000 = 192,000,000,000원, "
        "제3순위 24,000,000,000 + 12,000,000,000 + 9,600,000,000 = 45,600,000,000원, 제1순위 2,132,586,000원 → 총 239,732,586,000원으로 "
        "아래 제4항 '금 이천삼백구십칠억삼천이백오십팔만육천원정(239,732,586,000)' 및 순위별 내역과 일치합니다. 각 금액은 대출약정서 v27 별지의 대주별 대출약정금 × 120%입니다. "
        "법인등록번호 [*]는 서류 수령 후 기재합니다.")

# ================================================================ 3) 별지 4 여신거래 표(표 6)
t6 = tables[6]
rows6 = t6._tbl.findall(qn("w:tr"))
tmpl6 = copy.deepcopy(rows6[5])
rpr6 = direct_runs(t6.cell(2, 2).paragraphs[0])[0]
x0 = TA[0]
track_replace(t6.cell(1, 3).paragraphs[0], "Tranche A 대출", x0["tranche"])
track_set(t6.cell(1, 4).paragraphs[0], f"금{won(x0['amt'])}원", rpr6)
for i, x in enumerate(TA[1:5]):
    r = 2 + i
    for c, v in ((2, x["name"]), (3, x["tranche"]), (4, f"금{won(x['amt'])}원"), (5, "인출일"), (6, "인출일로부터 24개월")):
        track_set(t6.cell(r, c).paragraphs[0], v, rpr6)
anchor = rows6[5]
for x in TA[5:]:
    nr = mark_row_inserted(copy.deepcopy(tmpl6), [None, None, x["name"], x["tranche"], f"금{won(x['amt'])}원", "인출일", "인출일로부터 24개월"], rpr6)
    anchor.addnext(nr)
    anchor = nr
t6r = len(t6.rows)
track_replace(t6.cell(t6r - 2, 2).paragraphs[0], "DB저축은행 [정식 상호 확인 필요]", "주식회사 디비저축은행")
comment(doc, t6.cell(1, 2).paragraphs[0]._p,
        f"[별지 4 여신거래 기재] 대출약정서 v27 별지대로 Tranche A {len(TA)}개사(우리은행 Tranche A-2 변동 700억원, 나머지 {N_OTHERS}개사 Tranche A-1 고정, 합계 1,600억원)와 "
        f"Tranche B {len(TB)}개사(380억원)를 대주별로 기재했습니다(1행 추가). 합계 금198,000,000,000원은 기존 합계행과 일치합니다. 신규일·만기일은 대출약정서 기준(인출일, 인출일로부터 24개월)입니다.")

# ================================================================ 4) 특약 별첨 1 (목록 문단 정정 + 세부 목록 표)
ps = doc.paragraphs
k_b1 = find(ps, find(ps, 0, "부동산담보신탁계약에 대한 특약"), "2. 공동 2순위 우선수익자(Tranche A)")
track_replace(ps[k_b1], "외 Tranche A 대주 [확정 시 기재]", f"외 {N_OTHERS}개사")
track_replace(ps[k_b1 + 1], "DB저축은행[정식 상호 확인 필요]", "주식회사 디비저축은행")
cap2 = ins_para(ps[k_b1 + 1]._p, "※ 우선수익자별 세부 내역", b=True)
tbl_r2 = make_table(ROSTER_W, roster_rows(), sz=18)
cap2.addnext(tbl_r2)
comment(doc, cap2, "[우선수익자 목록 기재] 본계약 별첨 1과 같은 표입니다. " + ROSTER_NOTE.split("] ", 1)[1])

# ================================================================ 5) 특약 별첨 2 통지처 목록
ps = doc.paragraphs
k_n = find(ps, k_b1, "우선수익자들의 통지처 목록")
k_star = find(ps, k_n, "[*]")
assert ps[k_star].text.strip() == "[*]"
track_replace(ps[k_star], "[*]", "")
nrows = [[("순위", {}), ("우선수익자", {}), ("주소", {}), ("담당자", {}), ("전화 / 팩스", {}), ("전자우편", {})]]
nrows.append([("제1순위", {}), ("쿠팡 주식회사", {"jc": "left"}), ("[*]", {}), ("[*]", {}), ("[*]", {}), ("[*]", {})])


def tel_fax(x):
    t = x["tel"]
    return t + (f"\n(F) {x['fax']}" if x["fax"] else "")


for grp, lab in ((TA, "(공동)\n제2순위\n(Tranche A)"), (TB, "(공동)\n제3순위\n(Tranche B)")):
    for i, x in enumerate(grp):
        nrows.append([(lab, {"vm": "restart"}) if i == 0 else ("", {"vm": "cont"}),
                      (x["name"], {"jc": "left"}), (x["addr"], {"jc": "left"}), (x["person"], {}),
                      (tel_fax(x), {}), (x["email"], {"jc": "left"})])
tbl_n = make_table(NOTICE_W, nrows, sz=16)
ps[k_star]._p.addnext(tbl_n)
comment(doc, ps[k_n]._p,
        f"[통지처 목록 기재] 담보대출약정서 v27 부록 5 '대주 통지처'(9개사)를 옮겼습니다(Tranche A {len(TA)}개사·Tranche B {len(TB)}개사). "
        "① 담당자 휴대전화 번호는 넣지 않았습니다. 특약·별첨이 신탁원부로 등기소에 제출·열람될 수 있기 때문이며, 휴대전화는 대출약정서 통지처에서 관리합니다. 필요하면 추가하겠습니다. "
        "② 1순위 쿠팡 주식회사 통지처는 [미접수]이므로 [*]로 두었습니다(임대차계약서 통지 조항 또는 쿠팡 회신으로 확인 필요). "
        "③ 특약 제19조 통지처 표의 '별첨 2.와 같음'과 연결됩니다.")

# ================================================================ 5-1) 신탁사(우리은행 신탁부) 검토 의견 반영
TRUST = "[신탁사 의견 반영] "
ps = doc.paragraphs

# (1) 등기용 약관 제33조: 2부 작성·각 1부 보관·수탁자 보관용 등기소 제출
k33 = find(ps, 0, "제33조(계약서의 보관)")
track_replace(ps[k33], "1부", "2부", ctx="신탁계약서 1부를")
track_replace(ps[k33], "원본은 위탁자가 보관한다.", "원본은 위탁자 및 은행이 각각 1부씩 보관하기로 하며, 은행 보관용을 등기소에 제출하기로 한다.")
comment(doc, ps[k33]._p,
        TRUST + "신탁사 요청 문안대로 '신탁계약서 2부를 작성하여 … 원본은 위탁자 및 수탁자가 각각 1부씩 보관하기로 하며, 수탁자 보관용을 등기소에 제출'로 고쳤습니다. "
        "본계약(약관)은 수탁자를 '은행'으로 정의하므로 '수탁자' 대신 '은행'으로 적었고, '각자'는 v03의 '위탁자와 은행이 각각'(당사자 한정)을 유지했습니다. "
        "신탁사 표준 문구 그대로('수탁자', '각자')를 원하면 바꾸겠습니다. 날인 원본은 총 2부입니다.")

# (2) 등기용 특약 말미: 같은 취지 — v03에서 넣은 ' 각각 기명날인한 후'는 신탁사 문안으로 대체(자기 삽입분이므로 삭제 표시 없이 제거)
k_tk2 = find(ps, 0, "부동산담보신탁계약에 대한 특약")
k22 = find(ps, k_tk2, "본 특약이 체결된 증거로써")
p22 = ps[k22]
for ins in p22._p.findall(qn("w:ins")):
    if ins.get(qn("w:author")) == AUTHOR and "".join(t.text or "" for t in ins.iter(qn("w:t"))) == " 각각 기명날인한 후":
        p22._p.remove(ins)
track_replace(p22, "작성하여", " 각자 서명 또는 기명날인하도록 한 후,", ctx="1부를 작성하여", after=True)
track_replace(p22, "1부", "2부", ctx="본 특약 1부를")
track_replace(p22, "위탁자가 원본을 보관하기로 한다.", "원본은 위탁자 및 수탁자가 각각 1부씩 보관하기로 하며, 수탁자 보관용을 등기소에 제출하기로 한다.")
comment(doc, p22._p,
        TRUST + "신탁사 요청 문안대로 '본 특약 2부를 작성하여 각자 서명 또는 기명날인하도록 한 후, 원본은 위탁자 및 수탁자가 각각 1부씩 보관하기로 하며, 수탁자 보관용을 등기소에 제출'로 고쳤습니다. "
        "주어는 v03의 '위탁자와 수탁자는'(당사자 한정)을 유지했고, v03에서 넣었던 '각각 기명날인한 후'는 신탁사 문안으로 대체했습니다. 날인 원본은 총 2부입니다.")

# (3) 특약 제12조 제1항 제2호 및 단서 — v03에서 이미 삭제(신탁사 의견과 같음)
k12 = find(ps, k_tk2, "제12조【신탁부동산의 처분(담보권의 실행)】")
comment(doc, ps[k12]._p,
        TRUST + "신탁사 의견(제2호에 따른 담보권 실행이 위탁자 사전동의로 사실상 불가능 → 단서 삭제 고려)은 v03에서 이미 반영했습니다: 제①항 단서와 제2호를 삭제하고, "
        "제16조 제③항의 '제1, 2호'를 '제1호'로 맞췄습니다. 처분사유는 대출약정서 제12조의 기한의 이익 상실(제1호)로 한정됩니다.")

# (4) 특약 제15조 제②항 ↔ 대출약정서 제11조 제6항 제(2)호 상충 여부 — 검토 의견(문언 변경 없음)
k15 = find(ps, k12, "신탁계약서 제10조 제②항, 제11조 및 본 특약 제6조에도 불구하고")
comment(doc, ps[k15]._p,
        TRUST + "[제15조 제②항 ↔ 대출약정서 제11조 제6항 제(2)호 검토] 정면 상충은 없다고 봅니다(내부 검토). "
        "① 대출약정 제11조⑥(2)는 약정 체결 후 새로 맺는 임대차·담보계약으로 본건 부동산에 '설정되는 담보권'(임차인을 위한 우선수익권 포함)이 대주 우선수익권보다 후순위여야 한다는 조항입니다"
        "(제8조⑦(1)의 임대차보증금계좌 근질권은 예외). "
        "② 특약 제15조②는 위탁자의 임대 권한과, 처분대금 정산 시 임대차보증금(쿠팡 1순위분 제외)의 정산 순서(제23조①3호 다음·5호 앞)를 정한 것입니다. 임차인에게 담보권이나 우선수익권을 설정하는 조항은 아닙니다. "
        "③ 다만 신규 임대차의 보증금도 정산 단계에서 대주 우선수익자보다 먼저 받게 되므로, 실질적으로는 제11조⑥(2)의 '후순위' 취지와 어긋나 보일 수 있습니다. "
        "대항력 있는 임차인의 보증금은 매수인이 인수하는 것이 실무라 경제적 차이는 크지 않습니다. "
        "필요하면 끝에 다음 문장을 넣는 방안이 있습니다: '본 항에 따른 임대차보증금의 정산 순서는 임차인을 위한 담보권 또는 우선수익권을 설정하는 것이 아니며, "
        "대출약정서 제11조 제6항 제(2)호에 따른 대주 우선수익자들의 권리에 영향을 미치지 아니한다.' "
        "→ BKL(대주 측) 및 신탁사 확인이 필요합니다(외부 법무법인 확인 필요).")

# (5) 특약 제19조 단서 → 대출약정서 제14조와 같은 수령간주 문언
k19 = find(ps, k15, "본 특약상 발송이 요구되거나 허용되는 모든 통지는")
track_replace(ps[k19], "다만, 팩시밀리 또는 전자우편으로 송부하는 경우에는 그 수령을 확인하는 내용을 팩시밀리 또는 전자우편으로 받아야 한다.",
              "모든 통지는 인편으로 교부될 경우에는 그 교부일에, 등기우편으로 발송하는 경우에는 그 도달일에, 전자우편으로 발송하는 경우에는 수신확인일에 각 수령된 것으로 본다.")
comment(doc, ps[k19]._p,
        TRUST + "단서를 대출약정서 v27 제14조 후단과 같은 문언('모든 통지는 인편 … 교부일에, 등기우편 … 도달일에, 전자우편 … 수신확인일에 각 수령된 것으로 본다')으로 바꿨습니다. "
        "본문에는 팩시밀리 발송이 있지만, 대출약정서와 같이 팩시밀리의 수령 시점은 정하지 않았습니다. 필요하면 '팩시밀리로 발송하는 경우에는 수신확인일에'를 추가할 수 있습니다.")

# ================================================================ 6) 서식 통일 — 특약 이후 본문 문단을 본계약 첫 페이지 기준으로
#   기준(본계약 제1면): 줄간격 1.5줄(line 360 auto)·문단 뒤 0, 격자 맞춤 해제, 글꼴 바탕체 / Times New Roman, 11pt
#   제목은 본계약 제목(20pt 굵게 밑줄), 별첨·별지 제목은 본계약 별첨 제목(15pt 굵게 밑줄)과 같게 한다. 표 안 서식은 손대지 않는다.
SUBTITLES = ("우선수익자들 목록", "우선수익자들의 통지처 목록", "임대차계약등 내역")


def set_spacing(p_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p_el.insert(0, ppr)
    para = Paragraph(p_el, doc._body)
    pf = para.paragraph_format
    pf.space_after = Pt(0)
    pf.line_spacing = 1.5
    sp = ppr.find(qn("w:spacing"))
    if qn("w:before") in sp.attrib:
        del sp.attrib[qn("w:before")]
    if ppr.find(qn("w:snapToGrid")) is None:
        sg = OxmlElement("w:snapToGrid")
        sg.set(qn("w:val"), "0")
        sp.addprevious(sg)
    else:
        ppr.find(qn("w:snapToGrid")).set(qn("w:val"), "0")


def set_rpr(rpr, sz, title=False):
    f = rpr.get_or_add_rFonts()
    for a in list(f.attrib):
        if a.endswith("Theme"):
            del f.attrib[a]
    for k in ("ascii", "hAnsi", "cs"):
        f.set(qn(f"w:{k}"), "Times New Roman")
    f.set(qn("w:eastAsia"), "바탕체")
    for tag in ("w:sz", "w:szCs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag)
            succ = [qn(x) for x in ("w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
                                    "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange")]
            if tag == "w:sz":
                succ.insert(0, qn("w:szCs"))
            nxt = next((c for c in rpr if c.tag in succ), None)
            (nxt.addprevious(e) if nxt is not None else rpr.append(e))
        e.set(qn("w:val"), str(sz))
    cs = rpr.find(qn("w:spacing"))
    if cs is not None:
        rpr.remove(cs)
    if title:
        col = rpr.get_or_add_color()
        col.set(qn("w:val"), "000000")
        for a in list(col.attrib):
            if a != qn("w:val"):
                del col.attrib[a]
        rpr.get_or_add_b()
        u = rpr.get_or_add_u()
        u.set(qn("w:val"), "single")


def normalize_para(p_el, sz, title=False):
    set_spacing(p_el)
    ppr = p_el.find(qn("w:pPr"))
    mrpr = ppr.find(qn("w:rPr"))
    if mrpr is None:
        mrpr = OxmlElement("w:rPr")
        ppr.append(mrpr)
    # 문단 기호 rPr은 CT_RPr가 아닐 수 있어 직접 처리
    f = mrpr.find(qn("w:rFonts"))
    if f is None:
        f = OxmlElement("w:rFonts")
        tr = [c for c in mrpr if c.tag in (qn("w:ins"), qn("w:del"))]
        (tr[-1].addnext(f) if tr else mrpr.insert(0, f))
    for a in list(f.attrib):
        if a.endswith("Theme"):
            del f.attrib[a]
    for k in ("ascii", "hAnsi", "cs"):
        f.set(qn(f"w:{k}"), "Times New Roman")
    f.set(qn("w:eastAsia"), "바탕체")
    for tag in ("w:sz", "w:szCs"):
        e = mrpr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag)
            mrpr.append(e)
        e.set(qn("w:val"), str(sz))
    for r in p_el.iter(qn("w:r")):
        set_rpr(r.get_or_add_rPr(), sz, title)


body = doc.element.body
started = False
n_norm = 0
for el in list(body):
    if el is tk_title:
        started = True
    if not started or el.tag != qn("w:p"):
        continue
    txt = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
    if el is tk_title:
        normalize_para(el, 40, title=True)
    elif txt in SUBTITLES:
        normalize_para(el, 30)
    else:
        normalize_para(el, 22)
    n_norm += 1

# 특약 구역 여백·격자를 본계약과 같게
sect = body.find(qn("w:sectPr"))
pm = sect.find(qn("w:pgMar"))
for k, v in (("top", 1701), ("right", 1440), ("bottom", 1440), ("left", 1440)):
    pm.set(qn(f"w:{k}"), str(v))
dg = sect.find(qn("w:docGrid"))
if dg is not None and qn("w:type") in dg.attrib:
    del dg.attrib[qn("w:type")]

# 본계약: 일부 글자에 섞인 대체 글꼴(바탕;Batang, MS Mincho)을 바탕체로
for el in body:
    if el is tk_title:
        break
    for f in el.iter(qn("w:rFonts")):
        if f.get(qn("w:eastAsia")) in ("바탕;Batang", "MS Mincho"):
            f.set(qn("w:eastAsia"), "바탕체")

comment(doc, tk_title,
        "[서식 통일] 특약·별첨·별지 본문을 본계약 첫 페이지 서식으로 맞췄습니다(내용 변경 없음, 변경추적 미표시): 줄간격 1.5줄(종전 일부 문단 '고정 10pt'로 12pt 글자가 겹쳐 보이던 문제 해소)·문단 뒤 0pt, "
        "글꼴 바탕체/Times New Roman 11pt(종전 바탕·바탕체 혼용 12pt, 일부 8pt), 제목 20pt·별첨 제목 15pt 굵게 밑줄, 여백 위 30mm·좌우하 25.4mm(종전 위 35mm·좌우하 30mm). "
        f"대상 문단 {n_norm}개. 표 안 서식은 그대로 두었습니다.")

doc.core_properties.last_modified_by = AUTHOR
doc.save(OUT)
print("saved", OUT, "| normalized paragraphs:", n_norm, "| TA", len(TA), "TB", len(TB))
