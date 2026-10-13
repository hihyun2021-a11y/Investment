# -*- coding: utf-8 -*-
"""부동산담보신탁계약서 및 특약 v03 — 당사자를 위탁자·수탁자로 한정 + 특약 제12조 제1항 제2호 삭제.

입력(원본 불변): 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v02통합Clean_법무팀검토_WJL2-1.docx
출력: 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v03_당사자정리_변경추적.docx
- 모든 변경은 Word 변경추적(작성자 "코람코 상장리츠팀") + 메모로 표시. 법무팀(이원준 차장) 기존 변경추적·메모는 그대로 둔다.
실행: python3 build_담보신탁_v03_당사자정리_261013.py
"""
import copy
import os
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v02통합Clean_법무팀검토_WJL2-1.docx")
OUT = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v03_당사자정리_변경추적.docx")
if len(sys.argv) > 1:
    OUT = sys.argv[1]

AUTHOR, INIT, DATE = "코람코 상장리츠팀", "KR", "2026-10-13T09:00:00Z"
_id = [9000]


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
    """run 하나를 삭제 표시로 바꾼다(w:t → w:delText)."""
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
    """run을 pos 위치에서 둘로 나눈다. 앞 run을 반환(뒤 run은 바로 다음 형제)."""
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
    """문단의 직접 run들에서 old를 찾아 삭제 표시하고 그 자리에 new를 삽입 표시한다.
    ctx: old를 찾을 문맥(ctx 안의 old 위치만 변경). after=True면 old는 그대로 두고 그 뒤에 new를 삽입."""
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
    # 경계에서 run 분할
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


def track_delete_paragraph(p):
    for c in list(p._p):
        if c.tag == qn("w:r"):
            to_del(c)
        elif c.tag == qn("w:ins"):
            for r in list(c.findall(qn("w:r"))):
                d = mark("w:del")
                r.addprevious(d)
                for t in r.findall(qn("w:t")):
                    dt = OxmlElement("w:delText")
                    dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    dt.text = t.text
                    t.addprevious(dt)
                    r.remove(t)
                d.append(r)
    ppr = p._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    rpr.append(mark("w:del"))


def track_insert_paragraph_after(ref_el, text, ppr_src_p, rpr_src_run):
    np_ = OxmlElement("w:p")
    if ppr_src_p._p.pPr is not None:
        ppr = copy.deepcopy(ppr_src_p._p.pPr)
        for old in ppr.findall(qn("w:rPr")):
            ppr.remove(old)
    else:
        ppr = OxmlElement("w:pPr")
    rpr = OxmlElement("w:rPr")
    rpr.append(mark("w:ins"))
    ppr.append(rpr)
    np_.append(ppr)
    ins = mark("w:ins")
    ins.append(new_run(text, rpr_src_run))
    np_.append(ins)
    ref_el.addnext(np_)
    return np_


def comment(doc, p_el, text):
    """문단 전체에 메모를 단다(앞뒤 빈 run을 기준점으로 사용)."""
    from docx.text.paragraph import Paragraph
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


doc = Document(SRC)
ps = doc.paragraphs

# ================================================================ 본계약
# (1) 제33조 계약서 보관
k33 = find(ps, 0, "제33조(계약서의 보관)")
track_replace(ps[k33], "각자", "위탁자와 은행이 각각", ctx="각자 서명 또는 기명날인하도록")
comment(doc, ps[k33]._p,
        "[당사자 정리] 신탁계약의 당사자를 위탁자·수탁자(은행)로 한정합니다. 신탁계약서·특약에 우선수익자가 당사자로 기재되면 "
        "신탁등기 시 등기관이 우선수익자 전원의 간인을 요구할 수 있다는 법무사 의견에 따른 변경입니다. 대주 우선수익자가 25개사(Tranche A 22개사(신협 조합 17곳 포함)·Tranche B 3개사, 통지처 목록 261012 vf 기준)라 "
        "전원 간인은 현실적으로 어렵습니다. 우선수익자는 별첨 1의 지정과 수익권증서 발급으로 지위를 취득합니다(특약 제1조의2 신설).")

# (2) 본계약 우선수익자 서명란 삭제
k1 = find(ps, k33, "제 1 순 위")
k_end = find(ps, k1, "대표이사 [*] (인)")  # (공동) 2·3순위 블록 끝
first_del = ps[k1]._p
for k in range(k1, k_end + 1):
    track_delete_paragraph(ps[k])
comment(doc, first_del,
        "[당사자 정리] 본계약 서명란에서 1순위(쿠팡) 및 공동 2·3순위 우선수익자 기명날인란을 삭제합니다. 신탁계약서는 위탁자·수탁자만 기명날인합니다. "
        "우선수익자 명단·순위·한도금액은 별첨 1(우선수익자들 목록)에 그대로 둡니다.")

# ================================================================ 특약
kp = find(ps, k_end, "주식회사 코람코라이프로지스위탁관리부동산투자회사(이하 “위탁자”라 함), 주식회사 우리은행")
pre = ps[kp]
track_replace(pre, ", ", "와 ", ctx="(이하 “위탁자”라 함), 주식회사 우리은행")
track_replace(pre, ", 쿠팡 주식회사(이하 “1순위 우선수익자”라 함), 별첨 1.에 공동 2순위 우선수익자로 기재된 회사들(이하 총칭하여 “2순위 우선수익자”라 함), "
                   "별첨 1.에 공동 3순위 우선수익자로 기재된 회사들(이하 총칭하여 “3순위 우선수익자”라 하고, 2순위 우선수익자 및 3순위 우선수익자를 총칭하여 "
                   "“대주 우선수익자들”이라 하며, 1순위 우선수익자, 2순위 우선수익자 및 3순위 우선수익자를 개별적으로 “우선수익자”, 총칭하여 “우선수익자들”이라 함)", "")
track_replace(pre, "“본 신탁계약”이라 함).", after=True,
              new=" 본 신탁계약에서 쿠팡 주식회사를 “1순위 우선수익자”라 하고, 별첨 1.에 공동 2순위 우선수익자로 기재된 회사들을 총칭하여 "
              "“2순위 우선수익자”, 별첨 1.에 공동 3순위 우선수익자로 기재된 회사들을 총칭하여 “3순위 우선수익자”라 하며, 2순위 우선수익자 및 3순위 우선수익자를 "
              "총칭하여 “대주 우선수익자들”, 1순위 우선수익자, 2순위 우선수익자 및 3순위 우선수익자를 개별적으로 “우선수익자”, 총칭하여 “우선수익자들”이라 한다.")
comment(doc, pre._p,
        "[당사자 정리] 특약 전문의 약정 당사자를 위탁자·수탁자로 한정했습니다. 우선수익자 관련 정의(1순위·2순위·3순위·대주 우선수익자들·우선수익자들)는 "
        "당사자 열거에서 빼고 정의 문장으로 옮겨 내용은 그대로 유지합니다.")

# 제1조의2 신설 (제1조 본문 다음)
k1h = find(ps, kp, "제1조【특약의 목적】")
k1b = k1h + 1
head_src, body_src = ps[k1h], ps[k1b]
h_run = direct_runs(head_src)[0]
b_run = direct_runs(body_src)[0]
texts = [
    ("h", "제1조의2【우선수익자의 지위】"),
    ("b", "① 우선수익자들은 본 신탁계약의 당사자가 아니며, 위탁자의 지정에 따라 별첨 1.에 기재된 순위 및 우선수익한도금액으로 본 신탁계약상 우선수익권을 취득한다. "
          "수탁자는 신탁등기 시 별첨 1.에 따라 우선수익자들을 신탁원부에 기재하고, 본 특약 제4조에 따라 수익권증서를 발급한다."),
    ("b", "② 신탁계약서 및 본 특약 중 우선수익자들의 권리에 관한 규정은 우선수익자들을 위하여 그 효력이 있다. 신탁계약서 및 본 특약 중 우선수익자들의 확인, 동의, "
          "승낙 또는 이의 불제기에 관한 규정은 우선수익자가 수익권증서를 수령한 때(대주 우선수익자들의 경우 대출약정서 체결 시를 포함한다) 해당 우선수익자가 이를 승낙한 것으로 본다."),
    ("b", "③ 위탁자와 수탁자는 우선수익자들(해당 변경이 특정 우선수익자의 권리에만 영향을 미치는 경우에는 해당 우선수익자)의 사전 서면 동의 없이 "
          "신탁계약서 및 본 특약을 변경하거나 해지할 수 없다."),
]
ref = body_src._p
new_head = None
for kind, t in texts:
    if kind == "h":
        ref = track_insert_paragraph_after(ref, t, head_src, h_run)
        new_head = ref
    else:
        ref = track_insert_paragraph_after(ref, t, body_src, b_run)
comment(doc, new_head,
        "[신설] 우선수익자가 당사자에서 빠지더라도 권리 보호와 기존 조항의 효력을 유지하기 위한 조항입니다. "
        "① 우선수익권 취득 근거(위탁자 지정·별첨 1·수익권증서) "
        "② 특약 중 우선수익자들의 확인·이의 불제기 조항(제6조 제④항, 제6조의2, 제14조, 제20조 등)이 수익권증서 수령·대출약정 체결로 승낙된 것으로 봄 "
        "③ 우선수익자 동의 없는 신탁계약 변경·해지 금지(대주 보호). "
        "※ 문언은 BKL(대주 측) 및 우리은행 신탁부 확인이 필요합니다(내부 검토안).")

# 제12조 제1항 단서 삭제 + 제2호 삭제
k12 = find(ps, kp, "제12조【신탁부동산의 처분(담보권의 실행)】")
p12 = ps[k12 + 1]
track_replace(p12, " 단, 본 항 제2호의 경우에는 위탁자의 사전 동의가 있어야 한다.", "")
comment(doc, p12._p,
        "[변경 요청 반영] 제2호(위탁자의 본 신탁계약 위반 시 처분요청) 삭제에 따라 '단, 본 항 제2호의 경우에는 위탁자의 사전 동의가 있어야 한다' 단서도 삭제합니다. "
        "처분사유는 대출약정서상 기한의 이익 상실(제1호)로 한정되며, 대출약정 제12조(기한이익상실)와 일치합니다.")
p12_2 = ps[k12 + 3]
assert "위탁자가 본 신탁계약" in p12_2.text
track_delete_paragraph(p12_2)
comment(doc, p12_2._p,
        "[변경 요청 반영] 제2호 '위탁자가 본 신탁계약의 중요한 사항을 위반한 경우'를 삭제합니다(법무팀 10/08 '중요한 사항' 수정분 포함). "
        "신탁계약 위반은 대출약정상 의무 위반·치유기간(30영업일) 미치유 시 기한이익상실 사유로 포섭되므로, 신탁계약만으로 별도 처분사유를 두지 않습니다.")

# 제16조 제3항 '제1, 2호' → '제1호'
k16 = find(ps, k12, "본 특약 제12조 제①항 제1, 2호에 의한")
track_replace(ps[k16], ", 2", "", ctx="제1, 2호")
comment(doc, ps[k16]._p, "[연동 정리] 제12조 제1항 제2호 삭제에 따라 처분보수 근거 조항을 '제1호'로 맞춥니다.")

# 제18조 간인대체: 당사자 = 위탁자·수탁자
k18 = find(ps, k16, "신탁계약서 및 본 특약의 당사자들은 신탁계약서 및 본 특약의 간인을")
track_replace(ps[k18], "신탁계약서 및 본 특약의 당사자들은", "위탁자와 수탁자는")
track_replace(ps[k18], "당사자들이", "위탁자와 수탁자가", ctx="당사자들이 합의하여 지정하는 자")
comment(doc, ps[k18]._p,
        "[당사자 정리] 간인은 위탁자·수탁자 간 합의로 지정하는 자의 직인 또는 천공으로 대체합니다. 우선수익자 간인은 필요 없다는 점을 분명히 하는 변경입니다. "
        "실제 등기 실무(간인 요구 여부)는 법무사·등기소 확인이 필요합니다.")

# 특약 말미 보관 문구
k22 = find(ps, k18, "본 특약이 체결된 증거로써 당사자들은")
track_replace(ps[k22], "당사자들은", "위탁자와 수탁자는", ctx="당사자들은 본 특약 1부를")
track_replace(ps[k22], "작성하여", " 각각 기명날인한 후", ctx="1부를 작성하여", after=True)
comment(doc, ps[k22]._p, "[당사자 정리] 특약 기명날인 당사자를 위탁자·수탁자로 명시합니다.")

# 특약 우선수익자 서명란 삭제
ks = find(ps, k22, "1 순 위 우선수익자")
ke = find(ps, ks, "대표자 [*]")
ke = find(ps, ke + 1, "대표자 [*]")
first = ps[ks]._p
for k in range(ks, ke + 1):
    track_delete_paragraph(ps[k])
comment(doc, first, "[당사자 정리] 특약 서명란에서 1순위(쿠팡) 및 공동 2·3순위 우선수익자 기명날인란을 삭제합니다. 위탁자·수탁자만 기명날인합니다.")

doc.core_properties.last_modified_by = AUTHOR
doc.save(OUT)
print("saved", OUT)
