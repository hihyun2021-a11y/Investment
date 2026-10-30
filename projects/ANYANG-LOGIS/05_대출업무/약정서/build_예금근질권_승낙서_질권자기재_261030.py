# -*- coding: utf-8 -*-
"""예금채권근질권설정계약서(BKL 25746316.vfinal) 별지 3-2 승낙서 수신란에 근질권자 기재 + 승낙서만 별도 파일로 분리 — 운용역 요청.
원본(불변): 20261030_대출약정_예금채권근질권설정계약서_BKL_vfinal_원본.docx
출력: ① 20261030_대출약정_예금채권근질권설정계약서_BKL_vfinal_승낙서질권자기재.docx (계약서 전체)
      ② 20261030_대출약정_예금근질권설정승낙서_별지3-2_질권자기재.docx (승낙서만)
"""
import copy
import os
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261030_대출약정_예금채권근질권설정계약서_BKL_vfinal_원본.docx")
OUT_ALL = os.path.join(HERE, "20261030_대출약정_예금채권근질권설정계약서_BKL_vfinal_승낙서질권자기재.docx")
OUT_ONE = os.path.join(HERE, "20261030_대출약정_예금근질권설정승낙서_별지3-2_질권자기재.docx")


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def replace_span(p, pattern, new):
    """문단 안 정규식 일치 구간 1곳을 new로 바꾼다(시작 런 서식 유지)."""
    ts = list(p.iter(qn("w:t")))
    full = "".join(t.text or "" for t in ts)
    m = re.search(pattern, full)
    assert m, (pattern, full[:80])
    s, e = m.span()
    pos, placed = 0, False
    for t in ts:
        txt = t.text or ""
        a, b = pos, pos + len(txt)
        pos = b
        if b <= s or a >= e:
            continue
        lo, hi = max(s, a) - a, min(e, b) - a
        t.text = (txt[:lo] + new + txt[hi:]) if not placed else (txt[:lo] + txt[hi:])
        placed = True
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


doc = Document(SRC)
body = doc.element.body

# 별지 1. 근질권자 목록 추출
groups, rank = {"1순위": [], "2순위": []}, None
for tbl in body.iter(qn("w:tbl")):
    rows = [[ptext(tc).strip() for tc in tr.findall(qn("w:tc"))] for tr in tbl.findall(qn("w:tr"))]
    if rows and rows[0][:2] == ["순위", "근질권자"]:
        for r in rows[1:]:
            if r[0]:
                rank = r[0]
            if r[1] and "합계" not in r[1] and rank in groups:
                groups[rank].append(r[1])
        break
assert len(groups["1순위"]) == 22 and len(groups["2순위"]) == 3, groups

# 별지 3-2 수신란
paras = list(body.iter(qn("w:p")))
rp = [p for p in paras if "[*근질권자 기재]" in ptext(p)]
assert len(rp) == 1
rp = rp[0]
line1 = "(1순위 근질권자) " + ", ".join(groups["1순위"])
line2 = "(2순위 근질권자) " + ", ".join(groups["2순위"])
replace_span(rp, r"\[\*근질권자 기재\]", line1)
run = [t for t in rp.iter(qn("w:t")) if line1 in (t.text or "")][0].getparent()
r2 = copy.deepcopy(run)
for c in list(r2):
    if c.tag != qn("w:rPr"):
        r2.remove(c)
r2.append(OxmlElement("w:br"))
t2 = OxmlElement("w:t")
t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
t2.text = line2
r2.append(t2)
run.addnext(r2)
ppr = rp.find(qn("w:pPr"))
if ppr is None:
    ppr = OxmlElement("w:pPr")
    rp.insert(0, ppr)
jc = ppr.find(qn("w:jc"))
if jc is None:
    jc = OxmlElement("w:jc")
    ppr.append(jc)
jc.set(qn("w:val"), "left")
doc.save(OUT_ALL)
print("수신:", ptext(rp))

# 승낙서만 분리: '별지 3-2.' 문단 앞의 본문 요소 삭제(구역 설정·머리글은 유지)
doc = Document(OUT_ALL)
body = doc.element.body
els = list(body.iterchildren())
start = next(i for i, e in enumerate(els) if e.tag == qn("w:p") and ptext(e).strip() == "별지 3-2.")
for e in els[:start]:
    body.remove(e)
first = list(body.iterchildren())[0]
for br in first.iter(qn("w:br")):  # 첫 쪽 앞 페이지 나누기 제거
    if br.get(qn("w:type")) == "page":
        br.getparent().remove(br)
doc.save(OUT_ONE)
print("saved", OUT_ALL, OUT_ONE)
