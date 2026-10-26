# -*- coding: utf-8 -*-
"""보험근질권설정계약서(BKL 25746312.vfinal 수정2, 2026-10-26) 보완 — 운용역 요청.
1) 서식 가.(통지서) 본문의 대출약정·보험근질권설정계약 체결일 공란 → 2026년 10월 27일 (본문 첫머리·전문과 같게)
2) 서식 나.(승낙서) 수신란 "[*근질권자 기재]" → 별지 1. 근질권자 25개사 명칭(1순위·2순위) 기재
원본(불변): 20261026_대출약정_보험근질권설정계약서_BKL_vfinal_수정2_원본.docx
출력: 20261026_대출약정_보험근질권설정계약서_BKL_vfinal_체결일·서식나보완.docx (변경추적 없음)
"""
import copy
import os
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261026_대출약정_보험근질권설정계약서_BKL_vfinal_수정2_원본.docx")
OUT = os.path.join(HERE, "20261026_대출약정_보험근질권설정계약서_BKL_vfinal_체결일·서식나보완.docx")


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def replace_span(p, pattern, new):
    """문단 안 정규식 일치 구간 1곳을 new로 바꾼다. 시작 런의 서식을 유지하고, 걸친 런에서는 해당 글자만 지운다."""
    ts = [t for t in p.iter(qn("w:t"))]
    full = "".join(t.text or "" for t in ts)
    m = re.search(pattern, full)
    assert m, (pattern, full[:80])
    s, e = m.span()
    pos = 0
    placed = False
    for t in ts:
        txt = t.text or ""
        a, b = pos, pos + len(txt)
        pos = b
        if b <= s or a >= e:
            continue
        lo, hi = max(s, a) - a, min(e, b) - a
        if not placed:
            t.text = txt[:lo] + new + txt[hi:]
            placed = True
        else:
            t.text = txt[:lo] + txt[hi:]
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return m.group(0)


doc = Document(SRC)
body = doc.element.body
paras = list(body.iter(qn("w:p")))

# 별지 1. 근질권자 목록(같은 문서) 추출
first_tbl = None
for tbl in body.iter(qn("w:tbl")):
    rows = [[ptext(tc) for tc in tr.findall(qn("w:tc"))] for tr in tbl.findall(qn("w:tr"))]
    if rows and rows[0][:2] == ["순위", "근질권자"]:
        first_tbl = rows
        break
rank, groups = None, {"1순위": [], "2순위": []}
for r in first_tbl[1:]:
    if r[0].strip():
        rank = r[0].strip()
    if r[1].strip() in ("총 합계", ""):
        continue
    groups[rank].append(r[1].strip())
assert len(groups["1순위"]) == 22 and len(groups["2순위"]) == 3, groups

# 1) 서식 가. 체결일
target = next(p for p in paras if ptext(p).startswith("당사는 2026년 10월") and "보험근질권설정계약(추후 수정" in ptext(p))
replace_span(target, r"2026년 10월 \[\s*\]일(?= 대출약정에)", "2026년 10월 27일")
replace_span(target, r"2026년 10월 \[\s*\]일(?= 근질권자와 보험근질권설정계약)", "2026년 10월 27일")
assert "[" not in ptext(target), ptext(target)

# 2) 서식 나. 수신란
recv = [p for p in paras if ptext(p).strip().startswith("수신:") and "[*근질권자 기재]" in ptext(p)]
assert len(recv) == 1, len(recv)
rp = recv[0]
line1 = "(1순위 근질권자) " + ", ".join(groups["1순위"])
line2 = "(2순위 근질권자) " + ", ".join(groups["2순위"])
replace_span(rp, r"\[\*근질권자 기재\]", line1)
# 2순위는 줄바꿈 후 같은 서식으로 이어 쓴다
last_t = [t for t in rp.iter(qn("w:t")) if line1 in (t.text or "")][0]
run = last_t.getparent()
br_run = copy.deepcopy(run)
for c in list(br_run):
    if c.tag != qn("w:rPr"):
        br_run.remove(c)
br_run.append(OxmlElement("w:br"))
t2 = OxmlElement("w:t")
t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
t2.text = line2
br_run.append(t2)
run.addnext(br_run)

# 이름 나열 문단은 양쪽 정렬 시 간격이 벌어지므로 왼쪽 정렬
ppr = rp.find(qn("w:pPr"))
if ppr is None:
    ppr = OxmlElement("w:pPr")
    rp.insert(0, ppr)
jc = ppr.find(qn("w:jc"))
if jc is None:
    jc = OxmlElement("w:jc")
    ppr.append(jc)
jc.set(qn("w:val"), "left")

doc.save(OUT)
print("saved", OUT)
print("서식 가:", ptext(target)[:160])
print("서식 나 수신:", ptext(rp))
