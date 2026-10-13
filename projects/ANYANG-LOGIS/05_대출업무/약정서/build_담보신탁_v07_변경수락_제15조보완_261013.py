# -*- coding: utf-8 -*-
"""부동산담보신탁계약서 및 특약 v07 — v06의 변경추적을 모두 수락 + 특약 제15조 제②항 끝에 보완 문장 추가(변경추적·메모).

입력(원본 불변): 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v06_신탁사의견반영_대주22개사_변경추적.docx
출력: 05_대출업무/약정서/20261013_대출약정_부동산담보신탁계약서및특약_v07_변경수락_제15조보완.docx

- 변경 수락(운용역 지시 2026-10-13): 삽입은 확정, 삭제는 제거, 삭제된 문단(서명란 등)은 문단째 제거. 법무팀·당사 변경 모두 대상.
  메모는 남긴다. 삭제 문단에 달린 메모는 바로 앞 문단으로 옮겨 설명이 사라지지 않게 한다.
- 새 변경은 제15조 제②항 보완 문장 1건뿐이다(변경추적으로 표시, 메모 갱신).
실행: python3 build_담보신탁_v07_변경수락_제15조보완_261013.py
"""
import copy
import os
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v06_신탁사의견반영_대주22개사_변경추적.docx")
OUT = os.path.join(HERE, "20261013_대출약정_부동산담보신탁계약서및특약_v07_변경수락_제15조보완.docx")
if len(sys.argv) > 1:
    OUT = sys.argv[1]

AUTHOR, DATE = "코람코 상장리츠팀", "2026-10-13T18:00:00Z"
_id = [20000]
COMMENT_TAGS = (qn("w:commentRangeStart"), qn("w:commentRangeEnd"))
ADDED = ("본 항에 따른 임대차보증금의 정산 순서는 임차인을 위한 담보권 또는 우선수익권을 설정하는 것이 아니며, "
         "대출약정서 제11조 제6항 제(2)호에 따른 대주 우선수익자들의 권리에 영향을 미치지 아니한다.")


def mark(tag):
    _id[0] += 1
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(_id[0]))
    el.set(qn("w:author"), AUTHOR)
    el.set(qn("w:date"), DATE)
    return el


def accept_all(root):
    """root(본문·머리글 등) 안의 삽입/삭제 변경을 모두 수락한다. 반환: 제거한 문단 수."""
    # 1) 글자 단위 삭제 제거(문단 기호·행 표시는 제외)
    for d in list(root.iter(qn("w:del"))):
        if d.getparent().tag not in (qn("w:rPr"), qn("w:trPr")):
            for c in list(d.iter()):                 # 삭제 구간 안의 메모 기준점은 살린다
                if c.tag in COMMENT_TAGS or (c.tag == qn("w:r") and c.find(qn("w:commentReference")) is not None):
                    d.addprevious(c)
            d.getparent().remove(d)
    # 2) 글자 단위 삽입은 감싼 표시만 벗긴다
    for ins in list(root.iter(qn("w:ins"))):
        par = ins.getparent()
        if par.tag in (qn("w:rPr"), qn("w:trPr")):
            par.remove(ins)                      # 삽입된 문단 기호·행 → 확정
            continue
        idx = par.index(ins)
        for c in list(ins):
            par.insert(idx, c)
            idx += 1
        par.remove(ins)
    # 3) 삭제된 행 제거
    for tr in list(root.iter(qn("w:tr"))):
        trpr = tr.find(qn("w:trPr"))
        if trpr is not None and trpr.find(qn("w:del")) is not None:
            tr.getparent().remove(tr)
    # 4) 문단 기호가 삭제된 문단: 남은 글자가 없으면 문단째 제거(메모 기준점은 앞 문단으로 이동)
    removed = 0
    for p in list(root.iter(qn("w:p"))):
        ppr = p.find(qn("w:pPr"))
        rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
        dm = rpr.find(qn("w:del")) if rpr is not None else None
        if dm is None:
            continue
        rpr.remove(dm)
        if "".join(t.text or "" for t in p.iter(qn("w:t"))).strip():
            # 남은 글자가 있으면 다음 문단과 합친다(Word의 문단 기호 삭제 수락과 같음)
            nxt = p.getnext()
            if nxt is not None and nxt.tag == qn("w:p"):
                npr = nxt.find(qn("w:pPr"))
                pos = 1 if npr is not None else 0
                for c in [c for c in p if c.tag != qn("w:pPr")]:
                    nxt.insert(pos, c)
                    pos += 1
                p.getparent().remove(p)
                removed += 1
            continue
        keep = [c for c in p.iter() if c.tag in COMMENT_TAGS or
                (c.tag == qn("w:r") and c.find(qn("w:commentReference")) is not None)]
        if keep:
            prev = p.getprevious()
            while prev is not None and (prev.tag != qn("w:p") or _mark_deleted(prev)):
                prev = prev.getprevious()
            assert prev is not None, "메모를 옮길 앞 문단이 없음"
            for c in keep:
                prev.append(c)
        p.getparent().remove(p)
        removed += 1
    return removed


def _mark_deleted(p):
    ppr = p.find(qn("w:pPr"))
    rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
    return rpr is not None and rpr.find(qn("w:del")) is not None


doc = Document(SRC)
n_removed = accept_all(doc.element.body)
for part in [r.target_part for r in doc.part.rels.values() if not r.is_external]:
    el = getattr(part, "element", None)
    if el is not None and part.partname.endswith(("header1.xml", "header2.xml", "footer1.xml", "footer2.xml",
                                                   "footer3.xml", "footer4.xml", "footnotes.xml", "endnotes.xml")):
        accept_all(el)
body = doc.element.body
left = [e for e in body.iter(qn("w:ins"), qn("w:del"))]
assert not left, f"남은 변경 {len(left)}건"

# ---------------------------------------------------------------- 특약 제15조 제②항 보완 문장(새 변경추적 1건)
target = None
for p in body.iter(qn("w:p")):
    t = "".join(x.text or "" for x in p.iter(qn("w:t")))
    if t.startswith("신탁계약서 제10조 제②항, 제11조 및 본 특약 제6조에도 불구하고"):
        target = p
        break
assert target is not None
runs = [r for r in target if r.tag == qn("w:r") and r.find(qn("w:t")) is not None]
last = runs[-1]
lt = last.find(qn("w:t"))
if lt.text.endswith(" "):
    lt.text = lt.text.rstrip()
ins = mark("w:ins")
r = OxmlElement("w:r")
if last.find(qn("w:rPr")) is not None:
    r.append(copy.deepcopy(last.find(qn("w:rPr"))))
t = OxmlElement("w:t")
t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
t.text = " " + ADDED
r.append(t)
ins.append(r)
last.addnext(ins)

# 제15조 메모 갱신(반영 사실 덧붙임)
cpart = next(rel.target_part for rel in doc.part.rels.values() if rel.reltype.endswith("/comments"))
hit = 0
for c in cpart.element.iter(qn("w:comment")):
    ts = list(c.iter(qn("w:t")))
    if ts and "[제15조 제②항 ↔ 대출약정서 제11조 제6항 제(2)호 검토]" in "".join(x.text or "" for x in ts):
        ts[-1].text = (ts[-1].text or "") + (" ⇒ [반영 2026-10-13] 운용역 요청에 따라 위 문장을 제15조 제②항 끝에 추가했습니다(변경추적 표시). "
                                             "BKL(대주 측) 및 신탁사 확인이 필요합니다(외부 법무법인 확인 필요).")
        hit += 1
assert hit == 1, hit

# 메모 참조 무결성
ref = {e.get(qn("w:id")) for e in body.iter(qn("w:commentReference"))}
have = {c.get(qn("w:id")) for c in cpart.element.iter(qn("w:comment"))}
assert have <= ref, f"기준점 없는 메모: {have - ref}"

doc.core_properties.last_modified_by = AUTHOR
doc.save(OUT)
print("saved", OUT, "| removed paragraphs:", n_removed, "| comments:", len(have))
