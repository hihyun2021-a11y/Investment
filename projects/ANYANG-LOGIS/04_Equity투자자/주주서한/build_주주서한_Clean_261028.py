# -*- coding: utf-8 -*-
"""주주서한 v3.3 변경추적본 → Clean본 (모든 변경 내용 적용, 서식 변경 기록 제거). 원본은 수정하지 않음."""
import os
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
q = lambda t: f"{{{W}}}{t}"

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261028_주주서한_KLI주주서한2026년10월_v3.3_변경추적.docx")
OUT = os.path.join(HERE, "20261028_주주서한_KLI주주서한2026년10월_v3.3_Clean.docx")


def accept(root):
    # 1) 서식·문단속성 변경 기록 제거(현재 서식 유지)
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "numberingChange"):
        for e in root.iter(q(tag)):
            pass
        for e in root.xpath(f"//w:{tag}", namespaces=NS):
            e.getparent().remove(e)
    # 2) 삭제된 문단 기호 → 다음 문단과 합침(뒤에서부터 처리)
    for p in reversed(root.xpath("//w:p[w:pPr/w:rPr/w:del]", namespaces=NS)):
        p.find(q("pPr")).find(q("rPr")).remove(p.find(q("pPr")).find(q("rPr")).find(q("del")))
        nxt = p.getnext()
        if nxt is None or nxt.tag != q("p"):
            continue
        content = [c for c in p if c.tag != q("pPr")]
        npr = nxt.find(q("pPr"))
        pos = 0 if npr is None else list(nxt).index(npr) + 1
        for c in reversed(content):
            nxt.insert(pos, c)
        p.getparent().remove(p)
    # 3) 삽입된 문단 기호 표시 제거
    for e in root.xpath("//w:pPr/w:rPr/w:ins", namespaces=NS):
        e.getparent().remove(e)
    # 4) 삭제된 내용 제거
    for e in root.xpath("//w:del|//w:moveFrom|//w:moveFromRangeStart|//w:moveFromRangeEnd|//w:moveToRangeStart|//w:moveToRangeEnd", namespaces=NS):
        e.getparent().remove(e)
    # 5) 삽입된 내용은 감싼 태그만 풀기
    for e in root.xpath("//w:ins|//w:moveTo", namespaces=NS):
        parent = e.getparent()
        i = parent.index(e)
        for c in reversed(list(e)):
            parent.insert(i, c)
        parent.remove(e)
    # 6) 내용이 사라진 run 정리
    for r in root.xpath("//w:r[not(w:t|w:tab|w:br|w:drawing|w:pict|w:fldChar|w:instrText|w:sym|w:footnoteReference|w:endnoteReference|w:object|w:cr|w:noBreakHyphen|w:softHyphen|w:ptab)]", namespaces=NS):
        r.getparent().remove(r)


tmp = tempfile.mkdtemp()
with zipfile.ZipFile(SRC) as z:
    z.extractall(tmp)
    names = z.namelist()
for n in names:
    if n.startswith("word/") and n.endswith(".xml") and n.split("/")[-1] in ("document.xml", "footnotes.xml", "endnotes.xml") or n.startswith("word/header") or n.startswith("word/footer"):
        path = os.path.join(tmp, n)
        tree = etree.parse(path)
        accept(tree.getroot())
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
# 변경 추적 기록 설정 해제
sp = os.path.join(tmp, "word/settings.xml")
st = etree.parse(sp)
for e in st.xpath("//w:trackRevisions", namespaces=NS):
    e.getparent().remove(e)
st.write(sp, xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(os.path.join(tmp, n), n)
shutil.rmtree(tmp)
print("saved", OUT)
