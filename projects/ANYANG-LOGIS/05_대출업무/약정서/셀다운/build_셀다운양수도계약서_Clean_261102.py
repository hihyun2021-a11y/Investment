# -*- coding: utf-8 -*-
"""셀다운 양수도계약서(농협은행 건) 운용역 수정본 → Clean본: 변경추적 전부 반영, 메모 삭제. 원본은 수정하지 않음."""
import os
import re
import shutil
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
q = lambda t: f"{{{W}}}{t}"

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261102_대출약정_셀다운_대출채권양수도계약서_농협은행_차주의견반영초안_v01_운용역수정.docx")
OUT = os.path.join(HERE, "20261102_대출약정_셀다운_대출채권양수도계약서_농협은행_차주의견반영초안_v01_Clean.docx")
COMMENT_PARTS = ("word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml")


def accept(root):
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "numberingChange"):
        for e in root.xpath(f"//w:{tag}", namespaces=NS):
            e.getparent().remove(e)
    for p in reversed(root.xpath("//w:p[w:pPr/w:rPr/w:del]", namespaces=NS)):  # 삭제된 문단 기호 → 다음 문단과 합침
        rpr = p.find(q("pPr")).find(q("rPr"))
        rpr.remove(rpr.find(q("del")))
        nxt = p.getnext()
        if nxt is None or nxt.tag != q("p"):
            continue
        content = [c for c in p if c.tag != q("pPr")]
        npr = nxt.find(q("pPr"))
        pos = 0 if npr is None else list(nxt).index(npr) + 1
        for c in reversed(content):
            nxt.insert(pos, c)
        p.getparent().remove(p)
    for e in root.xpath("//w:pPr/w:rPr/w:ins", namespaces=NS):
        e.getparent().remove(e)
    for e in root.xpath("//w:del|//w:moveFrom|//w:moveFromRangeStart|//w:moveFromRangeEnd|//w:moveToRangeStart|//w:moveToRangeEnd", namespaces=NS):
        e.getparent().remove(e)
    for e in root.xpath("//w:ins|//w:moveTo", namespaces=NS):
        parent, i = e.getparent(), e.getparent().index(e)
        for c in reversed(list(e)):
            parent.insert(i, c)
        parent.remove(e)
    # 메모 표시 제거
    for e in root.xpath("//w:commentRangeStart|//w:commentRangeEnd", namespaces=NS):
        e.getparent().remove(e)
    for e in root.xpath("//w:r[w:commentReference]", namespaces=NS):
        e.getparent().remove(e)


tmp = tempfile.mkdtemp()
with zipfile.ZipFile(SRC) as z:
    z.extractall(tmp)
    names = [n for n in z.namelist() if n not in COMMENT_PARTS]
for n in names:
    if n in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml") or re.match(r"word/(header|footer)\d*\.xml$", n):
        path = os.path.join(tmp, n)
        tree = etree.parse(path)
        accept(tree.getroot())
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
# 메모 파트 연결 정보 정리
rels = os.path.join(tmp, "word/_rels/document.xml.rels")
txt = open(rels, encoding="utf-8").read()
txt = re.sub(r'<Relationship [^>]*Target="comments(Extended|Ids|Extensible)?\.xml"[^>]*/>', "", txt)
open(rels, "w", encoding="utf-8").write(txt)
ct = os.path.join(tmp, "[Content_Types].xml")
txt = open(ct, encoding="utf-8").read()
txt = re.sub(r'<Override PartName="/word/comments(Extended|Ids|Extensible)?\.xml"[^>]*/>', "", txt)
open(ct, "w", encoding="utf-8").write(txt)
st = os.path.join(tmp, "word/settings.xml")
tree = etree.parse(st)
for e in tree.xpath("//w:trackRevisions", namespaces=NS):
    e.getparent().remove(e)
tree.write(st, xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(os.path.join(tmp, n), n)
shutil.rmtree(tmp)
print("saved", OUT)
