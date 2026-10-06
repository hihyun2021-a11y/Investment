# -*- coding: utf-8 -*-
"""2단계: 1단계 시트(standalone xlsx)를 기준 재무모델 .xlsm 패키지에 직접 이식한다.
매크로(vbaProject)·양식컨트롤·웹확장·스레드 메모 등 기존 파트는 바이트 그대로 유지하고,
styles.xml에 서식만 추가 병합, workbook/rels/content-types에 시트 1개만 추가한다.
사용법: python3 -I stage_b.py <model.xlsm> <stage_a.xlsx> <out.xlsm> [values.json]
values.json(선택): {셀주소: 값} — 수식 셀의 캐시값(<v>)으로 기록.
"""
import sys, json, re, zipfile, copy
from lxml import etree

MODEL, SRC, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
VALS = json.load(open(sys.argv[4], encoding="utf-8")) if len(sys.argv) > 4 else {}
SHEET_NAME = "A&R(사내양식)"
AFTER = "A&R"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
q = lambda t: f"{{{NS}}}{t}"
ser = lambda el: b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + etree.tostring(el, encoding="UTF-8")

zm = zipfile.ZipFile(MODEL)
zs = zipfile.ZipFile(SRC)
parts = {n: zm.read(n) for n in zm.namelist()}
infos = {i.filename: i for i in zm.infolist()}

def src_sheet_path():
    wbx = etree.fromstring(zs.read("xl/workbook.xml"))
    rid = wbx.find(f"{q('sheets')}/{q('sheet')}").get(f"{{{RNS}}}id")
    rels = etree.fromstring(zs.read("xl/_rels/workbook.xml.rels"))
    for r in rels:
        if r.get("Id") == rid:
            t = r.get("Target").lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t

# ── 서식 병합 ───────────────────────────────────────────────────────────
mst = etree.fromstring(parts["xl/styles.xml"])
sst = etree.fromstring(zs.read("xl/styles.xml"))

def sect(root, name, create_before=None):
    e = root.find(q(name))
    if e is None:
        e = etree.Element(q(name))
        if create_before is not None:
            root.find(q(create_before)).addprevious(e)
        else:
            root.insert(0, e)
    return e

def key(el):
    return etree.tostring(el, method="c14n")

def merge_list(name):
    mdst = sect(mst, name)
    src = sst.find(q(name))
    existing = {key(e): i for i, e in enumerate(mdst)}
    mapping = {}
    for i, e in enumerate(src if src is not None else []):
        k = key(e)
        if k not in existing:
            mdst.append(copy.deepcopy(e)); existing[k] = len(mdst) - 1
        mapping[i] = existing[k]
    mdst.set("count", str(len(mdst)))
    return mapping

fmap, fillmap, bmap = merge_list("fonts"), merge_list("fills"), merge_list("borders")
mnf = mst.find(q("numFmts"))
if mnf is None:
    mnf = etree.Element(q("numFmts")); mst.insert(0, mnf)
code2id = {e.get("formatCode"): int(e.get("numFmtId")) for e in mnf}
nextid = max([int(e.get("numFmtId")) for e in mnf] + [163]) + 1
nfmap = {}
snf = sst.find(q("numFmts"))
for e in (snf if snf is not None else []):
    code, sid = e.get("formatCode"), int(e.get("numFmtId"))
    if code not in code2id:
        etree.SubElement(mnf, q("numFmt"), numFmtId=str(nextid), formatCode=code)
        code2id[code] = nextid; nextid += 1
    nfmap[sid] = code2id[code]
mnf.set("count", str(len(mnf)))
mxf = mst.find(q("cellXfs"))
xkeys = {key(e): i for i, e in enumerate(mxf)}
xmap = {}
for i, xf in enumerate(sst.find(q("cellXfs"))):
    x = copy.deepcopy(xf)
    x.set("fontId", str(fmap.get(int(x.get("fontId", 0)), 0)))
    x.set("fillId", str(fillmap.get(int(x.get("fillId", 0)), 0)))
    x.set("borderId", str(bmap.get(int(x.get("borderId", 0)), 0)))
    n = int(x.get("numFmtId", 0))
    x.set("numFmtId", str(nfmap.get(n, n)))
    x.set("xfId", "0")
    k = key(x)
    if k not in xkeys:
        mxf.append(x); xkeys[k] = len(mxf) - 1
    xmap[i] = xkeys[k]
mxf.set("count", str(len(mxf)))
parts["xl/styles.xml"] = ser(mst)

# ── 시트 XML 정리 ────────────────────────────────────────────────────────
sx = etree.fromstring(zs.read(src_sheet_path()))
ss = []
if "xl/sharedStrings.xml" in zs.namelist():
    ss = list(etree.fromstring(zs.read("xl/sharedStrings.xml")))
for c in sx.iter(q("c")):
    if c.get("s") is not None:
        c.set("s", str(xmap[int(c.get("s"))]))
    if c.get("t") == "s":
        si = ss[int(c.find(q("v")).text)]
        c.remove(c.find(q("v")))
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, q("is"))
        for ch in si:
            is_.append(copy.deepcopy(ch))
    ref = c.get("r")
    f = c.find(q("f"))
    if f is not None and ref in VALS:
        v = c.find(q("v"))
        if v is None:
            v = etree.SubElement(c, q("v"))
        val = VALS[ref]
        if isinstance(val, bool):
            c.set("t", "b"); v.text = "1" if val else "0"
        elif isinstance(val, (int, float)):
            c.attrib.pop("t", None); v.text = repr(float(val)) if isinstance(val, float) else str(val)
        elif isinstance(val, str) and val.startswith("#"):
            c.set("t", "e"); v.text = val
        else:
            c.set("t", "str"); v.text = str(val)
for r in sx.iter(q("row")):
    if r.get("s") is not None:
        r.set("s", str(xmap[int(r.get("s"))]))
for col in sx.iter(q("col")):
    if col.get("style") is not None:
        col.set("style", str(xmap[int(col.get("style"))]))
for tag in ("drawing", "legacyDrawing", "picture", "tableParts", "legacyDrawingHF"):
    for e in sx.findall(q(tag)):
        sx.remove(e)
for e in sx.findall(q("pageSetup")):
    e.attrib.pop(f"{{{RNS}}}id", None)
for sv in sx.iter(q("sheetView")):
    sv.attrib.pop("tabSelected", None)

# ── 패키지에 시트 추가 ───────────────────────────────────────────────────
nums = [int(m.group(1)) for n in parts for m in [re.match(r"xl/worksheets/sheet(\d+)\.xml$", n)] if m]
new_path = f"xl/worksheets/sheet{max(nums) + 1}.xml"
parts[new_path] = ser(sx)

rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
rid_nums = [int(re.sub(r"\D", "", r.get("Id")) or 0) for r in rels]
new_rid = f"rId{max(rid_nums) + 1}"
etree.SubElement(rels, f"{{{PR}}}Relationship", Id=new_rid,
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                 Target=new_path.replace("xl/", ""))
parts["xl/_rels/workbook.xml.rels"] = ser(rels)

wbx = etree.fromstring(parts["xl/workbook.xml"])
sheets = wbx.find(q("sheets"))
names = [s.get("name") for s in sheets]
pos = names.index(AFTER) + 1
new_id = max(int(s.get("sheetId")) for s in sheets) + 1
ne = etree.Element(q("sheet"), name=SHEET_NAME, sheetId=str(new_id))
ne.set(f"{{{RNS}}}id", new_rid)
sheets.insert(pos, ne)
dns = wbx.find(q("definedNames"))
if dns is not None:
    for d in dns:
        ls = d.get("localSheetId")
        if ls is not None and int(ls) >= pos:
            d.set("localSheetId", str(int(ls) + 1))
    pa = etree.SubElement(dns, q("definedName"), name="_xlnm.Print_Area", localSheetId=str(pos))
    pa.text = f"'{SHEET_NAME}'!$A$1:$AM$142"
for bv in wbx.iter(q("workbookView")):
    for a in ("activeTab", "firstSheet"):
        if bv.get(a) is not None and int(bv.get(a)) >= pos:
            bv.set(a, str(int(bv.get(a)) + 1))
parts["xl/workbook.xml"] = ser(wbx)

ct = etree.fromstring(parts["[Content_Types].xml"])
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
etree.SubElement(ct, f"{{{CTNS}}}Override", PartName="/" + new_path,
                 ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
parts["[Content_Types].xml"] = ser(ct)

# docProps/app.xml 시트 목록 갱신 (선택)
try:
    ap = etree.fromstring(parts["docProps/app.xml"])
    VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    tvec = ap.find(f"{{{EP}}}TitlesOfParts/{{{VT}}}vector")
    lps = list(tvec)
    texts = [e.text for e in lps]
    if AFTER in texts:
        i = texts.index(AFTER)
        el = etree.Element(f"{{{VT}}}lpstr"); el.text = SHEET_NAME
        lps[i].addnext(el)
        tvec.set("size", str(int(tvec.get("size")) + 1))
        hp = ap.find(f"{{{EP}}}HeadingPairs/{{{VT}}}vector")
        vs = list(hp)
        for j in range(0, len(vs), 2):
            nm = vs[j].find(f"{{{VT}}}lpstr")
            if nm is not None and nm.text in ("워크시트", "Worksheets"):
                cnt = vs[j + 1].find(f"{{{VT}}}i4"); cnt.text = str(int(cnt.text) + 1)
        parts["docProps/app.xml"] = ser(ap)
except Exception as ex:  # noqa
    print("app.xml skip:", ex)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zo:
    order = [i.filename for i in zm.infolist()] + [new_path]
    for n in order:
        info = infos.get(n)
        if info is not None and n not in ("xl/styles.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                                          "[Content_Types].xml", "docProps/app.xml"):
            zo.writestr(info, parts[n])          # 기존 파트 원본 바이트·압축방식 유지
        else:
            zo.writestr(n, parts[n])
print("saved", OUT, "sheet", new_path, "pos", pos, "xf+", len(mxf))
