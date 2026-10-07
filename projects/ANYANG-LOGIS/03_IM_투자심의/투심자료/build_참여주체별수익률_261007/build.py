# -*- coding: utf-8 -*-
"""운용역 제공 양식(2. 참여 주체별 역할 및 수익률, 261007)에 안양물류센터 투자구조를 반영.
사용법: python3 -I build.py <양식.pptx> <출력.pptx>
수치 출처: 기준 재무모델 v04(2026-10-06) IM·Sell-down 탭, 확정 T/S(9/22)·대출약정 v16 중순위 의견본(10/06),
          주주간계약서 BKL v30.5, 투심자료 v12(투자자 구성).
"""
import copy, sys
from pptx import Presentation
from lxml import etree

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
s = prs.slides[0]
sh = {x.name: x for x in s.shapes}


# ---------- text helpers ----------
def set_paras(txBody, lines):
    ps = txBody.findall(A + 'p')
    tmpl = [copy.deepcopy(p) for p in ps] or [etree.SubElement(txBody, A + 'p')]
    for p in ps:
        txBody.remove(p)
    for i, line in enumerate(lines):
        t = tmpl[min(i, len(tmpl) - 1)]
        p = etree.SubElement(txBody, A + 'p')
        ppr = t.find(A + 'pPr')
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        r0 = t.find(A + 'r')
        rpr = r0.find(A + 'rPr') if r0 is not None else t.find(A + 'endParaRPr')
        r = etree.SubElement(p, A + 'r')
        if rpr is not None:
            rp = copy.deepcopy(rpr); rp.tag = A + 'rPr'; r.append(rp)
        etree.SubElement(r, A + 't').text = line


def shape_text(shape, text):
    set_paras(shape.text_frame._txBody, text.split('\n'))


def tc_text(tc, text):
    set_paras(tc.find(A + 'txBody'), text.split('\n'))


# ---------- 상단·좌측 ----------
s.shapes._spTree.remove(sh['직사각형 9']._element)  # '예시' 표시 삭제
shape_text(sh['리츠'], '코람코라이프로지스리츠 (KLL)\n( 총 투자비 2,989억  ·  AMC 코람코자산신탁 )')
shape_text(sh['자산명'], '안양물류센터')

grp = sh['그룹 5']
g = {x.name: x for x in grp.shapes}
Y0 = 2407008


def place(shape, top, h):
    shape.top = Y0 + top
    shape.height = h


# 투자비 (좌): 매입 2,700 / 취득부대비용 등 189 / 예비비 100  (공사비 해당 없음 → 삭제)
grp._element.remove(g['투자비 공사비 등']._element)
place(g['투자비 부동산 매입금액'], 0, 1936000); shape_text(g['투자비 부동산 매입금액'], '부동산 매입금액\n2,700')
place(g['투자비 취득부대비용 등'], 1936000, 400000); shape_text(g['투자비 취득부대비용 등'], '취득부대비용 등 189')
place(g['투자비 보증금·예비비 등'], 2336000, 400000); shape_text(g['투자비 보증금·예비비 등'], '예비비 100')

# 재원 (우): 보증금 21 / 선순위 1,600 / 중순위 380 / 1종 240 / 2종 500 / 3종 100 / 보통주 148
eq_tmpl = g['재원 1종 종류주']
new2 = copy.deepcopy(eq_tmpl._element); eq_tmpl._element.addnext(new2)
new3 = copy.deepcopy(eq_tmpl._element); new2.addnext(new3)
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
maxid = max(int(e.get('id')) for e in s.shapes._spTree.iter(P + 'cNvPr'))
for k, el in enumerate((new2, new3), 1):
    el.find('.//' + P + 'cNvPr').set('id', str(maxid + k))
g = {x.name: x for x in grp.shapes}
boxes = [x for x in grp.shapes if x.name == '재원 1종 종류주']
b2, b3 = boxes[1], boxes[2]
b2.name, b3.name = '재원 2종 종류주', '재원 3종 종류주'
plan = [(g['재원 보증금'], 300000, '보증금 21'),
        (g['재원 선순위 대출'], 776000, '선순위 대출\n1,600\nLTV 48.5%'),
        (g['재원 중순위 대출'], 360000, '중순위 대출 380\n(누적 LTV 59.8%)'),
        (boxes[0], 300000, '1종 종류주 240'),
        (b2, 400000, '2종 종류주 500'),
        (b3, 300000, '3종 종류주 100'),
        (g['재원 보통주'], 300000, '보통주 148')]
y = 0
for shp, h, txt in plan:
    place(shp, y, h); shape_text(shp, txt); y += h
assert y == 2736000

# ---------- 표 ----------
tshape = sh['주체별 수익 표']
tbl = tshape._element.find('.//' + A + 'tbl')
rows = tbl.findall(A + 'tr')
TC = lambda tr: tr.findall(A + 'tc')


def fill(tr, vals):
    """vals: dict col -> text"""
    cs = TC(tr)
    for c, v in vals.items():
        tc_text(cs[c], v)


def merge(tc, span=None, cont=False):
    for k in ('rowSpan', 'vMerge'):
        tc.attrib.pop(k, None)
    if span and span > 1:
        tc.set('rowSpan', str(span))
    if cont:
        tc.set('vMerge', '1')


# Loan
r2, r6, r7, r8, r9 = rows[2], rows[6], rows[7], rows[8], rows[9]
for r in rows[3:6]:
    tbl.remove(r)
merge(TC(r2)[0], 6)
merge(TC(r2)[1], None)
fill(r2, {1: '선순위\n(1,600억,\n5.50%)', 2: '우리은행', 3: '은행', 4: '· 선순위 Tr.A\n· LOC 1,600억 (9/21)',
          5: '1,600', 6: '5.50%', 7: '0.73%', 8: '4.78%', 9: '-'})
merge(TC(r6)[1], 3)
fill(r6, {1: '중순위\n(380억,\n6.50%)', 2: '키움캐피탈', 3: '캐피탈사', 4: '· 중순위 (LOC 200억)', 5: '200',
          6: '6.50%', 7: '0.50%', 8: '6.00%', 9: '-'})
fill(r7, {2: '엠지캐피탈', 3: '캐피탈사', 4: '· 중순위 (LOC 100억)', 5: '100', 6: '6.50%', 7: '0.50%', 8: '6.00%', 9: '-'})
r7b = copy.deepcopy(r7); r7.addnext(r7b)
fill(r7b, {2: 'DB저축은행', 3: '저축은행', 4: '· 중순위 (LOC [미접수])', 5: '80', 6: '6.50%', 7: '0.50%', 8: '6.00%', 9: '-'})
fill(r8, {1: '보증금\n(21억)', 2: '쿠팡', 3: '임차인', 4: '· 임대차보증금 (무수익)', 5: '21', 6: '0.0%'})
r2.set('h', '330000')
fill(r9, {1: 'Loan 소계 (가중평균)', 5: '2,001', 6: '5.63%', 7: '0.67%', 8: '4.96%'})

# Equity
rows = tbl.findall(A + 'tr')
e_start = [i for i, r in enumerate(rows) if ''.join(x.text or '' for x in TC(r)[0].iter(A + 't')) == 'Equity'][0]
r13, r14, r15, r16 = rows[e_start:e_start + 4]
T_first_c0 = copy.deepcopy(TC(r13)[0])
T_grp_c1 = copy.deepcopy(TC(r14)[1])
T_black, T_red, T_sub = copy.deepcopy(r15), copy.deepcopy(r13), r16


def build(tmpl, c0=None, c1=None, vals=None, h=None):
    tr = copy.deepcopy(tmpl)
    cs = TC(tr)
    if c0 is not None:
        tr.replace(cs[0], copy.deepcopy(c0))
    else:
        merge(cs[0], cont=True)
        tc_text(cs[0], '')
    cs = TC(tr)
    if c1 is not None:
        new = copy.deepcopy(T_grp_c1)
        tr.replace(cs[1], new)
        span, txt = c1
        merge(new, span)
        tc_text(new, txt)
    else:
        merge(cs[1], cont=True)
        tc_text(cs[1], '')
    fill(tr, vals or {})
    if h:
        tr.set('h', str(h))
    return tr


H = 250000
AMT_BLACK = copy.deepcopy(TC(r14)[5])


def blk(tr):
    cs = TC(tr)
    amt = ''.join(x.text or '' for x in cs[5].iter(A + 't'))
    new = copy.deepcopy(AMT_BLACK)
    tr.replace(cs[5], new)
    tc_text(new, amt)
    return tr


E = [
    blk(build(T_black, c0=T_first_c0, c1=(1, '보통주\n(148억)'),
              vals={2: 'KLI', 3: '상장 리츠\n(모리츠)', 4: '· 보통주 출자 (100%)\n· 의결권 71.4%',
                    5: '148', 6: '16.18%', 7: '0.75%', 8: '31.80%', 9: '-', 10: ''}, h=H)),
    blk(build(T_black, c1=(1, '1종\n종류주\n(240억)'),
              vals={2: '1종 FI 4개사', 3: 'FI', 4: '· 1종 인수 (LOC)\n· 2년 후 KLI 매도',
                    5: '240', 6: '7.12%', 7: '7.00%', 8: '-', 9: '-', 10: ''}, h=H)),
    blk(build(T_black, c1=(2, '2종\n종류주\n(500억)'),
              vals={2: '2종 FI 3개사', 3: 'FI', 4: '· 2종 인수 (LOC)\n· 1년 후 KLI 매도',
                    5: '340', 6: '7.59%', 7: '7.50%', 8: '-', 9: '-', 10: ''}, h=H)),
    blk(build(T_black,
              vals={2: 'KLI', 3: '상장 리츠', 4: '· 2종 출자\n· C.G 20% 배분',
                    5: '160', 6: '9.15%', 7: '7.50%', 8: '2.35%', 9: '-', 10: ''}, h=H)),
    build(T_red, c1=(1, '3종\n종류주\n(100억)'),
          vals={2: '코크렙안양', 3: '매도인\n(LF 계열)', 4: '· 매매대금 상계납입\n· 2년 내 유상감자',
                5: '100', 6: '0.00%', 7: '0.00%', 8: '0.00%', 9: '본건 매도', 10: '2,700'}, h=H),
    build(T_black, c1=(1, 'KLI 합산'),
          vals={2: 'KLI', 3: '상장 리츠', 4: '· 1·2종 580 매입\n· 최종 888억 보유',
                5: '(580)', 6: '10.99%', 7: '5.44%', 8: '6.63%', 9: '-', 10: ''}, h=H),
]
merge(TC(E[0])[0], len(E) + 1)
for r in (r13, r14, r15):
    tbl.remove(r)
anchor = tbl.findall(A + 'tr')[e_start - 1]
for r in E:
    anchor.addnext(r); anchor = r
fill(T_sub, {1: 'Equity 소계 (3종 제외 888억 기준 수익률)', 5: '988', 6: '10.21%', 7: '6.10%', 8: '6.48%'})
last = tbl.findall(A + 'tr')[-1]
fill(last, {5: '2,989'})

# ---------- 서식 통일 (편집성) ----------
# 양식 예시 서식(빨간 수치·회색 금액·큰 기관명 등)이 남지 않도록 데이터 셀(기관~부수용역 열)을
# 같은 열의 대주 행(키움캐피탈) 서식으로 통일. 붉은색은 작성요령 규칙대로 출자·용역 겸업 셀만 유지.
DARK, RED = '262626', 'C00000'
ref = {}
for ci, tc in enumerate(TC(r6)):
    r0 = tc.find('.//' + A + 'r')
    if r0 is not None and r0.find(A + 'rPr') is not None:
        ref[ci] = copy.deepcopy(r0.find(A + 'rPr'))


def recolor(rpr, val):
    for sf in rpr.findall(A + 'solidFill'):
        rpr.remove(sf)
    sf = etree.Element(A + 'solidFill'); c = etree.SubElement(sf, A + 'srgbClr'); c.set('val', val)
    ln = rpr.find(A + 'ln')
    (ln.addnext(sf) if ln is not None else rpr.insert(0, sf))
    rpr.attrib.pop('dirty', None)
    return rpr


data_rows = [r2, r6, r7, r7b, r8] + E
red_cells = {(id(E[4]), 2), (id(E[4]), 9), (id(E[4]), 10)}
for tr in data_rows:
    for ci, tc in enumerate(TC(tr)):
        if ci < 2 or ci not in ref:
            continue
        col = RED if (id(tr), ci) in red_cells else DARK
        for r in tc.iter(A + 'r'):
            old = r.find(A + 'rPr')
            new = recolor(copy.deepcopy(ref[ci]), col)
            if old is not None:
                new.set('lang', old.get('lang', 'ko-KR'))
                r.replace(old, new)
            else:
                r.insert(0, new)
# 소계·합계 행, 라벨 열은 양식 서식 유지하되 색만 기본색으로(빨간 예시값 제거)
for tr in (r9, T_sub, tbl.findall(A + 'tr')[-1]):
    for ci, tc in enumerate(TC(tr)):
        for r in tc.iter(A + 'r'):
            rp = r.find(A + 'rPr')
            if rp is not None and rp.find(A + 'solidFill') is not None:
                v = rp.find(A + 'solidFill')[0].get('val')
                if v in ('C00000', 'FF0000', 'A6A6A6', '808080', '7F7F7F', 'BFBFBF'):
                    recolor(rp, DARK)
for tr in data_rows:
    for ci, tc in enumerate(TC(tr)[:2]):
        for r in tc.iter(A + 'r'):
            rp = r.find(A + 'rPr')
            if rp is not None and rp.find(A + 'solidFill') is not None and rp.find(A + 'solidFill')[0].get('val') in ('A6A6A6', '808080', '7F7F7F', 'BFBFBF', 'C00000'):
                recolor(rp, '595959')
# 빈 run 제거, 모든 문단에 endParaRPr(입력 시 같은 서식 유지), dirty 속성 제거
for txb in s.shapes._spTree.iter(A + 'txBody', '{http://schemas.openxmlformats.org/presentationml/2006/main}txBody'):
    for p in txb.findall(A + 'p'):
        runs = p.findall(A + 'r')
        for r in runs:
            t = r.find(A + 't')
            if (t is None or not (t.text or '')) and len(runs) > 1:
                p.remove(r)
        runs = p.findall(A + 'r')
        if p.find(A + 'endParaRPr') is None and runs and runs[-1].find(A + 'rPr') is not None:
            e = copy.deepcopy(runs[-1].find(A + 'rPr')); e.tag = A + 'endParaRPr'; p.append(e)
for el in s.shapes._spTree.iter():
    if 'dirty' in el.attrib:
        del el.attrib['dirty']
# 표 스타일: 기본 스타일 영향 차단(서식은 셀에 직접 지정됨)
tblPr = tbl.find(A + 'tblPr')
if tblPr.find(A + 'tableStyleId') is None:
    sid = etree.SubElement(tblPr, A + 'tableStyleId'); sid.text = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

# 표 높이·주석 위치
total_h = sum(int(r.get('h')) for r in tbl.findall(A + 'tr'))
tshape.height = total_h
note = sh['작성요령']
note.top = 6250000
shape_text(note,
           '※ 대주 : All-in = 취급수수료(24개월 연환산) + Coupon. 보증금 21억(쿠팡 5~7층, 담보신탁 1순위)은 무이자 조달로 Loan에 포함(그 외 보증금 50.7억은 예금질권 예치로 S&U 제외). 중순위 구성은 10/6 의견본 기준\n'
           '※ 투자자 : IRR(C.G 포함)·I.G(운영배당 CoC)·C.G(매각차익÷투자금÷10년), 재무모델 Base case. KLI = 코람코라이프인프라리츠. 1종 FI = 이지스리츠3호·키움·애큐온·엠지캐피탈, 2종 FI = 삼성증권·이지스K리츠·기계설비조합(발행가 매도, C.G 없음)\n'
           '※ 출자·용역 겸업 기관은 붉은색 / [ ] 는 미확정 / 출처 : 재무모델 v04(2026-10-06) IM·Sell-down 탭, 확정 T/S(9/22), 주주간계약서 BKL v30.5, 투심자료 v12')
prs.save(OUT)
print('saved', OUT, 'table h', total_h, 'note bottom', note.top + note.height)
