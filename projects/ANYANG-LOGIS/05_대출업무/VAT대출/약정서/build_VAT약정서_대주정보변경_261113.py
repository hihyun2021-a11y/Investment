# -*- coding: utf-8 -*-
"""VAT 대출 약정서 3종 — 대주(우리은행) 표기를 투자금융부 → 영업점으로 변경 (2026-11-13 우리은행 송준민 계장 요청).
원본(BKL clean)은 그대로 두고 새 파일을 만든다. 서식(run 서식)은 유지하고 해당 run의 글자만 바꾼다.
"""
import copy
import pathlib
import docx
from docx.oxml.ns import qn

HERE = pathlib.Path(__file__).resolve().parent
ADDR = "서울특별시 서초구 방배천로 91, 1층(방배동, 구산타워)"
TITLE = "[지점장]"          # 메일에 직위 없음 → 우리은행 확인 후 확정
TEL = "02-523-2171"
PIC = "이택주 부지점장/송준민 계장"
MAIL = "[*]"               # 메일에 이메일 없음 → 우리은행 확인 후 기재


def put(p, i, j, text):
    """runs[i..j] 를 text 하나로 바꾼다(첫 run 서식 유지)."""
    rs = p.runs
    rs[i].text = text
    for r in rs[i + 1:j + 1]:
        r._r.getparent().remove(r._r)


def find(doc, start, pred):
    for k, p in enumerate(doc.paragraphs[start:], start):
        if pred(p.text):
            return k
    raise ValueError("문단을 찾지 못함")


def runs_idx(p, word):
    for k, r in enumerate(p.runs):
        if r.text == word:
            return k
    raise ValueError(word)


def sign_block(doc, spaced):
    k = find(doc, 0, lambda t: t.startswith("위 대리인 투자금융부장"))
    ps = doc.paragraphs
    # 주소 문단: 대출약정은 주소와 '대표이사'가 한 문단(줄바꿈 없음)
    a = ps[k - 1] if "소공로" in ps[k - 1].text else ps[k - 2]
    end = runs_idx(a, ")")
    put(a, 0, end, ADDR)
    if "대표이사" in a.text:   # 원본에 줄바꿈이 빠져 있어 주소 뒤에 줄바꿈 추가
        a.runs[0]._r.append(a.runs[0]._r.makeelement(qn("w:br"), {}))
    s = ps[k]
    t = runs_idx(s, "투자금융부장")
    s.runs[t].text = TITLE
    if spaced:  # '이',' ','광',' ','옥'
        n0 = t + 2
        for off, ch in zip((0, 2, 4), "배동호"):
            s.runs[n0 + off].text = ch
    else:
        s.runs[runs_idx(s, "이광옥")].text = "배동호"


def loan(src, dst):
    d = docx.Document(src)
    sign_block(d, spaced=True)
    k = find(d, 600, lambda t: t.replace(" ", "") == "대주")
    ps = d.paragraphs
    assert ps[k + 1].text == "주식회사 우리은행"
    a, tel, pic, em = ps[k + 2], ps[k + 3], ps[k + 4], ps[k + 5]
    put(a, 6, len(a.runs) - 1, ADDR)
    put(tel, 6, len(tel.runs) - 1, TEL)
    put(pic, 3, len(pic.runs) - 1, PIC)
    put(em, 2, len(em.runs) - 1, MAIL)
    d.save(dst)


def security(src, dst):
    d = docx.Document(src)
    sign_block(d, spaced=False)
    d.save(dst)


if __name__ == "__main__":
    J = [
        (loan, "KLL_부가가치세대출약정서(26376115.v6)_clean_BKL원본.docx", "20261113_대출약정_VAT_부가가치세대출약정서_대주정보변경_v01.docx"),
        (security, "KLL_예금채권근질권설정계약서_부가세환급계좌(26376116.v4)_clean_BKL원본.docx", "20261113_대출약정_VAT_예금채권근질권설정계약서_대주정보변경_v01.docx"),
        (security, "KLL_채권양도담보권설정계약서_부가가치세(26376117.v4)_clean_BKL원본.docx", "20261113_대출약정_VAT_채권양도담보권설정계약서_대주정보변경_v01.docx"),
    ]
    for fn, s, t in J:
        fn(HERE / s, HERE / t)
        print("작성:", t)
