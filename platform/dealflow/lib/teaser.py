# -*- coding: utf-8 -*-
"""Teaser PPT(KLI v5.8 서식, 4장): ① 투자개요 ② 자산·시설 ③ 임대현황 ④ 거래구조·일정."""
from __future__ import annotations

import pathlib

from pptx.enum.text import PP_ALIGN
from pptx.parts.image import Image
from pptx.util import Inches

from .common import fmt, kdate, kli_template, num, to_py

import kli_deck as K  # noqa: E402  (common 이 경로를 잡는다)


def _pic(slide, path, x, y, w, h):
    """비율을 유지해 상자 안에 맞춘다."""
    try:
        iw, ih = Image.from_file(str(path)).size
    except Exception:
        return None
    r = min(w / iw, h / ih)
    pw, ph = iw * r, ih * r
    return slide.shapes.add_picture(str(path), Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))


def _placeholder(deck, s, x, y, w, h, text="사진 [미접수]\n01_입력자료/사진 폴더에 넣으면 자동 배치"):
    box = deck.rect(s, x, y, w, h, fill=K.COL1, line=K.LINE_C)
    tf = box.text_frame
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        K.add_run(p, ln, size=10, color=K.GRAY)


def build(d: dict, model: dict | None, path: pathlib.Path) -> pathlib.Path:
    a, rr = d["asset"], d["rr"]
    name = a.get("name", "[자산명]")
    deck = K.KLIDeck("Teaser", footer=f"{name} 투자 Teaser", template=kli_template())
    price = num(a.get("price"))
    gfa = num(a.get("gfa_m2"))
    gfa_py = to_py(gfa)
    ppy = (price * 1e8 / gfa_py / 1e4) if price and gfa_py else None
    photos = [pathlib.Path(p) for p in d["photos"]]
    hl = [a.get(f"highlight{i}") for i in (1, 2, 3) if a.get(f"highlight{i}")]

    # ① 투자개요
    head = a.get("highlight1") or f"{a.get('addr', '')} 소재 {a.get('sector', '')} 매입 기회"
    s = deck.slide("투자개요", head)
    deck.band(s, "자산 개요", top=2.17, left=0.50, width=5.55)
    occ = rr.get("occupancy")
    rows = [["구분", "내용"],
            ["자산명", name],
            ["소재지", a.get("addr") or a.get("jibun") or "[확인 필요]"],
            ["자산유형", a.get("sector", "[확인 필요]")],
            ["대지 / 연면적", f"{fmt(num(a.get('land_m2')), 1)}㎡ / {fmt(gfa, 1)}㎡ ({fmt(gfa_py, 0)}평)"],
            ["규모 / 사용승인", f"{a.get('floors', '[확인 필요]')} / {a.get('approval', '[확인 필요]')}"],
            ["임대율 / WALE", f"{fmt(occ * 100 if occ is not None else None, 1)}% / {fmt(rr.get('wale'), 1)}년"],
            ["매도인", a.get("owner", "[확인 필요]")],
            ["매각 방식", a.get("deal_type", "[확인 필요]")],
            ["희망 매입가", f"{fmt(price, 0)}억원" + (f" (연면적 평당 {fmt(ppy, 0)}만원)" if ppy else "")]]
    deck.table(s, rows, top=2.55, left=0.56, width=5.45, col_widths=[1.45, 4.0], row_h=0.36)
    if photos:
        _pic(s, photos[0], 6.30, 2.17, 4.90, 3.30)
    else:
        _placeholder(deck, s, 6.30, 2.17, 4.90, 3.30)
    deck.band(s, "Investment Highlight", top=5.60, left=6.30, width=4.90)
    deck.text(s, 6.36, 5.95, 4.85, 1.40, [f"{i}. {t}" for i, t in enumerate(hl or ["[투자포인트 입력 필요]"], 1)])
    deck.footnote(s, "* 자료: 매도인 제공 자료·공부(대장·등기부), 당사 정리. 수치는 실사 전 [추정]이며 변경될 수 있음")

    # ② 자산·시설
    s = deck.slide("자산 및 시설 현황", a.get("location") or f"{name} 입지·시설 개요")
    deck.band(s, "건축 개요(공부 기준)", top=2.17, left=0.50, width=4.55)
    srcs = d.get("asset_src", {})
    def v(k, unit=""):
        x = a.get(k)
        return "[확인 필요]" if x in (None, "") else (fmt(num(x), 1) + unit if unit else str(x))
    rows = [["구분", "내용"], ["대지지번", v("jibun")], ["대지면적", v("land_m2", "㎡")], ["연면적", v("gfa_m2", "㎡")],
            ["규모", v("floors")], ["주구조", v("structure")], ["용도지역", v("zoning")], ["주용도", v("use")],
            ["사용승인일", v("approval")], ["소유자", v("owner")], ["제한물권", v("encumbrance")]]
    deck.table(s, rows, top=2.55, left=0.56, width=4.45, col_widths=[1.2, 3.25], row_h=0.36)
    deck.band(s, "사진", top=2.17, left=5.30, width=5.90)
    pics = photos[1:5] if len(photos) > 1 else photos[:1]
    if pics:
        cells = [(5.36, 2.55), (8.30, 2.55), (5.36, 4.95), (8.30, 4.95)] if len(pics) > 1 else [(5.36, 2.55)]
        size = (2.84, 2.30) if len(pics) > 1 else (5.80, 4.70)
        for p, (x, y) in zip(pics, cells):
            _pic(s, p, x, y, *size)
    else:
        _placeholder(deck, s, 5.36, 2.55, 5.80, 4.70)
    cited = sorted({srcs[k] for k in ("jibun", "gfa_m2", "owner") if k in srcs})
    deck.footnote(s, "* 출처: " + (", ".join(cited) if cited else "딜입력") + " — 공부 원본과 대조 필요")

    # ③ 임대현황
    s = deck.slide("임대 현황", (f"임대율 {fmt(occ * 100, 1)}%, WALE {fmt(rr.get('wale'), 1)}년, "
                               f"주요 임차인 {rr['tenants'][0]['tenant']}({rr['tenants'][0]['share'] * 100:.1f}%)")
                   if rr["tenants"] else "임대 현황 [렌트롤 미접수]")
    deck.band(s, "주요 임차인", top=2.17, left=0.50, width=6.10)
    rows = [["임차인", "용도", "면적(평)", "비중", "월임대료(백만원)", "만기"]]
    for t in rr["tenants"][:8]:
        rows.append([t["tenant"], t["uses"] or "-", fmt(t["area_py"], 0), f"{t['share'] * 100:.1f}%",
                     fmt(t["rent_m"] / 1e6, 1), "-" if t["tenant"] == "공실" else kdate(t["expiry"], wd=False)])
    rows.append(["합계", "", fmt(rr["area_py"], 0), "100.0%", fmt(rr["rent_month"] / 1e6, 1), ""])
    deck.table(s, rows, top=2.55, left=0.56, width=6.0, col_widths=[1.35, 1.05, 0.95, 0.7, 1.15, 0.8],
               row_h=0.34, align=["c", "c", "r", "r", "r", "c"])
    deck.band(s, "만기 분포(면적 기준)", top=2.17, left=6.85, width=4.35)
    exp = rr["expiry_by_year"]
    if exp:
        top, h, x0, w = 2.85, 2.55, 7.05, 4.0
        mx = max(exp.values())
        n = len(exp)
        bw = min(0.55, w / n * 0.6)
        for i, (yr, ar) in enumerate(exp.items()):
            bh = h * ar / mx
            x = x0 + (i + 0.5) * w / n - bw / 2
            deck.rect(s, x, top + h - bh, bw, bh, fill=K.BAND)
            deck.text(s, x - 0.3, top + h - bh - 0.27, bw + 0.6, 0.25, [f"{ar / rr['area_py'] * 100:.0f}%"], align=PP_ALIGN.CENTER)
            deck.text(s, x - 0.3, top + h + 0.05, bw + 0.6, 0.25, [f"'{str(yr)[2:]}"], align=PP_ALIGN.CENTER)
        deck.hline(s, x0, top + h, x0 + w, color=K.LINE_TOP, lw=0.75)
    rows = [["구분", "값"], ["평균 임대료", f"{fmt(rr.get('avg_rent_py'), 0)}원/평/월"],
            ["보증금 합계", f"{fmt(rr['deposit'] / 1e8, 1)}억원"], ["월 관리비 합계", f"{fmt(rr['cam_month'] / 1e6, 1)}백만원"]]
    deck.table(s, rows, top=5.95, left=6.90, width=4.25, col_widths=[1.5, 2.75], row_h=0.32)
    deck.footnote(s, "* 자료: 매도인 제공 렌트롤(딜입력 렌트롤 시트). WALE = 임대료 가중 잔여기간, 거래종결 예정일 기준 [추정]")

    # ④ 거래구조·일정
    s = deck.slide("거래 개요 및 일정", f"거래종결 목표 {kdate(d.get('closing'))}, "
                                      f"{a.get('vehicle') or '매수 비히클 [협의 중]'}")
    deck.band(s, "거래 개요(협의안)", top=2.17, left=0.50, width=5.0)
    rows = [["구분", "내용"], ["매수 비히클", a.get("vehicle", "[협의 중]")], ["운용기간", f"{a.get('term', '[ ]')}년"],
            ["매입가", f"{fmt(price, 0)}억원"]]
    if model:
        rows += [["예상 NOI(1년차)", f"{fmt(model['noi'][0], 1)}억원 [추정]"],
                 ["Going-in Cap", f"{model['cap_in'] * 100:.2f}% [추정]"],
                 ["LTV(매입가)", f"{model['ltv_price'] * 100:.1f}% [협의 중]" if model.get("ltv_price") else "[협의 중]"]]
    rows += [["매각자문사", a.get("seller_advisor", "[확인 필요]")]]
    deck.table(s, rows, top=2.55, left=0.56, width=4.9, col_widths=[1.5, 3.4], row_h=0.36)
    deck.band(s, "주요 일정", top=2.17, left=5.75, width=5.45)
    keys = ["bid_notice", "bid_due", "preferred", "mou", "dd_start", "ic", "board", "closing"]
    rows = [["일정", "일자", "비고"]]
    for m in d["schedule"]:
        if m["key"] in keys:
            rows.append([m["name"], kdate(m["date"]), "[추정]" if m["estimated"] else ""])
    deck.table(s, rows, top=2.55, left=5.81, width=5.35, col_widths=[2.95, 1.45, 0.95], row_h=0.36,
               align=["l", "c", "c"])
    contact = "  |  ".join(x for x in [f"{a.get('company', '코람코자산신탁')} {a.get('manager', '[담당자]')} {a.get('manager_title', '')}".strip(),
                                     a.get('manager_tel'), a.get('manager_email')] if x)
    deck.text(s, 0.56, 6.65, 10.6, 0.55, [("문의: " + contact,
                                         {"font": K.F_MEDIUM})])
    deck.footnote(s, "* 본 자료는 투자 검토를 위한 요약이며 수익을 보장하지 않음. 일정·조건은 협의에 따라 변경될 수 있음. [추정] = 표준 일정 간격 적용")
    bad = deck.check_min_font()
    if bad:
        raise RuntimeError(f"10pt 미만 글자: {bad[:3]}")
    return deck.save(path)
