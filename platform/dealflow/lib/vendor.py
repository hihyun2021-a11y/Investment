# -*- coding: utf-8 -*-
"""실사업체 선정: 견적 비교·선정 보고장표(PPT 1장, KLI v5.8) + 기안문(Word, 사내 기안 양식) + 전자결재 붙여넣기용 본문(.txt)
+ 선정·미선정 통보 임시 메일(.eml)."""
from __future__ import annotations

import pathlib

from . import eml
from .common import TODAY, dd_label, fmt, kdate, kli_template, load_json, outpath
from .docx_util import CENTER, Doc

import kli_deck as K  # noqa: E402

ORDER = ["감정", "법률", "재무", "시장", "물리"]


def _sched(d, key):
    return next((m["date"] for m in d["schedule"] if m["key"] == key), None)


def summary(d):
    by = {}
    for v in d["vendors"]:
        by.setdefault(v["실사구분"], []).append(v)
    sel = {k: next((v for v in vs if v["선정"]), None) for k, vs in by.items()}
    return by, sel


def build(d: dict, folder: pathlib.Path) -> dict:
    by, sel = summary(d)
    chosen = {k: v for k, v in sel.items() if v}
    if not chosen:
        return {}
    a = d["asset"]
    name = a.get("name", "[자산명]")
    scopes = load_json("rfp_scopes.json")["types"]
    kinds = [k for k in ORDER if k in by] + [k for k in by if k not in ORDER]
    total = sum((v["견적"] or 0) for v in chosen.values())
    missing = [k for k in kinds if k not in chosen]

    # ── PPT 1장
    deck = K.KLIDeck("실사업체 선정", footer=f"{name} 실사업체 선정 보고", template=kli_template())
    picked = ", ".join(f"{k}({chosen[k]['업체명']})" for k in kinds if k in chosen)
    head = f"{picked} 선정, 수수료 합계 {fmt(total, 1)}백만원(VAT 별도)"
    if missing:
        head += f" — {', '.join(missing)} 선정 대기"
    s = deck.slide("제반 실사업체 선정 현황", head, title_lins=deck._title_lins() + 0.45)
    deck.band(s, "선정 결과", top=2.17)
    rows = [["구분", "선정업체", "주요 용역 범위", "수수료(백만원)", "기간", "선정 사유"]]
    for k in kinds:
        v = chosen.get(k)
        scope = " · ".join(g for g, _ in scopes.get(k, {}).get("scope", [])[:3])
        rows.append([dd_label(k), v["업체명"] if v else "[선정 대기]", scope, fmt(v["견적"], 1) if v and v["견적"] is not None else "-",
                     f"{fmt(v['기간'], 0)}주" if v and v["기간"] else "-", (v.get("선정 사유") or "") if v else ""])
    rows.append(["합계", "", "", fmt(total, 1), "", "VAT 별도"])
    deck.table(s, rows, top=2.55, col_widths=[1.0, 1.6, 3.4, 1.25, 0.75, 2.63], row_h=0.31,
               align=["c", "c", "l", "r", "c", "l"])
    y = 2.55 + 0.31 * len(rows) + 0.22
    deck.band(s, "견적 비교", top=y)
    rows = [["구분", "업체", "견적(백만원)", "기간(주)", "제안 요지", "결과"]]
    for k in kinds:
        for v in by[k]:
            rows.append([dd_label(k), v["업체명"], fmt(v["견적"], 1) if v["견적"] is not None else "-",
                         fmt(v["기간"], 0) if v["기간"] else "-", v.get("제안 요지") or "", "선정" if v["선정"] else "-"])
    room = 7.36 - y - 0.36
    rh = max(0.26, min(0.30, room / max(len(rows), 1)))
    fit = int(room / rh)
    if len(rows) > fit:          # 넘치면 선정 업체를 남기고 나머지를 줄여 표시
        keep = [rows[0]] + [r for r in rows[1:] if r[5] == "선정"]
        keep += [r for r in rows[1:] if r[5] != "선정"][: max(0, fit - len(keep) - 1)]
        hidden = len(rows) - len(keep)
        rows = keep + [["", f"외 {hidden}건", "", "", "견적 비교 전체는 딜입력.xlsx 실사업체 시트 참조", ""]]
    deck.table(s, rows, top=y + 0.36, col_widths=[1.0, 1.6, 1.2, 0.9, 5.08, 0.85],
               row_h=rh, align=["c", "c", "r", "c", "l", "c"], bold_cells={(i, 5) for i, r in enumerate(rows) if r[5] == "선정"})
    deck.footnote(s, f"* 자료: 업체별 제안서·견적서({TODAY:%Y.%m.%d} 기준). 수수료는 VAT 별도, 거래 무산 시 보수(Drop fee)는 계약 시 확정")
    if deck.check_min_font():
        raise RuntimeError("10pt 미만 글자")
    ppt = deck.save(outpath(folder, "보고", f"{name}_실사업체선정현황", "pptx"))

    # ── 기안문
    title = f"{name} 매입 검토 관련 실사업체 선정 및 용역계약 체결의 건"
    lines = [f"{name} 매입 검토를 위하여 아래와 같이 실사업체를 선정하고 용역계약을 체결하고자 하오니 승인하여 주시기 바랍니다.",
             "", "- 아 래 -", "", "1. 선정 내역"]
    for i, k in enumerate([k for k in kinds if k in chosen]):
        v = chosen[k]
        lines.append(f"  {'가나다라마바사아'[i]}. {dd_label(k)} : {v['업체명']} (수수료 {fmt(v['견적'], 1)}백만원, VAT 별도, {fmt(v['기간'], 0)}주)")
    lines += [f"  합계 : {fmt(total, 1)}백만원 (VAT 별도)", "", "2. 선정 사유"]
    for k in kinds:
        v = chosen.get(k)
        if v:
            cands = by[k]
            why = v.get("선정 사유") or (f"후보 {len(cands)}개사 중 견적·수행 실적·일정을 종합 고려" if len(cands) > 1 else "단독 제안(수의) — 수행 실적·일정 고려")
            lines.append(f"  - {dd_label(k)} : {why}")
    lines += ["", "3. 용역 개요",
              f"  - 대상 : {name} ({a.get('addr') or a.get('jibun') or ''})",
              f"  - 일정 : 실사 착수 {kdate(_sched(d, 'dd_start'))} → 초안 {kdate(_sched(d, 'dd_draft'))} → 최종 {kdate(_sched(d, 'dd_final'))}",
              "  - 비용 부담 : 매수 비히클 설립 후 비히클 부담(계약 승계), 거래 무산 시 당사 부담분은 용역계약에서 확정",
              "", "4. 승인을 구하는 사항",
              "  - 상기 실사업체 선정 및 용역계약(또는 업무 의뢰서) 체결",
              "", "5. 계약 방법",
              "  - 업체 표준 용역계약서 검토 후 체결(법률 조항은 내부 검토 초안이며 법무 확인 필요)",
              "", "첨부1. 업체별 제안서 및 견적서", "첨부2. 실사업체 선정 현황(견적 비교).  끝."]
    doc = Doc(size=10.5, margin=2.5)
    doc.p(f"제목: {title}", bold=True, size=12)
    doc.p()
    for ln in lines:
        if ln == "- 아 래 -":
            doc.p(ln, align=CENTER, before=6)
        elif ln[:2].rstrip(".").isdigit() and ln[1:2] == ".":
            doc.p(ln, bold=True, before=6)
        else:
            doc.p(ln.strip(), indent=0.4 if ln.startswith("  ") else 0)
    doc.p()
    doc.table([["구분", "업체", "수수료(백만원)", "기간", "비고"]] +
              [[dd_label(k), chosen[k]["업체명"], fmt(chosen[k]["견적"], 1), f"{fmt(chosen[k]['기간'], 0)}주", chosen[k].get("선정 사유") or ""]
               for k in kinds if k in chosen] + [["합계", "", fmt(total, 1), "", "VAT 별도"]], widths=[2.4, 3.2, 2.6, 1.6, 6.2])
    docx = doc.save(outpath(folder, "기안", f"{name}_실사업체선정및용역계약", "docx"), author=a.get("company", "코람코자산신탁"))
    txt = outpath(folder, "기안", f"{name}_실사업체선정_결재본문", "txt")
    txt.write_text(f"제목: {title}\n\n" + "\n".join(lines) + "\n", encoding="utf-8")

    # ── 통보 메일
    mails = []
    me = a.get("manager", "")
    sign = f"\n\n감사합니다.\n\n{me} 드림\n{a.get('company', '코람코자산신탁')} {a.get('team', '').split('/')[-1]} | Tel. {a.get('manager_tel', '')} | {a.get('manager_email', '')}"
    for k in kinds:
        for v in by[k]:
            if not v.get("이메일"):
                continue
            if v["선정"]:
                body = (f"{v['업체명']} {v.get('담당자') or ''}님, 안녕하십니까. {a.get('company', '코람코자산신탁')} {me}입니다.\n\n"
                        f"{name} {dd_label(k)} 용역 제안에 감사드립니다. 내부 검토 결과 귀사를 {dd_label(k)} 수행사로 선정하였음을 알려드립니다"
                        " (내부 결재 완료 후 확정).\n"
                        f"실사 착수 예정일은 {kdate(_sched(d, 'dd_start'))}이며, 용역계약서(또는 업무 의뢰서) 초안과 착수 미팅 가능 일정을 회신 부탁드립니다."
                        + sign)
                subj = f"[{name}] {dd_label(k)} 수행사 선정 안내"
            else:
                body = (f"{v['업체명']} {v.get('담당자') or ''}님, 안녕하십니까. {a.get('company', '코람코자산신탁')} {me}입니다.\n\n"
                        f"{name} {dd_label(k)} 용역에 제안해 주셔서 감사합니다. 검토 결과 이번에는 다른 업체와 진행하게 되었음을 알려드립니다.\n"
                        "좋은 제안에 다시 한번 감사드리며, 다음 기회에 함께할 수 있기를 바랍니다." + sign)
                subj = f"[{name}] {dd_label(k)} 제안 검토 결과 안내"
            mails.append(str(eml.draft(outpath(folder / "발송메일", "메일", f"{dd_label(k)}_{'선정' if v['선정'] else '미선정'}_{v['업체명']}", "eml"),
                                       to=[v["이메일"]], subject=subj, body=body, sender=a.get("manager_email", ""))))
    return {"ppt": str(ppt), "docx": str(docx), "txt": str(txt), "eml": mails, "total": total, "missing": missing}
