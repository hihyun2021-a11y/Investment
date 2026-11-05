# -*- coding: utf-8 -*-
"""실사 RFP 5종(시장·재무·물리·감정·법률) Word + 후보 업체별 발송용 임시 메일(.eml)."""
from __future__ import annotations

import datetime as dt
import pathlib

from . import eml
from .common import dd_label, fmt, kdate, load_json, num, outpath, prev_bizday, to_py
from .docx_util import CENTER, GREY, Doc


def _sched(d, key):
    return next((m["date"] for m in d["schedule"] if m["key"] == key), None)


def _deal_rows(d, model):
    a, rr = d["asset"], d["rr"]
    gfa = num(a.get("gfa_m2"))
    occ = rr.get("occupancy")
    top = ", ".join(f"{t['tenant']}({t['share'] * 100:.0f}%)" for t in rr["tenants"][:3]) or "[렌트롤 미접수]"
    rows = [["항목", "내용"],
            ["대상 자산", f"{a.get('name', '[자산명]')} — {a.get('addr') or a.get('jibun') or '[소재지]'}"],
            ["자산 개요", f"{a.get('sector', '')}, 연면적 {fmt(gfa, 1)}㎡({fmt(to_py(gfa), 0)}평), {a.get('floors', '[규모]')}, "
                       f"사용승인 {a.get('approval', '[확인 필요]')}"],
            ["임대 현황", f"임대율 {fmt(occ * 100 if occ is not None else None, 1)}%, WALE {fmt(rr.get('wale'), 1)}년, 주요 임차인 {top}"],
            ["거래 개요", f"매도인 {a.get('owner', '[확인 필요]')}, 매입(희망)가 {fmt(num(a.get('price')), 0)}억원, "
                       f"매수 비히클 {a.get('vehicle', '[협의 중]')}"],
            ["일정", f"우협 선정 {kdate(_sched(d, 'preferred'))} → 실사 착수 {kdate(_sched(d, 'dd_start'))} → "
                    f"거래종결 목표 {kdate(d.get('closing'))}"]]
    return rows


def build_one(d, kind: str, model, folder: pathlib.Path, cfg=None) -> pathlib.Path:
    cfg = cfg or load_json("rfp_scopes.json")
    T = cfg["types"][kind]
    a = d["asset"]
    sector = a.get("sector", "")
    name = a.get("name", "[자산명]")
    sub = lambda s: s.replace("{자산유형}", sector or "부동산")
    send = _sched(d, "dd_rfp") or dt.date.today()
    due = prev_bizday(send + dt.timedelta(days=7))
    start = _sched(d, "dd_start")
    draft = _sched(d, "dd_draft")
    final = _sched(d, "dd_final")
    me = f"{a.get('manager', '[성명]')} {a.get('manager_title', '')}".strip()

    doc = Doc()
    doc.p("※ 작성 안내: 자산·일정 수치는 딜입력.xlsx 기준이며, [ ] 표시는 발송 전 확인·기입할 부분입니다.", size=8.5, color=GREY)
    doc.p("RFP (제안요청서)", bold=True, size=16, align=CENTER, before=6)
    doc.p(f"{name} 매입 관련 {T['title']} 제안 요청", bold=True, size=12, align=CENTER)
    doc.p(f"{a.get('company', '코람코자산신탁')} {a.get('team', '').split('/')[-1]} | {send:%Y년 %m월}", size=9, color=GREY, align=CENTER, after=8)
    doc.table([["제목", f"[RFP] {name} 매입 관련 {T['title']} 제안 요청"],
               ["수신", "[업체명] [담당자]님"], ["발신", f"{a.get('company', '코람코자산신탁')} {me}"]], widths=[2.5, 14], head=False)
    doc.p("[업체명] [담당자]님께,")
    doc.p(f"안녕하십니까. {a.get('company', '코람코자산신탁')} {me}입니다.")
    doc.p(f"당사가 검토 중인 {name} 매입 건과 관련하여 {T['title']}을 의뢰드리고자 합니다. "
          "아래 내용을 검토하시어 제안서와 견적을 보내주시면 감사하겠습니다.")
    doc.h("1. 배경 및 목적")
    doc.bullets([T["purpose"] + "."])
    doc.h("2. 대상 거래 개요")
    doc.table(_deal_rows(d, model), widths=[3.2, 13.3])
    doc.h("3. 용역 범위 (Scope of Work)")
    extra = cfg.get("sector_extra", {}).get(sector, {}).get(kind, [])
    for i, (grp, items) in enumerate(T["scope"], 1):
        doc.p(f"({i}) {grp}", bold=True)
        doc.bullets([sub(x) for x in items])
    if extra:
        doc.p(f"({len(T['scope']) + 1}) {sector} 특화 항목", bold=True)
        doc.bullets(extra)
    doc.h("4. 산출물 및 일정")
    doc.bullets([f"산출물: {T['deliverable']}",
                 f"희망 일정: 착수 {kdate(start)} → 초안 {kdate(draft)} → 최종 {kdate(final)} "
                 f"(표준 소요 {T['weeks']}주 기준, 가능한 일정을 제안 부탁드립니다)",
                 "보고서 이용 범위: 매수 비히클 주주·대주·투자심의 제출, 필요 시 감독당국 제출 — 이에 따른 Reliance Letter·이용 동의 조건을 함께 제시"])
    doc.h("5. 제공 예정 자료 (비밀유지약정 체결 후)")
    doc.bullets(T["data_provide"])
    doc.h("6. 제안서 포함 요청 사항")
    for i, t in enumerate(cfg["common_request"], 1):
        doc.p(f"{i}. {sub(t)}", indent=0.3, hanging=0.4)
    doc.h("7. 제출 안내")
    doc.bullets([f"제출 기한: {kdate(due)} 17:00까지", "제출 방법: 본 메일 회신(PDF)",
                 f"문의: {me}, {a.get('manager_tel', '[전화]')}, {a.get('manager_email', '[이메일]')}"])
    doc.p("본 메일과 첨부 정보는 대외비입니다. 제안 검토 목적으로만 사용해 주시고 제3자에게 공개하지 말아 주시기 바랍니다.",
          size=9, color=GREY, before=6)
    doc.p("감사합니다.")
    doc.p(f"{a.get('manager', '[성명]')} 드림")
    doc.p(f"{a.get('company', '코람코자산신탁')} {a.get('team', '').split('/')[-1]}")
    if kind == "법률":
        doc.p("※ 법률 검토 범위는 내부 초안이며, 최종 범위는 법무법인 제안 내용을 반영해 확정합니다.", size=8.5, color=GREY)
    out = outpath(folder, "RFP", f"{name}_{dd_label(kind)}", "docx")
    doc.save(out, author=a.get("company", "코람코자산신탁"))
    return out, due


def build_all(d, model, folder: pathlib.Path) -> dict:
    cfg = load_json("rfp_scopes.json")
    a = d["asset"]
    name = a.get("name", "[자산명]")
    made = {"rfp": [], "eml": []}
    kinds = [k for k in cfg["types"]]
    for kind in kinds:
        doc, due = build_one(d, kind, model, folder, cfg)
        made["rfp"].append(str(doc))
        for v in [x for x in d["vendors"] if x["실사구분"] == kind and x.get("이메일")]:
            body = (f"{v['업체명']} {v.get('담당자') or ''}님, 안녕하십니까.\n"
                    f"{a.get('company', '코람코자산신탁')} {a.get('manager', '')}입니다.\n\n"
                    f"당사가 검토 중인 {name} 매입 건의 {cfg['types'][kind]['title']} 제안을 요청드립니다.\n"
                    f"첨부 RFP를 검토하시어 {kdate(due)} 17:00까지 제안서와 견적을 회신해 주시면 감사하겠습니다.\n"
                    "자료 제공을 위한 비밀유지약정서가 필요하시면 말씀 주십시오.\n\n"
                    f"감사합니다.\n\n{a.get('manager', '')} 드림\n{a.get('company', '코람코자산신탁')} {a.get('team', '').split('/')[-1]}"
                    f" | Tel. {a.get('manager_tel', '')} | {a.get('manager_email', '')}")
            p = eml.draft(outpath(folder / "발송메일", "메일", f"RFP_{kind}_{v['업체명']}", "eml"),
                          to=[v["이메일"]], subject=f"[RFP] {name} {cfg['types'][kind]['title']} 제안 요청",
                          body=body, attachments=[doc], sender=a.get("manager_email", ""))
            made["eml"].append(str(p))
    return made
