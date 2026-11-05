# -*- coding: utf-8 -*-
"""사업일정 → 마일스톤별 체크리스트(엑셀) + 캘린더 등록 파일(.ics).

.ics 는 구글 캘린더(설정 → 가져오기, 딜 전용 캘린더 선택)·아웃룩(파일 열기)에서 한 번에 등록된다.
일정마다 고정 UID를 쓰므로 일정이 바뀌어 다시 가져오면 같은 일정이 갱신된다(구글은 '가져오기'가 덮어쓰기 대신
중복을 만들 수 있으니 딜 전용 캘린더를 지우고 다시 가져오는 방식을 권장).
"""
from __future__ import annotations

import datetime as dt
import hashlib
import pathlib

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .common import WK, kdate, outpath, prev_bizday

FONT = "맑은 고딕"
HEAD = PatternFill("solid", fgColor="D9D9D9")
MS = PatternFill("solid", fgColor="DDEBF7")
EST = Font(name=FONT, size=10, color="C00000", italic=True)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def tasks(d: dict) -> list[dict]:
    out = []
    closing = d.get("closing")
    for m in d["schedule"]:
        if not m["date"]:
            continue
        out.append({"date": m["date"], "ms": m["name"], "task": f"◆ {m['name']}", "owner": "", "milestone": True,
                    "est": m["estimated"], "key": m["key"], "i": -1})
        for i, c in enumerate(m["checklist"]):
            day = prev_bizday(m["date"] + dt.timedelta(days=c["d"]))
            out.append({"date": day, "ms": m["name"], "task": c["task"], "owner": c["owner"], "milestone": False,
                        "est": m["estimated"], "key": m["key"], "i": i})
    for t in out:
        t["dday"] = (t["date"] - closing).days if closing else None
    out.sort(key=lambda t: (t["date"], not t["milestone"], t["ms"]))
    return out


def _uid(code, key, i):
    return hashlib.sha1(f"{code}|{key}|{i}".encode()).hexdigest()[:16] + "@dealflow"


def _esc(s):
    return s.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")


def write_ics(d, ts, path: pathlib.Path):
    code = d["asset"].get("code") or d["asset"].get("name", "DEAL")
    name = d["asset"].get("name", code)
    now = dt.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//dealflow//KO", "CALSCALE:GREGORIAN",
         f"X-WR-CALNAME:{_esc(name)} 딜 일정"]
    for t in ts:
        start = t["date"]
        summ = f"[{name}] " + (t["task"] if t["milestone"] else f"{t['task']} ({t['owner']})")
        if t["est"]:
            summ += " [추정]"
        desc = f"마일스톤: {t['ms']}" + (f"\\n거래종결 D{t['dday']:+d}" if t["dday"] is not None else "")
        L += ["BEGIN:VEVENT", f"UID:{_uid(code, t['key'], t['i'])}", f"DTSTAMP:{now}",
              f"DTSTART;VALUE=DATE:{start:%Y%m%d}", f"DTEND;VALUE=DATE:{start + dt.timedelta(days=1):%Y%m%d}",
              f"SUMMARY:{_esc(summ)}", f"DESCRIPTION:{desc}", "TRANSP:TRANSPARENT"]
        if not t["est"] or t["milestone"]:
            L += ["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_esc(summ)}", "TRIGGER:-PT15H", "END:VALARM"]
        L.append("END:VEVENT")
    L.append("END:VCALENDAR")
    # RFC 5545: 75옥텟 접기
    out = []
    for ln in L:
        b = ln.encode("utf-8")
        while len(b) > 73:
            cut = 73
            while (b[cut] & 0xC0) == 0x80:
                cut -= 1
            out.append(b[:cut].decode("utf-8"))
            b = b" " + b[cut:]
        out.append(b.decode("utf-8"))
    path.write_text("\r\n".join(out) + "\r\n", encoding="utf-8", newline="")
    return path


def write_xlsx(d, ts, path: pathlib.Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "일정표"
    ws.append(["마일스톤", "구분", "일자", "요일", "D-day(종결)", "근거"])
    for m in d["schedule"]:
        dd = (m["date"] - d["closing"]).days if m["date"] and d.get("closing") else None
        ws.append([m["name"], m["group"], m["date"], WK[m["date"].weekday()] if m["date"] else "",
                   f"D{dd:+d}" if dd is not None else "", "[추정] 안양 일정 간격" if m["estimated"] else "입력"])
    for c in ws[1]:
        c.fill, c.font, c.border = HEAD, Font(name=FONT, bold=True, size=10), BOX
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font, c.border = Font(name=FONT, size=10), BOX
        r[2].number_format = "yyyy-mm-dd"
        if r[5].value and "추정" in r[5].value:
            r[5].font = EST
    for c, w in zip("ABCDEF", (38, 8, 12, 6, 12, 22)):
        ws.column_dimensions[c].width = w

    ws = wb.create_sheet("체크리스트")
    ws.append(["일자", "요일", "D-day", "마일스톤", "할 일", "담당", "완료", "비고"])
    for t in ts:
        ws.append([t["date"], WK[t["date"].weekday()], f"D{t['dday']:+d}" if t["dday"] is not None else "", t["ms"],
                   t["task"], t["owner"], "", "[추정]" if t["est"] else ""])
    for c in ws[1]:
        c.fill, c.font, c.border = HEAD, Font(name=FONT, bold=True, size=10), BOX
        c.alignment = Alignment(horizontal="center")
    for r in ws.iter_rows(min_row=2):
        ms = r[4].value.startswith("◆")
        for c in r:
            c.font = Font(name=FONT, size=10, bold=ms)
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if ms:
                c.fill = MS
        r[0].number_format = "yyyy-mm-dd"
    dv = DataValidation(type="list", formula1='"O,진행중,보류"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G2:G{ws.max_row}")
    for c, w in zip("ABCDEFGH", (11, 5, 8, 30, 56, 11, 7, 8)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{ws.max_row}"

    # 주차별(월~일, 안양 추진일정 양식처럼 한 주에 한 줄)
    ws = wb.create_sheet("주차별")
    ws.append(["주 시작(월)", "주요 일정", "할 일"])
    weeks = {}
    for t in ts:
        wk = t["date"] - dt.timedelta(days=t["date"].weekday())
        w = weeks.setdefault(wk, {"ms": [], "task": []})
        (w["ms"] if t["milestone"] else w["task"]).append(
            f"{t['date'].month}/{t['date'].day} " + (t["ms"] if t["milestone"] else t["task"]))
    for wk in sorted(weeks):
        ws.append([wk, "\n".join(weeks[wk]["ms"]), "\n".join(weeks[wk]["task"])])
    for c in ws[1]:
        c.fill, c.font, c.border = HEAD, Font(name=FONT, bold=True, size=10), BOX
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font, c.border = Font(name=FONT, size=10), BOX
            c.alignment = Alignment(vertical="top", wrap_text=True)
        r[0].number_format = "yyyy-mm-dd"
    for c, w in zip("ABC", (12, 44, 70)):
        ws.column_dimensions[c].width = w
    wb.save(path)
    return path


def build(d, folder: pathlib.Path) -> dict:
    ts = tasks(d)
    name = d["asset"].get("name", "딜")
    x = write_xlsx(d, ts, outpath(folder, "일정", f"{name}_사업일정및체크리스트", "xlsx"))
    i = write_ics(d, ts, outpath(folder, "일정", f"{name}_캘린더등록", "ics"))
    soon = [t for t in ts if dt.date.today() <= t["date"] <= dt.date.today() + dt.timedelta(days=14) and not t["milestone"]]
    return {"xlsx": str(x), "ics": str(i), "n_tasks": len(ts),
            "next_2w": [f"{kdate(t['date'])} {t['task']}" for t in soon][:12]}
