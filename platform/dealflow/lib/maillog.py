# -*- coding: utf-8 -*-
"""WGL 메일 자동 분류 → 일자별 이슈·논의사항 기록.

입력: 01_입력자료/메일/ 의 .eml (아웃룩에서 메일을 폴더로 끌어다 놓으면 .msg — extract-msg 패키지가 있으면 읽음)
분류: 보낸·받는 주소를 WGL 이메일(정확히 일치) → 도메인(사내·포털 도메인 제외) 순으로 기관에 매칭
기록: 제목·요지(인용문·서명 제거 후 앞부분)·요청/기한 문장·주제 태그·첨부 → deal.json에 누적(Message-ID 중복 제거)
"""
from __future__ import annotations

import datetime as dt
import email
import email.policy
import hashlib
import html as H
import pathlib
import re
from email.utils import getaddresses, parsedate_to_datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .common import WK, outpath

FREE = {"gmail.com", "naver.com", "daum.net", "hanmail.net", "kakao.com", "nate.com", "hotmail.com", "outlook.com",
        "icloud.com", "yahoo.com", "cu.co.kr"}
TOPICS = {
    "일정": ["일정", "기한", "마감", "D-", "연기", "순연", "까지"],
    "가격·조건": ["가격", "매매대금", "할인", "조정", "T/S", "텀싯", "조건", "금리", "수수료"],
    "계약": ["계약서", "SPA", "주주간", "약정", "MOU", "양해각서", "markup", "마크업", "조항"],
    "실사": ["실사", "DD", "RFI", "보고서", "감정", "물리", "시장조사"],
    "대출": ["대출", "대주", "LTV", "DSCR", "담보", "인출", "LOC"],
    "Equity": ["출자", "투자확약", "종류주", "보통주", "배당", "IRR"],
    "임대차": ["임대차", "임차인", "렌트롤", "보증금", "임대료", "승계"],
    "인허가·설립": ["인가", "등록", "설립", "국토부", "공정위", "등기"],
    "세무·회계": ["세금", "취득세", "부가세", "VAT", "회계", "감사"],
}
ASK = re.compile(r"(요청|부탁|회신|확인\s*부탁|검토\s*부탁|송부\s*부탁|필요합니다|해\s*주시기|주시면|까지|기한|마감)")
QUOTE = re.compile(r"^(-{2,}\s*(Original Message|원본 메시지|Forwarded)|From:|보낸 사람:|보낸사람:|발신:|On .+wrote:|20\d\d[.\-년].+작성:)", re.I)
SIGN = re.compile(r"^(감사합니다|감사드립니다|Best regards|Regards|드림\s*$|올림\s*$)")


def _text(msg) -> str:
    body = msg.get_body(preferencelist=("plain", "html"))
    if body is None:
        return ""
    try:
        t = body.get_content()
    except Exception:
        t = body.get_payload(decode=True).decode("utf-8", "replace")
    if body.get_content_type() == "text/html":
        t = re.sub(r"(?is)<(script|style).*?</\1>", "", t)
        t = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", t)
        t = H.unescape(re.sub(r"<[^>]+>", "", t))
    return t


def clean(t: str) -> list[str]:
    out = []
    for ln in t.replace("\r", "").split("\n"):
        s = ln.strip()
        if QUOTE.match(s) or s.startswith(">"):
            break
        if SIGN.match(s) and out:
            break
        if s:
            out.append(s)
    return out


def read_eml(p: pathlib.Path) -> dict | None:
    msg = email.message_from_bytes(p.read_bytes(), policy=email.policy.default)
    try:
        when = parsedate_to_datetime(msg["Date"]) if msg["Date"] else None
    except Exception:
        when = None
    att = [part.get_filename() for part in msg.iter_attachments() if part.get_filename()]
    return {"id": (msg["Message-ID"] or "").strip() or hashlib.sha1(p.read_bytes()).hexdigest(),
            "date": when, "subject": str(msg["Subject"] or "").strip(),
            "from": getaddresses([str(msg["From"] or "")]), "to": getaddresses([str(msg["To"] or "")]),
            "cc": getaddresses([str(msg["Cc"] or "")]), "body": _text(msg), "att": att, "file": p.name}


def read_msg(p: pathlib.Path) -> dict | None:
    try:
        import extract_msg
    except ImportError:
        return None
    m = extract_msg.Message(str(p))
    try:
        when = m.date if isinstance(m.date, dt.datetime) else None
        split = lambda s: getaddresses([s or ""])
        return {"id": (m.messageId or "").strip() or hashlib.sha1(p.read_bytes()).hexdigest(), "date": when,
                "subject": m.subject or "", "from": split(m.sender), "to": split(m.to), "cc": split(m.cc),
                "body": m.body or "", "att": [a.longFilename or a.shortFilename for a in m.attachments], "file": p.name}
    finally:
        m.close()


def matcher(d):
    by_mail, by_dom = {}, {}
    mine = (d["asset"].get("manager_email") or "").split("@")[-1].lower()
    for r in d["wgl"]:
        for e in re.split(r"[;,/\s]+", r["이메일"] or ""):
            e = e.strip().lower()
            if "@" not in e:
                continue
            by_mail[e] = r
            dom = e.split("@")[-1]
            if dom not in FREE and dom != mine:
                by_dom.setdefault(dom, r)

    def find(addr):
        a = addr.lower()
        if a in by_mail:
            r = by_mail[a]
            return r["기관명"], r["성명"], r["구분"]
        dom = a.split("@")[-1]
        if mine and dom == mine:
            return "사내", "", "매수인(당사)"
        if dom in by_dom:
            r = by_dom[dom]
            return r["기관명"], "", r["구분"]
        return None
    return find


def analyze(rec, find):
    lines = clean(rec["body"])
    people = rec["from"] + rec["to"] + rec["cc"]
    orgs, unknown = [], []
    for nm, ad in people:
        if not ad:
            continue
        hit = find(ad)
        if hit and hit[0] != "사내":
            if hit[0] not in orgs:
                orgs.append(hit[0])
        elif not hit:
            unknown.append(f"{nm} <{ad}>" if nm else ad)
    snd = rec["from"][0] if rec["from"] else ("", "")
    sh = find(snd[1]) if snd[1] else None
    text = rec["subject"] + "\n" + "\n".join(lines)
    tags = [k for k, ws in TOPICS.items() if any(w.lower() in text.lower() for w in ws)]
    sents = re.split(r"(?<=[다요\.\?])\s+|\n", "\n".join(lines))
    asks = [s.strip() for s in sents if ASK.search(s) and 6 < len(s.strip()) < 160][:4]
    gist = " / ".join(x for x in lines[1:7] if len(x) > 4)[:300] if len(lines) > 1 else (lines[0] if lines else "")
    when = rec["date"]
    if when and when.tzinfo:
        when = when.astimezone(dt.timezone(dt.timedelta(hours=9))).replace(tzinfo=None)
    return {"id": rec["id"], "date": when.date().isoformat() if when else "", "time": when.strftime("%H:%M") if when else "",
            "from_org": sh[0] if sh else "[미분류]", "from": snd[0] or snd[1], "orgs": orgs or ["[미분류]"],
            "subject": rec["subject"], "tags": tags, "gist": gist, "asks": asks, "att": rec["att"], "file": rec["file"],
            "unknown": unknown}


def collect(d, folder: pathlib.Path, prior: list[dict]) -> tuple[list[dict], list[str], int]:
    find = matcher(d)
    known = {r["id"]: r for r in prior}
    new, notes = 0, []
    for p in sorted(folder.glob("*")):
        ext = p.suffix.lower()
        if ext not in (".eml", ".msg"):
            continue
        rec = read_eml(p) if ext == ".eml" else read_msg(p)
        if rec is None:
            notes.append(f"{p.name}: .msg 를 읽으려면 'pip install extract-msg' 필요(또는 아웃룩에서 .eml로 저장)")
            continue
        r = analyze(rec, find)
        if r["id"] not in known:
            new += 1
        known[r["id"]] = r
    out = sorted(known.values(), key=lambda r: (r["date"], r["time"]))
    return out, notes, new


def write(d, log, folder: pathlib.Path):
    name = d["asset"].get("name", "딜")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "일자별"
    cols = ["일자", "요일", "시각", "보낸 기관", "보낸 사람", "관련 기관", "제목", "주제", "요지", "요청·기한", "첨부", "파일"]
    ws.append(cols)
    for r in log:
        day = dt.date.fromisoformat(r["date"]) if r["date"] else None
        ws.append([day, WK[day.weekday()] if day else "", r["time"], r["from_org"], r["from"], ", ".join(r["orgs"]),
                   r["subject"], ", ".join(r["tags"]), r["gist"], "\n".join(r["asks"]), ", ".join(r["att"]), r["file"]])
    thin = Side(style="thin", color="BFBFBF")
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.font = Font(name="맑은 고딕", bold=True, size=10)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="맑은 고딕", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = Border(bottom=thin)
        row[0].number_format = "yyyy-mm-dd"
    for c, w in zip("ABCDEFGHIJKL", (11, 5, 6, 16, 12, 22, 40, 16, 60, 44, 24, 20)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{ws.max_row}"
    ws2 = wb.create_sheet("기관별")
    ws2.append(["기관", "메일 수", "최근일", "주요 주제", "최근 요청·기한"])
    agg = {}
    for r in log:
        for o in r["orgs"]:
            a = agg.setdefault(o, {"n": 0, "last": "", "tags": {}, "asks": []})
            a["n"] += 1
            a["last"] = max(a["last"], r["date"])
            for t in r["tags"]:
                a["tags"][t] = a["tags"].get(t, 0) + 1
            a["asks"] += [f"{r['date'][5:]} {x}" for x in r["asks"]]
    for o, a in sorted(agg.items(), key=lambda x: -x[1]["n"]):
        ws2.append([o, a["n"], a["last"], ", ".join(t for t, _ in sorted(a["tags"].items(), key=lambda x: -x[1])[:4]),
                    "\n".join(a["asks"][-3:])])
    for c, w in zip("ABCDE", (22, 8, 11, 30, 80)):
        ws2.column_dimensions[c].width = w
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    unknown = sorted({u for r in log for u in r["unknown"]})
    ws3 = wb.create_sheet("WGL 미등록 주소")
    ws3.append(["주소", "조치"])
    for u in unknown:
        ws3.append([u, "딜입력.xlsx WGL 시트에 추가하면 다음 실행부터 분류"])
    ws3.column_dimensions["A"].width = 50
    x = outpath(folder, "메일로그", f"{name}_WGL메일_이슈로그", "xlsx")
    wb.save(x)

    md = [f"# {name} — WGL 메일 이슈·논의사항 (일자별)", "",
          "자동 분류 결과입니다. 요지·요청 문장은 메일 본문에서 기계적으로 뽑은 것이므로 원문을 확인하세요.", ""]
    cur = None
    for r in reversed(log):
        if r["date"] != cur:
            cur = r["date"]
            md += ["", f"## {cur or '날짜 미상'}"]
        md.append(f"- **[{', '.join(r['orgs'])}]** {r['subject']} — {r['from_org']} {r['from']} {r['time']}"
                  + (f" `{'·'.join(r['tags'])}`" if r["tags"] else ""))
        if r["gist"]:
            md.append(f"  - 요지: {r['gist'][:200]}")
        for a in r["asks"]:
            md.append(f"  - 요청·기한: {a}")
    mp = outpath(folder, "메일로그", f"{name}_이슈로그_일자별", "md")
    mp.write_text("\n".join(md) + "\n", encoding="utf-8")
    return {"xlsx": str(x), "md": str(mp), "n": len(log), "unknown": len(unknown)}
