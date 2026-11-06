# -*- coding: utf-8 -*-
"""코스피 상장리츠 공시·금리 일일 보고 메일.

데이터
- 공시·종가·시가총액: 네이버증권(KOSCOM 제공). 모바일 API → 데스크톱 화면 순으로 시도한다.
- 금리: 한국은행 ECOS Open API (ECOS_API_KEY) — 기준금리·CD·CP·국고채·회사채 / FRED (FRED_API_KEY) — 미 연방기금·SOFR·국채 2·10년.
- 재무 지표: OpenDART (DART_API_KEY) — 최근 사업·반기보고서 주석과 배당결정 공시로 시가배당수익률·가중평균 차입금리·LTV·P/NAV (dart_metrics.py)

실행
  python reits_daily.py --check            # 접속·파싱 점검(메일 발송 안 함). 처음 한 번 꼭 실행
  python reits_daily.py --preview          # 수집 후 report_preview.html 저장(발송 안 함)
  python reits_daily.py                    # 수집 → 메일 발송(.env 설정 필요)
  python reits_daily.py --demo             # 인터넷 없이 샘플 데이터로 메일 모양만 확인
  python reits_daily.py --dart             # DART 재무 지표만 계산해 화면에 표시(근거: data/dart_trace/)

외부 패키지 없이 Python 3.9+ 기본 기능만 사용한다. 설정: config.json, 비밀값: env.txt(또는 .env, 깃에 올리지 않음).
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import smtplib
import ssl
import sys
import time
import urllib.parse
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from html.parser import HTMLParser
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
STATE_FILE = HERE / "data" / "state.json"
CACHE_DIR = HERE / "data" / "detail_cache"
KST = dt.timezone(dt.timedelta(hours=9))
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
WEEK = "월화수목금토일"
LOG: list[str] = []


def log(msg: str):
    LOG.append(msg)
    print(msg, file=sys.stderr)


# ────────────────────────────────────────────────────────────── 공통 도구
def load_env():
    """env.txt 또는 .env(KEY=VALUE) → 환경변수. 이미 설정된 값은 덮어쓰지 않는다.
    회사 PC에서 .env 확장자가 AnySign4PC(인증서 암호화 파일)와 연결돼 있어 env.txt를 기본으로 쓴다."""
    p = next((x for x in (HERE / "env.txt", HERE / ".env") if x.exists()), None)
    if p is None:
        return
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def http_get(url: str, *, referer: str | None = None, retries: int = 2, timeout: int = 15) -> tuple[bytes, str]:
    hdr = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "ko-KR,ko;q=0.9"}
    if referer:
        hdr["Referer"] = referer
    last = None
    for i in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers=hdr)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read(), r.headers.get_content_charset() or ""
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(1.0 + i)
    raise RuntimeError(f"GET 실패 {url}: {last}")


def get_json(url: str, **kw):
    raw, cs = http_get(url, **kw)
    return json.loads(raw.decode(cs or "utf-8", errors="replace"))


def get_text(url: str, **kw) -> str:
    raw, cs = http_get(url, **kw)
    for enc in ([cs] if cs else []) + ["utf-8", "euc-kr", "cp949"]:
        try:
            return raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", errors="replace")


class _Text(HTMLParser):
    def __init__(self):
        super().__init__()
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skip += 1
        if tag in ("br", "p", "tr", "div", "li", "h1", "h2", "h3", "table"):
            self.out.append("\n")
        if tag in ("td", "th"):
            self.out.append(" | ")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self.skip:
            self.skip -= 1

    def handle_data(self, d):
        if not self.skip:
            self.out.append(d)


def html_to_text(s: str) -> str:
    p = _Text()
    p.feed(s)
    t = html.unescape("".join(p.out)).replace("\xa0", " ")
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def num(s) -> float | None:
    if s is None:
        return None
    m = re.search(r"-?[\d,]+(?:\.\d+)?", str(s))
    return float(m.group(0).replace(",", "")) if m else None


def parse_dt(s) -> dt.datetime | None:
    if not s:
        return None
    s = str(s).strip()
    for pat in (r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?", r"(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?"):
        m = re.search(pat, s)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            hh = int(m.group(4) or 0)
            mm = int(m.group(5) or 0)
            try:
                return dt.datetime(y, mo, d, hh, mm, tzinfo=KST)
            except ValueError:
                return None
    return None


def won_to_eok(v: float | None) -> str:
    if v is None:
        return "-"
    eok = v / 1e8
    return f"{eok:,.0f}억원" if eok >= 10 else f"{eok:,.1f}억원"


# ────────────────────────────────────────────────────────────── 네이버증권
def resolve_code(name: str) -> str | None:
    q = urllib.parse.quote(name)
    for url in (f"https://ac.stock.naver.com/ac?q={q}&target=stock",
                f"https://m.stock.naver.com/front-api/search/autoComplete?query={q}&target=stock"):
        try:
            data = get_json(url)
        except Exception as e:  # noqa: BLE001
            log(f"[코드검색] {name}: {e}")
            continue
        blob = json.dumps(data, ensure_ascii=False)
        m = re.search(r'"code"\s*:\s*"([0-9A-Z]{6})"[^{}]*?"name"\s*:\s*"' + re.escape(name), blob) or \
            re.search(r'"name"\s*:\s*"' + re.escape(name) + r'"[^{}]*?"code"\s*:\s*"([0-9A-Z]{6})"', blob)
        if m:
            return m.group(1)
    return None


def fetch_quote(code: str) -> dict:
    """종가·등락률·시가총액(억원)."""
    q = {"close": None, "chg": None, "ratio": None, "mcap_eok": None, "asof": None}
    try:
        b = get_json(f"https://m.stock.naver.com/api/stock/{code}/basic")
        q["close"] = num(b.get("closePrice"))
        q["chg"] = num(b.get("compareToPreviousClosePrice"))
        q["ratio"] = num(b.get("fluctuationsRatio"))
        q["asof"] = (b.get("localTradedAt") or "")[:10] or None
        cp = b.get("compareToPreviousPrice") or {}
        if q["ratio"] is not None and str(cp.get("name", "")).upper() in ("FALLING", "LOWER_LIMIT") and q["ratio"] > 0:
            q["ratio"] = -q["ratio"]
    except Exception as e:  # noqa: BLE001
        log(f"[시세] {code} basic: {e}")
    try:
        it = get_json(f"https://m.stock.naver.com/api/stock/{code}/integration")
        for row in it.get("totalInfos", []) or []:
            if row.get("code") == "marketValue" or row.get("key") in ("시총", "시가총액"):
                v = str(row.get("value", ""))
                jo = num(re.search(r"([\d,.]+)조", v).group(1)) if "조" in v else 0
                eok = num(re.search(r"([\d,.]+)억", v).group(1)) if "억" in v else 0
                q["mcap_eok"] = (jo or 0) * 10000 + (eok or 0)
    except Exception as e:  # noqa: BLE001
        log(f"[시세] {code} integration: {e}")
    if q["close"] is None or q["mcap_eok"] is None:  # 데스크톱 화면 대체
        try:
            t = get_text(f"https://finance.naver.com/item/main.naver?code={code}")
            if q["close"] is None:
                m = re.search(r'<p class="no_today">.*?<span class="blind">([\d,]+)</span>', t, re.S)
                q["close"] = num(m.group(1)) if m else None
            if q["mcap_eok"] is None:
                m = re.search(r'id="_market_sum">(.*?)</em>', t, re.S)
                if m:
                    s = re.sub(r"<[^>]+>|\s", "", m.group(1))
                    jo = num(s.split("조")[0]) if "조" in s else 0
                    rest = s.split("조")[-1]
                    q["mcap_eok"] = (jo or 0) * 10000 + (num(rest) or 0)
        except Exception as e:  # noqa: BLE001
            log(f"[시세] {code} main: {e}")
    return q


def _walk_dicts(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from _walk_dicts(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk_dicts(v)


def fetch_disclosures(code: str, pages: int = 2) -> list[dict]:
    """[{id, title, date(datetime), url, src}] 최신순."""
    out: list[dict] = []
    # 1) 모바일 API
    for page in range(1, pages + 1):
        try:
            data = get_json(f"https://m.stock.naver.com/api/stock/{code}/disclosure?pageSize=20&page={page}",
                            referer=f"https://m.stock.naver.com/domestic/stock/{code}/notice")
        except Exception as e:  # noqa: BLE001
            log(f"[공시] {code} 모바일 API p{page}: {e}")
            break
        got = 0
        for d in _walk_dicts(data):
            title = d.get("title") or d.get("subject") or d.get("disclosureTitle")
            when = d.get("datetime") or d.get("date") or d.get("disclosureDate") or d.get("localDateTime")
            if not title or not when:
                continue
            did = str(d.get("disclosureId") or d.get("id") or d.get("no") or d.get("rcpNo") or f"{when}_{title}")
            out.append({"id": did, "title": str(title).strip(), "date": parse_dt(when), "src": "m",
                        "url": f"https://m.stock.naver.com/domestic/stock/{code}/notice/{did}"})
            got += 1
        if not got:
            break
    if out:
        return _dedupe(out)
    # 2) 데스크톱 화면
    for page in range(1, pages + 1):
        try:
            t = get_text(f"https://finance.naver.com/item/news_notice.naver?code={code}&page={page}",
                         referer=f"https://finance.naver.com/item/main.naver?code={code}")
        except Exception as e:  # noqa: BLE001
            log(f"[공시] {code} 데스크톱 p{page}: {e}")
            break
        rows = re.findall(r'<a href="(/item/news_notice_read\.naver\?[^"]+)"[^>]*>(.*?)</a>.*?'
                          r'<td class="date">\s*([\d.]+\s*[\d:]*)\s*</td>', t, re.S)
        for href, title, when in rows:
            href = html.unescape(href)
            no = (re.search(r"no=(\d+)", href) or re.search(r"(\d+)", href)).group(1)
            out.append({"id": no, "title": html_to_text(title), "date": parse_dt(when), "src": "d",
                        "url": "https://finance.naver.com" + href})
        if not rows:
            break
    return _dedupe(out)


def _dedupe(items):
    seen, res = set(), []
    for x in sorted(items, key=lambda z: z["date"] or dt.datetime.min.replace(tzinfo=KST), reverse=True):
        if x["id"] in seen:
            continue
        seen.add(x["id"])
        res.append(x)
    return res


def fetch_detail_text(code: str, d: dict) -> str:
    """공시 본문 텍스트(캐시)."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cf = CACHE_DIR / f"{code}_{re.sub(r'[^0-9A-Za-z_]', '_', d['id'])}.txt"
    if cf.exists():
        return cf.read_text(encoding="utf-8")
    urls = []
    if d["src"] == "m":
        urls += [f"https://m.stock.naver.com/api/stock/{code}/disclosure/{d['id']}",
                 f"https://finance.naver.com/item/news_notice_read.naver?no={d['id']}&code={code}"]
    else:
        urls += [d["url"]]
    text = ""
    for u in urls:
        try:
            raw = get_text(u, referer=f"https://finance.naver.com/item/news_notice.naver?code={code}")
        except Exception as e:  # noqa: BLE001
            log(f"[본문] {code} {d['id']}: {e}")
            continue
        if raw.lstrip().startswith("{"):
            try:
                j = json.loads(raw)
                parts = [str(v) for dd in _walk_dicts(j) for k, v in dd.items()
                         if isinstance(v, str) and len(v) > 40]
                raw = "\n".join(parts)
            except ValueError:
                pass
        # 데스크톱 상세 화면이 DART 뷰어 iframe을 쓰면 따라간다
        m = re.search(r'(https?://dart\.fss\.or\.kr/[^"\']+)', raw)
        text = html_to_text(raw)
        if m and len(text) < 300:
            try:
                text = html_to_text(get_text(html.unescape(m.group(1))))
            except Exception as e:  # noqa: BLE001
                log(f"[본문] DART {e}")
        if len(text) > 100:
            break
    cf.write_text(text, encoding="utf-8")
    time.sleep(0.3)
    return text


# ────────────────────────────────────────────────────────────── 공시 분류·추출
def norm_title(t: str) -> str:
    return re.sub(r"^\s*[\[(]?\s*정정\s*[\])]?\s*", "", t).strip()


def kind_of(title: str) -> str | None:
    t = norm_title(title)
    if "자금차입" in t or "자금 차입" in t:
        return "자금차입"
    if "유상증자결정" in t.replace(" ", "") and "신주발행가액" not in t:
        return "유상증자"
    if "유형자산" in t and "취득" in t:
        return "유형자산 취득"
    if "유형자산" in t and "처분" in t:
        return "유형자산 처분"
    if "증권 취득" in t or "부동산 취득" in t:
        return "유형자산 취득"
    return None


def _after(text: str, labels, span: int = 120) -> str | None:
    for lb in labels:
        m = re.search(lb + r"[^\n|:]*?[:\s]*(?:\|\s*)*([^\n|]{1," + str(span) + "})", text)
        if m:
            v = m.group(1).strip(" |:-").strip()
            if v and v not in ("-", "|"):
                return v
    return None


def _period(text: str) -> str | None:
    s = re.search(r"(?:시작일|차입일|발행일|기간)[^\d]{0,20}(\d{4}[.\-년]\s*\d{1,2}[.\-월]\s*\d{1,2})", text)
    e = re.search(r"(?:종료일|만기일|상환일)[^\d]{0,20}(\d{4}[.\-년]\s*\d{1,2}[.\-월]\s*\d{1,2})", text)
    if not (s and e):
        m = re.search(r"(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})\s*[~∼-]\s*(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})", text)
        if not m:
            return None
        a, b = parse_dt(m.group(1)), parse_dt(m.group(2))
    else:
        a, b = parse_dt(re.sub(r"[년월]", "-", s.group(1))), parse_dt(re.sub(r"[년월]", "-", e.group(1)))
    if not (a and b):
        return None
    days = (b - a).days
    dur = f"{days / 365:.1f}년" if days >= 300 else f"{days}일"
    return f"{a:%Y-%m-%d} ~ {b:%Y-%m-%d} ({dur})"


def extract(kind: str, text: str) -> dict:
    t = text
    if kind == "자금차입":
        amt = None
        m = re.search(r"차입\s*금액[^\d]{0,30}([\d,]{6,})", t) or re.search(r"(?:금액|규모)\s*\(원\)[^\d]{0,30}([\d,]{6,})", t)
        if m:
            amt = won_to_eok(num(m.group(1)))
        else:
            m = re.search(r"차입\s*금액[^\n]{0,30}?([\d,.]+)\s*억", t)
            amt = f"{m.group(1)}억원" if m else None
        rate = _after(t, [r"이자율\s*\(%\)", r"이자율", r"대출\s*금리", r"금리"], 60)
        if rate:
            rate = re.sub(r"\s+", " ", rate)[:60]
        return {"형태": _after(t, [r"차입\s*형태", r"차입\s*방법", r"사채의\s*종류", r"차입\s*종류"], 30),
                "금액": amt, "금리": rate, "기간": _period(t),
                "차입처": _after(t, [r"차입처", r"대주"], 40)}
    if kind == "유상증자":
        m = re.search(r"(?:신주\s*발행가액|발행가액|확정발행가)[^\d]{0,30}([\d,]+)\s*원?", t)
        price = f"{m.group(1)}원" if m else None
        if price:
            price += " (확정)" if "확정" in t[:2000] else " (예정)"
        sh = re.search(r"(?:보통주식|신주의\s*종류와\s*수)[^\d]{0,30}([\d,]{4,})\s*주?", t)
        tot = re.search(r"(?:시설자금|운영자금|채무상환자금|타법인\s*증권\s*취득자금|기타자금)[^\n]{0,80}", t)
        size = None
        if m and sh:
            size = f"{num(m.group(1)) * num(sh.group(1)) / 1e8:,.0f}억원 ({num(sh.group(1)) / 1e4:,.0f}만주)"
        return {"발행가액": price, "규모": size or (tot.group(0)[:40] if tot else None)}
    if kind.startswith("유형자산"):
        nm = _after(t, [r"자산\s*명", r"물건\s*명", r"취득\s*물건", r"처분\s*물건", r"자산구분", r"대상\s*자산"], 60)
        m = re.search(r"(?:취득|처분)\s*(?:금액|가액)[^\d]{0,30}([\d,]{6,})", t)
        dt_ = re.search(r"(?:취득|처분)\s*예정\s*일자?[^\d]{0,20}(\d{4}[.\-년]\s*\d{1,2}[.\-월]\s*\d{1,2})", t)
        return {"물건": nm, "금액": won_to_eok(num(m.group(1))) if m else None,
                "예정일": dt_.group(1).replace(" ", "") if dt_ else None}
    return {}


# ────────────────────────────────────────────────────────────── ECOS 금리
def fetch_rates(key: str | None, today: dt.date) -> list[dict]:
    res = []
    if not key:
        log("[금리] ECOS_API_KEY 없음 — 금리 표는 '[미입력]'으로 표시")
        return [{"name": r["name"], "value": None, "d1": None, "w1": None, "date": None} for r in CONFIG["rates"]]
    start = (today - dt.timedelta(days=30)).strftime("%Y%m%d")
    end = today.strftime("%Y%m%d")
    for r in CONFIG["rates"]:
        url = (f"https://ecos.bok.or.kr/api/StatisticSearch/{key}/json/kr/1/100/"
               f"{r['stat']}/D/{start}/{end}/{r['item']}")
        row = {"name": r["name"], "value": None, "d1": None, "w1": None, "date": None}
        try:
            j = get_json(url)
            rows = (j.get("StatisticSearch") or {}).get("row") or []
            if not rows:
                log(f"[금리] {r['name']}: 자료 없음 {json.dumps(j, ensure_ascii=False)[:150]}")
            vals = [(x["TIME"], float(x["DATA_VALUE"])) for x in rows if x.get("DATA_VALUE") not in (None, "", "-")]
            vals.sort()
            if vals:
                row["date"], row["value"] = vals[-1]
                if len(vals) >= 2:
                    row["d1"] = (vals[-1][1] - vals[-2][1]) * 100
                if len(vals) >= 6:
                    row["w1"] = (vals[-1][1] - vals[-6][1]) * 100
        except Exception as e:  # noqa: BLE001
            log(f"[금리] {r['name']}: {e}")
        res.append(row)
        time.sleep(0.2)
    return res


def fetch_fred(key: str | None, today: dt.date) -> list[dict]:
    """미국 금리(FRED). 값이 '.'인 날(휴일)은 건너뛴다."""
    items = CONFIG.get("fred") or []
    if not key:
        if items:
            log("[금리] FRED_API_KEY 없음 — 해외 금리 표는 '[미입력]'으로 표시")
        return [{"name": r["name"], "value": None, "d1": None, "w1": None, "date": None} for r in items]
    start = (today - dt.timedelta(days=30)).isoformat()
    res = []
    for r in items:
        row = {"name": r["name"], "value": None, "d1": None, "w1": None, "date": None}
        url = ("https://api.stlouisfed.org/fred/series/observations?"
               + urllib.parse.urlencode({"series_id": r["series"], "api_key": key, "file_type": "json",
                                         "observation_start": start}))
        try:
            obs = get_json(url).get("observations") or []
            vals = [(o["date"].replace("-", ""), float(o["value"])) for o in obs if o.get("value") not in (None, "", ".")]
            vals.sort()
            if vals:
                row["date"], row["value"] = vals[-1]
                if len(vals) >= 2:
                    row["d1"] = (vals[-1][1] - vals[-2][1]) * 100
                if len(vals) >= 6:
                    row["w1"] = (vals[-1][1] - vals[-6][1]) * 100
            else:
                log(f"[금리] FRED {r['series']}: 자료 없음")
        except Exception as e:  # noqa: BLE001
            log(f"[금리] FRED {r['series']}: {e}")
        res.append(row)
        time.sleep(0.2)
    return res


# ────────────────────────────────────────────────────────────── 수집
def collect(now: dt.datetime, state: dict) -> dict:
    load_env()
    today = now.date()
    since_major = now - dt.timedelta(days=CONFIG["major_window_days"])
    since_new = now - dt.timedelta(days=CONFIG["new_window_days"])
    last_run = parse_dt(state.get("last_run")) or (now - dt.timedelta(days=1))
    week_start = (now - dt.timedelta(days=now.weekday())).replace(hour=0, minute=0)
    reits, major = [], {"자금차입": [], "유상증자": [], "유형자산 취득": [], "유형자산 처분": []}
    for r in CONFIG["reits"]:
        code = r.get("code") or state.get("codes", {}).get(r["name"]) or resolve_code(r["name"])
        if not code:
            log(f"[종목] {r['name']}: 코드 확인 불가 — config.json에 code를 넣어 주세요")
            continue
        state.setdefault("codes", {})[r["name"]] = code
        q = fetch_quote(code)
        ds = fetch_disclosures(code)
        for d in ds:
            d["new"] = bool(d["date"] and d["date"] > last_run)
            d["week"] = bool(d["date"] and d["date"] >= week_start)
        reits.append({"name": r["name"], "code": code, **q, "disc": ds})
        for d in ds:
            if not d["date"] or d["date"] < since_major:
                continue
            k = kind_of(d["title"])
            if not k:
                continue
            info = extract(k, fetch_detail_text(code, d))
            major[k].append({"name": r["name"], "code": code, **d, **info})
        time.sleep(0.3)
    # 같은 건의 정정공시는 최신 1건만 (종목 + 정정 뗀 제목 + 금액·기간이 같으면 같은 건)
    for k, lst in major.items():
        seen, keep = set(), []
        for x in sorted(lst, key=lambda z: z["date"], reverse=True):
            sig = (x["code"], norm_title(x["title"]), x.get("금액"), x.get("기간"), x.get("발행가액") if k == "유상증자" else None)
            if not (x.get("금액") or x.get("기간") or x.get("발행가액")):
                sig = sig + (x["date"].date(),)
            if sig in seen:
                continue
            seen.add(sig)
            keep.append(x)
        major[k] = keep
    reits.sort(key=lambda z: z.get("mcap_eok") or 0, reverse=True)
    new_list = []
    for r in reits:
        cand = [d for d in r["disc"] if d["date"] and d["date"] >= since_new
                and not any(ex in d["title"] for ex in CONFIG["exclude_new_titles"])]
        if cand:
            new_list.append({"name": r["name"], **cand[0]})
    new_list.sort(key=lambda z: z["date"], reverse=True)
    price_date = next((r["asof"] for r in reits if r.get("asof")), None)
    rates = fetch_rates(os.environ.get("ECOS_API_KEY"), today)
    return {"now": now, "reits": reits, "major": major, "new": new_list, "rates": rates,
            "fred": fetch_fred(os.environ.get("FRED_API_KEY"), today), "price_date": price_date,
            "since_major": since_major, "since_new": since_new, "dart": collect_dart(reits, rates)}


def _cd(rates) -> float | None:
    return next((r["value"] for r in rates or [] if "CD" in r["name"] and r["value"] is not None), None)


def collect_dart(reits: list[dict], rates=None) -> list[dict] | None:
    """DART 재무 지표(키가 없거나 꺼져 있으면 None)."""
    key = os.environ.get("DART_API_KEY")
    if not CONFIG.get("dart", {}).get("enabled", True):
        return None
    if not key:
        log("[DART] DART_API_KEY 없음 — 재무 지표 표는 생략")
        return None
    import dart_metrics
    try:
        return dart_metrics.collect_all(key, reits, CONFIG.get("dart_override") or {}, log, _cd(rates))
    except Exception as e:  # noqa: BLE001
        log(f"[DART] {e}")
        return None


# ────────────────────────────────────────────────────────────── HTML
CSS_T = "border-collapse:collapse;width:100%;font-size:13px;margin:4px 0 14px 0;"
CSS_TH = "background:#D9D9D9;border-top:1px solid #404040;border-bottom:1px solid #BFBFBF;padding:5px 6px;text-align:center;font-weight:bold;"
CSS_TD = "border-bottom:1px solid #BFBFBF;padding:5px 6px;vertical-align:middle;"
CSS_H = "font-size:15px;font-weight:bold;color:#fff;background:#345B86;padding:5px 10px;margin:18px 0 6px 0;"
CSS_H2 = "font-size:13px;font-weight:bold;color:#002060;margin:8px 0 2px 2px;"
ORANGE = "#E46C0A"


def _d(x: dt.datetime | None) -> str:
    return x.strftime("%Y-%m-%d") if x else "-"


def _wd(x: dt.date) -> str:
    return f"{x:%Y-%m-%d}({WEEK[x.weekday()]})"


def table(head, rows, aligns=None) -> str:
    aligns = aligns or ["center"] * len(head)
    h = "".join(f'<th style="{CSS_TH}">{html.escape(c)}</th>' for c in head)
    body = []
    for r in rows:
        cells = []
        for c, a in zip(r, aligns):
            style, val = CSS_TD + f"text-align:{a};", c
            if isinstance(c, tuple):
                val, extra = c
                style += extra
            cells.append(f'<td style="{style}">{val}</td>')
        body.append("<tr>" + "".join(cells) + "</tr>")
    return f'<table style="{CSS_T}"><tr>{h}</tr>{"".join(body)}</table>'


def link(d) -> str:
    return f'<a href="{html.escape(d["url"])}" style="color:#1F1F1F;">{html.escape(d["title"])}</a>'


def bp(v):
    if v is None:
        return "-"
    color = "#C00000" if v > 0 else ("#0070C0" if v < 0 else "#646464")
    return (f'<span style="color:{color};">{"▲" if v > 0 else ("▼" if v < 0 else "")}'
            f'{abs(v):.1f}bp</span>')


def render(R: dict) -> tuple[str, str]:
    now = R["now"]
    n_new = sum(1 for r in R["reits"] for d in r["disc"] if d.get("new"))
    subj = f"[상장리츠 공시·금리] {now:%m/%d} {now:%H:%M} 기준 새 공시 {n_new}건"
    pdate = R["price_date"] or "-"
    out = [f'<div style="font-family:\'맑은 고딕\',Malgun Gothic,sans-serif;color:#1F1F1F;max-width:960px;">',
           f'<div style="font-size:18px;font-weight:bold;color:#002060;">{html.escape(CONFIG["title"])}</div>',
           f'<div style="font-size:12px;color:#646464;margin:2px 0 8px 0;">공시 기준 {_wd(now.date())} {now:%H:%M} · '
           f'종가·시가총액 기준 {pdate} · {len(R["reits"])}개 종목 · 출처: 네이버증권(KOSCOM), 한국은행 ECOS, FRED</div>']

    # 0. 금리
    def rate_rows(lst):
        rows = []
        for r in lst:
            v = f"{r['value']:.3f}%" if r["value"] is not None else "[미입력]"
            d = r["date"] or "-"
            if d != "-" and len(d) == 8:
                d = f"{d[:4]}-{d[4:6]}-{d[6:]}"
            rows.append([html.escape(r["name"]), (v, "font-weight:bold;"), bp(r["d1"]), bp(r["w1"]), d])
        return rows
    out.append(f'<div style="{CSS_H}">1. 주요 금리</div>')
    heads, al = ["구분", "금리", "전일 대비", "1주 전 대비", "기준일"], ["left", "right", "right", "right", "center"]
    out.append(f'<div style="{CSS_H2}">1) 국내 (한국은행 ECOS)</div>')
    out.append(table(heads, rate_rows(R["rates"]), al))
    if R.get("fred"):
        out.append(f'<div style="{CSS_H2}">2) 미국 (FRED, 미국 현지 기준일)</div>')
        out.append(table(heads, rate_rows(R["fred"]), al))

    # 2. 재무 지표 (DART)
    sec = 2
    if R.get("dart"):
        out.append(f'<div style="{CSS_H}">2. 리츠별 재무 지표 (DART 최근 사업·반기보고서)</div>')
        out.append(dart_table(R))
        sec = 3

    # 3. 주요사항 공시
    out.append(f'<div style="{CSS_H}">{sec}. 주요사항 공시</div>')
    out.append(f'<div style="font-size:12px;color:#646464;">최근 {CONFIG["major_window_days"]}일({_d(R["since_major"])}~) · '
               f'같은 건의 정정공시는 최신 것만 표시 · 본문에서 자동 추출(빈칸은 링크에서 확인)</div>')
    spec = {
        "자금차입": (["종목", "공시일", "차입형태", "대출금액", "대출금리", "대출기간", "공시"],
                 lambda x: [x["name"], _d(x["date"]), x.get("형태") or "-", x.get("금액") or "-",
                            x.get("금리") or "-", x.get("기간") or "-", link(x)]),
        "유상증자": (["종목", "공시일", "발행가액", "유상증자규모", "공시"],
                 lambda x: [x["name"], _d(x["date"]), x.get("발행가액") or "-", x.get("규모") or "-", link(x)]),
        "유형자산 취득": (["종목", "공시일", "취득물건명", "취득금액", "예정일", "공시"],
                      lambda x: [x["name"], _d(x["date"]), x.get("물건") or "-", x.get("금액") or "-",
                                 x.get("예정일") or "-", link(x)]),
        "유형자산 처분": (["종목", "공시일", "처분물건명", "처분금액", "예정일", "공시"],
                      lambda x: [x["name"], _d(x["date"]), x.get("물건") or "-", x.get("금액") or "-",
                                 x.get("예정일") or "-", link(x)]),
    }
    i = 0
    for k, (head, fn) in spec.items():
        lst = R["major"].get(k) or []
        if not lst:
            continue
        i += 1
        out.append(f'<div style="{CSS_H2}">{i}) {k} {len(lst)}건</div>')
        rows = [[(c, f"color:{ORANGE};font-weight:bold;") if x.get("new") and j == 0 else c
                 for j, c in enumerate(fn(x))] for x in lst]
        out.append(table(head, rows))
    if not i:
        out.append('<div style="font-size:13px;margin:6px 0 12px 2px;">해당 기간 주요사항 공시 없음</div>')

    # 2. 리츠별 공시
    out.append(f'<div style="{CSS_H}">{sec + 1}. 리츠별 공시사항</div>')
    out.append(f'<div style="{CSS_H2}">1) 새 공시 {len(R["new"])}건 · 최근 {CONFIG["new_window_days"]}일({_d(R["since_new"])}~) · '
               f'리츠별 최신 1건 · {"·".join(CONFIG["exclude_new_titles"])} 제외</div>')
    if R["new"]:
        out.append(table(["종목", "발표일자", "공시 제목"],
                         [[x["name"], _d(x["date"]), link(x)] for x in R["new"]], ["center", "center", "left"]))
    out.append(f'<div style="{CSS_H2}">2) 종목별 최근 공시 {CONFIG["recent_per_reit"]}건</div>')
    out.append(f'<div style="font-size:12px;color:#646464;">굵은 글씨 = 금주 공시 · '
               f'<span style="color:{ORANGE};">주황 글씨</span> = 직전 보고 이후 새 공시 · 바탕색 = 금주 공시가 있는 종목</div>')
    rows = []
    for r in R["reits"]:
        recent = r["disc"][:CONFIG["recent_per_reit"]]
        has_week = any(d.get("week") for d in recent)
        bg = "background:#FFF7E6;" if has_week else ""
        ratio = r.get("ratio")
        rc = "#C00000" if (ratio or 0) > 0 else ("#0070C0" if (ratio or 0) < 0 else "#646464")
        arrow = "▲" if (ratio or 0) > 0 else ("▼" if (ratio or 0) < 0 else "")
        price = (f'{r["close"]:,.0f}원<br><span style="color:{rc};font-size:12px;">{arrow}{abs(ratio):.2f}%</span>'
                 if r.get("close") is not None and ratio is not None else "-")
        mc = f'{r["mcap_eok"]:,.0f}억원' if r.get("mcap_eok") else "-"
        newtag = f' <span style="color:{ORANGE};font-weight:bold;">새 공시</span>' if any(d.get("new") for d in recent) else ""
        name = f'<b>{html.escape(r["name"])}</b><br><span style="font-size:12px;color:#646464;">{r["code"]} / {mc}</span>{newtag}'
        items = []
        for d in recent:
            st = ("font-weight:bold;" if d.get("week") else "") + (f"color:{ORANGE};" if d.get("new") else "")
            url, title = html.escape(d["url"]), html.escape(d["title"])
            items.append(f'<div>{_d(d["date"])} <a href="{url}" style="color:#1F1F1F;{st}">{title}</a></div>')
        rows.append([(name, bg), (price, bg + "text-align:right;"), ("".join(items) or "-", bg + "text-align:left;font-size:12px;")])
    out.append(table(["종목", "전일 종가", "최근 공시"], rows, ["left", "right", "left"]))
    if LOG:
        out.append('<div style="font-size:11px;color:#8C8C8C;margin-top:10px;">수집 경고 ' + str(len(LOG)) +
                   '건: ' + html.escape(" / ".join(LOG[:8])) + ("…" if len(LOG) > 8 else "") + '</div>')
    out.append('<div style="font-size:11px;color:#8C8C8C;margin-top:8px;">본 메일은 공개 정보를 자동 취합한 참고자료이며, '
               '세부 내용은 공시 원문을 확인하시기 바랍니다.</div></div>')
    return subj, "".join(out)


def dart_table(R: dict) -> str:
    by = {m["name"]: m for m in R["dart"]}
    f = lambda v, d=2, u="": "-" if v is None else f"{v:,.{d}f}{u}"  # noqa: E731
    rows = []
    for r in R["reits"]:
        m = by.get(r["name"])
        if not m:
            continue
        rate = f(m.get("rate"), 2, "%") + (" [추정]" if m.get("rate_est") and m.get("rate") is not None else "")
        div = f(m.get("yield"), 2, "%")
        if m.get("dps"):
            div += f'<br><span style="font-size:11px;color:#646464;">{m["dps"]:,.0f}원/{m.get("months") or "?"}개월</span>'
        ltv = f(m.get("ltv"), 1, "%")
        nav = f(m.get("nav_eok"), 0)
        pn = f(m.get("pnav"), 2, "x")
        rep = html.escape(m.get("report") or "-")
        note = html.escape(" · ".join(m.get("flags") or [])) or "-"
        rows.append([f'<b>{html.escape(r["name"])}</b>', (div, "text-align:right;"), (rate, "text-align:right;"),
                     (ltv, "text-align:right;"), (nav, "text-align:right;"), (pn, "text-align:right;font-weight:bold;white-space:nowrap;"),
                     (f'<span style="font-size:11px;">{rep}</span>', "text-align:left;"),
                     (f'<span style="font-size:11px;color:#646464;">{note}</span>', "text-align:left;")])
    head = ["종목", "시가배당수익률(연환산)", "가중평균 차입금리", "LTV", "NAV(억원)", "P/NAV", "기준 보고서", "비고"]
    note = ('<div style="font-size:11px;color:#646464;margin:-8px 0 12px 0;line-height:1.5;">'
            '시가배당수익률 = 최근 부동산투자회사금전배당결정 1주당 배당금 × (12 ÷ 결산기간 개월) ÷ 전일 종가 · '
            '가중평균 차입금리 = 「재무제표 주석」 차입금·사채 당기말 장부금액 가중평균(변동금리는 CD(91일)+가산금리 [추정]) · '
            'LTV = 차입금·사채 ÷ (종속·관계기업투자 장부금액 + 투자부동산 공정가치(서열체계, 없으면 장부금액)) · '
            'NAV = 위 자산가치 − 차입금·사채 − 재무상태표 임대보증금 · P/NAV = 시가총액(보통주) ÷ NAV. '
            '자동 추출값이므로 의사결정에 쓰기 전 보고서 원문을 확인하세요(근거: data/dart_trace/).</div>')
    return table(head, rows, ["left", "right", "right", "right", "right", "right", "left", "left"]) + note


# ────────────────────────────────────────────────────────────── 발송·상태
def _send_outlook(subject: str, html_body: str, to: list[str]):
    """회사 PC의 Outlook으로 발송(PowerShell COM). Gmail·외부 SMTP가 막힌 경우용."""
    import subprocess
    import tempfile
    tmp = Path(tempfile.mkdtemp(prefix="reits_"))
    (tmp / "subject.txt").write_text(subject, encoding="utf-8")
    (tmp / "body.html").write_text(html_body, encoding="utf-8")
    (tmp / "to.txt").write_text("; ".join(to), encoding="utf-8")
    ps = tmp / "send.ps1"
    ps.write_text("""
$ErrorActionPreference = 'Stop'
$d = Split-Path -Parent $MyInvocation.MyCommand.Path
$o = New-Object -ComObject Outlook.Application
$m = $o.CreateItem(0)
$m.To = Get-Content -Raw -Encoding UTF8 (Join-Path $d 'to.txt')
$m.Subject = (Get-Content -Raw -Encoding UTF8 (Join-Path $d 'subject.txt')).Trim()
$m.HTMLBody = Get-Content -Raw -Encoding UTF8 (Join-Path $d 'body.html')
$m.Send()
""", encoding="utf-8-sig")
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise RuntimeError("Outlook 발송 실패: " + (r.stderr or r.stdout).strip()[:400] +
                           " — Outlook이 설치·로그인돼 있는지 확인하세요.")


def _send_gmail(subject: str, html_body: str, to: list[str]):
    user = os.environ.get("GMAIL_USER", "").strip()
    pw = re.sub(r"\s+", "", os.environ.get("GMAIL_APP_PASSWORD", ""))   # 구글이 4자리씩 띄워 보여 주는 공백 제거
    if not (user and pw):
        raise SystemExit("GMAIL_USER / GMAIL_APP_PASSWORD 를 env.txt 에 넣어 주세요(또는 MAIL_METHOD=outlook).")
    if len(pw) != 16 or not pw.isascii():
        raise SystemExit(f"GMAIL_APP_PASSWORD 가 16자리 앱 비밀번호가 아닙니다(현재 {len(pw)}자). "
                         "구글 계정 → 보안 → 앱 비밀번호에서 만든 16자리를 넣어 주세요.")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((os.environ.get("MAIL_FROM_NAME", "상장리츠 일일보고"), user))
    msg["To"] = ", ".join(to)
    msg.attach(MIMEText(html_to_text(html_body), "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    errors = []
    # PC 이름에 한글이 있으면 EHLO 단계에서 ascii 오류가 나므로 접속 이름을 고정한다
    for port in (465, 587):
        try:
            if port == 465:
                s = smtplib.SMTP_SSL("smtp.gmail.com", 465, local_hostname="localhost",
                                     context=ssl.create_default_context(), timeout=30)
            else:
                s = smtplib.SMTP("smtp.gmail.com", 587, local_hostname="localhost", timeout=30)
                s.ehlo()
                s.starttls(context=ssl.create_default_context())
                s.ehlo()
            with s:
                s.login(user, pw)
                s.sendmail(user, to, msg.as_string())
            return
        except smtplib.SMTPAuthenticationError as e:
            raise SystemExit(f"Gmail 로그인 거부({e.smtp_code}): 주소·앱 비밀번호를 확인하세요. "
                             "2단계 인증이 켜져 있어야 앱 비밀번호를 쓸 수 있습니다.") from e
        except Exception as e:  # noqa: BLE001  (연결 끊김·차단·시간 초과 → 다음 포트)
            errors.append(f"{port}: {type(e).__name__} {e}")
    raise SystemExit("Gmail 서버 연결이 끊겼습니다(" + " / ".join(errors) + "). 회사 보안 프로그램·방화벽이 외부 메일 "
                     "발송을 막는 경우가 많습니다. env.txt 에 MAIL_METHOD=outlook 을 넣으면 회사 Outlook으로 보냅니다.")


def send(subject: str, html_body: str):
    load_env()
    to = [x.strip() for x in re.split(r"[,;]", os.environ.get("MAIL_TO", "")) if x.strip()]
    if not to:
        raise SystemExit("MAIL_TO(받는 사람)를 env.txt 에 넣어 주세요.")
    method = os.environ.get("MAIL_METHOD", "gmail").strip().lower()
    if method == "outlook":
        _send_outlook(subject, html_body, to)
    else:
        _send_gmail(subject, html_body, to)


def load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(state: dict, now: dt.datetime):
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    state["last_run"] = now.isoformat()
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")


# ────────────────────────────────────────────────────────────── 점검·데모
def check():
    load_env()
    code = "357120"
    print("1) 시세:", fetch_quote(code))
    ds = fetch_disclosures(code)
    print(f"2) 공시 목록: {len(ds)}건 (출처 {'모바일 API' if ds and ds[0]['src'] == 'm' else '데스크톱 화면' if ds else '없음'})")
    for d in ds[:5]:
        print("   ", _d(d["date"]), d["title"], d["id"])
    tgt = next((d for d in ds if kind_of(d["title"])), None)
    if tgt:
        txt = fetch_detail_text(code, tgt)
        print(f"3) 본문 추출: {tgt['title']} → {len(txt)}자")
        print("   ", extract(kind_of(tgt["title"]), txt))
    else:
        print("3) 본문 추출: 최근 목록에 주요사항 공시가 없어 건너뜀")
    rates = fetch_rates(os.environ.get("ECOS_API_KEY"), dt.datetime.now(KST).date())
    print("4) 국내 금리(ECOS):", [(r["name"], r["value"], r["date"]) for r in rates])
    fr = fetch_fred(os.environ.get("FRED_API_KEY"), dt.datetime.now(KST).date())
    print("   미국 금리(FRED):", [(r["name"], r["value"], r["date"]) for r in fr])
    if os.environ.get("DART_API_KEY"):
        q = fetch_quote(code)
        res = collect_dart([{"name": "코람코라이프인프라리츠", "code": code, "close": q["close"], "mcap_eok": q["mcap_eok"]}], rates)
        print("   DART 재무 지표:", res)
        print("   근거:", HERE / "data" / "dart_trace")
    else:
        print("   DART 재무 지표: DART_API_KEY 미설정(env.txt)")
    method = os.environ.get("MAIL_METHOD", "gmail").lower()
    need = ("MAIL_TO",) if method == "outlook" else ("GMAIL_USER", "GMAIL_APP_PASSWORD", "MAIL_TO")
    print(f"5) 메일 설정({method}):", "OK" if all(os.environ.get(k) for k in need) else "미설정(env.txt)")
    if LOG:
        print("\n경고:", *LOG, sep="\n - ")


def demo() -> dict:
    """인터넷 없이 메일 모양 확인용 샘플(값은 예시)."""
    now = dt.datetime.now(KST).replace(hour=8, minute=0, second=0, microsecond=0)
    D = lambda s: parse_dt(s)  # noqa: E731
    mk = lambda t, s, new=False: {"id": t + s, "title": t, "date": D(s), "url": "https://finance.naver.com", "src": "d",  # noqa: E731
                                  "new": new, "week": D(s) >= (now - dt.timedelta(days=now.weekday())).replace(hour=0)}
    reits = [
        {"name": "SK리츠", "code": "395400", "close": 4915, "ratio": -6.74, "mcap_eok": 14795,
         "disc": [mk("기업설명회(IR) 개최(안내공시)", "2026-10-13"), mk("부동산투자회사 자금차입", "2026-10-01")]},
        {"name": "코람코라이프인프라리츠", "code": "357120", "close": 3540, "ratio": -2.21, "mcap_eok": 3446,
         "disc": [mk("유형자산 취득결정(종속회사의 주요경영사항)", "2026-10-29"), mk("타법인주식및출자증권취득결정", "2026-10-26")]},
        {"name": "NH올원리츠", "code": "400760", "close": 2400, "ratio": -3.42, "mcap_eok": 1318,
         "disc": [mk("해산사유 발생(종속회사의 주요경영사항)", now.strftime("%Y-%m-%d"), True),
                  mk("(정정)부동산투자회사 자금차입", now.strftime("%Y-%m-%d"), True)]},
    ]
    major = {"자금차입": [
        {"name": "NH올원리츠", **mk("(정정)부동산투자회사 자금차입", now.strftime("%Y-%m-%d"), True),
         "형태": "주식담보대출", "금액": "450억원", "금리": "5.5%(고정금리)", "기간": "2026-05-13 ~ 2028-05-15 (2.0년)"},
        {"name": "코람코라이프인프라리츠", **mk("부동산투자회사 자금차입", "2026-10-16"),
         "형태": "사모사채", "금액": "140억원", "금리": "연 6.0%", "기간": "2026-10-16 ~ 2028-10-16 (2.0년)"}],
        "유상증자": [{"name": "삼성FN리츠", **mk("(정정)유상증자결정", "2026-10-30"), "발행가액": "4,755원 (확정)", "규모": "903억원 (1,900만주)"}],
        "유형자산 취득": [{"name": "코람코라이프인프라리츠", **mk("유형자산 취득결정(종속회사의 주요경영사항)", "2026-10-29"),
                       "물건": "안양물류센터", "금액": "2,700억원", "예정일": "2026-10-29"}],
        "유형자산 처분": []}
    rates = [{"name": r["name"], "value": None, "d1": None, "w1": None, "date": None} for r in CONFIG["rates"]]
    for r, (v, d1, w1) in zip(rates, [(2.50, 0, 0), (3.27, 0.0, 1.0), (3.45, -1.0, 2.0), (2.95, 2.1, 5.3),
                                     (3.05, 1.8, 4.9), (3.20, 1.5, 4.0), (3.55, 1.9, 4.7), (9.20, 0.3, 1.2)]):
        r.update(value=v, d1=d1, w1=w1, date=(now - dt.timedelta(days=1)).strftime("%Y%m%d"))
    fred = [{"name": r["name"], "value": v, "d1": d1, "w1": w1, "date": (now - dt.timedelta(days=2)).strftime("%Y%m%d")}
            for r, (v, d1, w1) in zip(CONFIG.get("fred", []), [(3.88, 0, -2.0), (3.92, 1.0, -3.0), (3.55, -2.4, -6.1), (4.12, -1.8, -4.4)])]
    new = [{"name": x["name"], **x["disc"][0]} for x in reits]
    LOG.append("[예시] --demo 샘플 데이터입니다. 수치는 실제와 다릅니다.")
    dart = [
        {"name": "SK리츠", "yield": 6.31, "dps": 77.5, "months": 3, "rate": 4.21, "rate_est": False, "ltv": 58.2, "nav_eok": 16890,
         "pnav": 0.88, "report": "사업보고서 (2026.06) (2026.09.12 공시)", "flags": []},
        {"name": "코람코라이프인프라리츠", "yield": 7.06, "dps": 125, "months": 6, "rate": 4.95, "rate_est": True, "ltv": 61.4,
         "nav_eok": 4120, "pnav": 0.84, "report": "사업보고서 (2026.06) (2026.09.10 공시)", "flags": ["변동금리는 CD 2.87%+가산금리로 추정"]},
        {"name": "NH올원리츠", "yield": None, "dps": None, "months": None, "rate": 5.12, "rate_est": False, "ltv": 55.0,
         "nav_eok": 1610, "pnav": 0.82, "report": "반기보고서 (2026.06) (2026.08.14 공시)", "flags": ["최근 배당결정 공시 없음"]}]
    return {"now": now, "reits": reits, "major": major, "new": new, "rates": rates, "fred": fred, "price_date": "2026-11-02",
            "since_major": now - dt.timedelta(days=30), "since_new": now - dt.timedelta(days=5), "dart": dart}


def dart_only():
    """DART 재무 지표만 계산해 화면에 표 형태로 출력(메일 발송 안 함)."""
    load_env()
    rates = fetch_rates(os.environ.get("ECOS_API_KEY"), dt.datetime.now(KST).date())
    reits = []
    for r in CONFIG["reits"]:
        code = r.get("code")
        if code:
            q = fetch_quote(code)
            reits.append({"name": r["name"], "code": code, "close": q["close"], "mcap_eok": q["mcap_eok"]})
    res = collect_dart(reits, rates) or []
    f = lambda v, d=2: "-" if v is None else f"{v:,.{d}f}"  # noqa: E731
    print(f"\n{'종목':<16}{'배당수익률':>9}{'차입금리':>9}{'LTV':>8}{'NAV(억)':>10}{'P/NAV':>7}  비고")
    for m in res:
        print(f"{m['name'][:14]:<16}{f(m.get('yield')):>9}{f(m.get('rate')):>9}{f(m.get('ltv'), 1):>8}"
              f"{f(m.get('nav_eok'), 0):>10}{f(m.get('pnav')):>7}  {' · '.join(m.get('flags') or [])}")
    print(f"\n계산 근거(종목별): {HERE / 'data' / 'dart_trace'}")
    if LOG:
        print("\n경고:", *LOG, sep="\n - ")


def run_once(preview: bool):
    now = dt.datetime.now(KST)
    state = load_state()
    R = collect(now, state)
    subj, body = render(R)
    if preview:
        p = HERE / "report_preview.html"
        p.write_text(f"<!doctype html><meta charset='utf-8'><title>{html.escape(subj)}</title>{body}", encoding="utf-8")
        print(f"미리보기 저장: {p}\n제목: {subj}")
        if os.name == "nt":
            os.startfile(str(p))  # noqa: S606
        return
    send(subj, body)
    save_state(state, now)
    print(f"발송 완료: {subj}")


def menu():
    env = HERE / "env.txt"
    if not env.exists() and (HERE / "env.example.txt").exists():
        env.write_text((HERE / "env.example.txt").read_text(encoding="utf-8"), encoding="utf-8")
    items = {"1": "점검 (접속·수집 확인, 메일 안 보냄)", "2": "미리보기 (오늘 메일을 브라우저로 열기)",
             "3": "지금 메일 보내기", "4": f"평일 {CONFIG.get('send_time', '09:10')} 자동 발송 등록(시간 변경도 이 메뉴)", "5": "설정 파일(env.txt) 열기",
             "6": "DART 재무 지표만 계산(배당수익률·차입금리·LTV·P/NAV, 근거 폴더 열기)", "0": "끝내기"}
    while True:
        print("\n ==========================================")
        print(f"  상장리츠 공시·금리 일일 보고   (Python {sys.version.split()[0]})")
        print(" ==========================================")
        for k, v in items.items():
            print(f"  {k}. {v}")
        sel = input("\n번호를 입력하고 Enter: ").strip()
        LOG.clear()
        try:
            if sel == "1":
                print("\n[점검] 1~3분 걸릴 수 있습니다...")
                check()
                print("\n경고가 있으면 이 창을 캡처해 보내 주세요.")
            elif sel == "2":
                print("\n[미리보기] 24개 종목을 수집합니다. 2~5분 걸릴 수 있습니다...")
                run_once(preview=True)
            elif sel == "3":
                print("\n[발송] 수집 후 메일을 보냅니다. 2~5분 걸릴 수 있습니다...")
                run_once(preview=False)
            elif sel == "4":
                bat = HERE / "run_daily.bat"
                cmd = (f'schtasks /create /tn ReitsDailyReport /tr "\\"{bat}\\" auto" '
                       f'/sc weekly /d MON,TUE,WED,THU,FRI /st {CONFIG.get("send_time", "09:10")} /f')
                rc = os.system(cmd)
                print(f"등록 완료: 평일 {CONFIG.get('send_time', '09:10')} 자동 발송 (해제: schtasks /delete /tn ReitsDailyReport /f)" if rc == 0
                      else "등록 실패 — 이 창을 캡처해 보내 주세요.")
            elif sel == "5":
                os.system(f'notepad "{env}"')
            elif sel == "6":
                print("\n[DART] 처음 실행은 보고서 원문을 받느라 3~10분 걸립니다(다음부터는 새 보고서만 받음)...")
                dart_only()
                tr = HERE / "data" / "dart_trace"
                if os.name == "nt" and tr.exists():
                    os.startfile(str(tr))  # noqa: S606
            elif sel == "0":
                return
        except SystemExit as e:
            print(f"\n[안내] {e}")
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"\n[오류] {e} — 이 창을 캡처해 보내 주세요.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--menu", action="store_true", help="번호 메뉴(run_daily.bat 더블클릭 시)")
    ap.add_argument("--check", action="store_true", help="접속·파싱 점검(발송 안 함)")
    ap.add_argument("--preview", action="store_true", help="수집 후 report_preview.html 저장(발송 안 함)")
    ap.add_argument("--demo", action="store_true", help="샘플 데이터로 report_preview.html 저장")
    ap.add_argument("--dart", action="store_true", help="DART 재무 지표만 계산해 표시")
    a = ap.parse_args()
    if a.dart:
        return dart_only()
    if a.menu:
        return menu()
    if a.check:
        return check()
    now = dt.datetime.now(KST)
    state = load_state()
    R = demo() if a.demo else collect(now, state)
    subj, body = render(R)
    if a.preview or a.demo:
        p = HERE / "report_preview.html"
        p.write_text(f"<!doctype html><meta charset='utf-8'><title>{html.escape(subj)}</title>{body}", encoding="utf-8")
        print(f"미리보기 저장: {p}\n제목: {subj}")
        return
    send(subj, body)
    save_state(state, now)
    print(f"발송 완료: {subj}")


if __name__ == "__main__":
    main()
