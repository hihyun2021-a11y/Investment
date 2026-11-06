# -*- coding: utf-8 -*-
"""OpenDART로 상장리츠 재무 지표 4종을 계산한다 (reits_daily.py에서 호출, 표준 라이브러리만 사용).

1. 시가배당수익률(연환산) = 최근 '부동산투자회사금전배당결정' 1주당 배당금 × (12 ÷ 결산기간 개월) ÷ 전일 종가
2. 가중평균 차입금리     = 최근 사업·반기보고서 「재무제표 주석」 차입금·사채 표의 당기말 장부금액 가중평균 이자율
3. LTV                 = 차입금·사채 당기말 장부금액 합계(재무상태표 과목 합계로 확정, 주석 합계는 대조용) ÷ 자산가치
                         자산가치 = 종속·관계기업투자 기말 장부금액 + 투자부동산 공정가치(서열체계 주석, 없으면 기말 장부금액)
4. P/NAV               = 시가총액 ÷ (자산가치 − 차입금·사채 − 재무상태표 임대보증금)

원문은 data/dart/에 캐시한다(접수번호 단위, 한 번 받으면 다시 받지 않음). 계산 근거(어느 표·행을 썼는지)는
data/dart_trace/ 에 종목별 텍스트로 남긴다. 자동 추출이 틀리면 config.json "dart_override"로 값을 덮어쓴다.
"""
from __future__ import annotations

import datetime as dt
import io
import json
import re
import time
import urllib.parse
import urllib.request
import zipfile
from html.parser import HTMLParser
from pathlib import Path

API = "https://opendart.fss.or.kr/api"
PARSER_VERSION = 12
HERE = Path(__file__).resolve().parent
DART_DIR = HERE / "data" / "dart"
TRACE_DIR = HERE / "data" / "dart_trace"
UNIT = {"원": 1, "천원": 1e3, "백만원": 1e6, "억원": 1e8}


# ────────────────────────────────────────────────────────────── API
def _get(url: str, params: dict, timeout=30) -> bytes:
    full = url + "?" + urllib.parse.urlencode(params)
    last = None
    for i in range(3):
        try:
            req = urllib.request.Request(full, headers={"User-Agent": "reits-daily-report"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(1 + i)
    raise RuntimeError(f"DART 접속 실패 {url.rsplit('/', 1)[-1]}: {last}")


def _api_error(raw: bytes) -> str | None:
    """오류면 메시지(예: 010 등록되지 않은 키), 정상이면 None."""
    head = raw[:400].decode("utf-8", "ignore")
    m = re.search(r"<status>(\d+)</status>\s*<message>([^<]*)", head) or \
        re.search(r'"status"\s*:\s*"(\d+)"\s*,\s*"message"\s*:\s*"([^"]*)"', head)
    if m and m.group(1) not in ("000", "013"):
        return f"{m.group(1)} {m.group(2)}"
    return None


def corp_codes(key: str, log) -> dict:
    """종목코드 → {corp_code, corp_name}. 30일마다 갱신."""
    cache = DART_DIR / "corpcode.json"
    if cache.exists() and time.time() - cache.stat().st_mtime < 30 * 86400:
        return json.loads(cache.read_text(encoding="utf-8"))
    raw = _get(f"{API}/corpCode.xml", {"crtfc_key": key}, timeout=60)
    err = _api_error(raw)
    if err:
        raise RuntimeError(f"DART 인증키 확인 필요: {err}")
    xml = zipfile.ZipFile(io.BytesIO(raw)).read("CORPCODE.xml").decode("utf-8", "replace")
    out = {}
    for m in re.finditer(r"<list>(.*?)</list>", xml, re.S):
        blk = m.group(1)
        g = lambda t: (re.search(rf"<{t}>(.*?)</{t}>", blk, re.S) or [None, ""])[1].strip()  # noqa: E731
        sc = g("stock_code")
        if sc:
            out[sc] = {"corp_code": g("corp_code"), "corp_name": g("corp_name")}
    DART_DIR.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    log(f"[DART] 회사 고유번호 {len(out)}건 갱신")
    return out


def filings(key: str, corp_code: str, days: int, ty: str | None = None) -> list[dict]:
    end = dt.date.today()
    p = {"crtfc_key": key, "corp_code": corp_code, "bgn_de": (end - dt.timedelta(days=days)).strftime("%Y%m%d"),
         "end_de": end.strftime("%Y%m%d"), "last_reprt_at": "Y", "page_count": 100, "page_no": 1}
    if ty:
        p["pblntf_ty"] = ty
    out = []
    for page in range(1, 6):
        p["page_no"] = page
        j = json.loads(_get(f"{API}/list.json", p).decode("utf-8", "replace"))
        if j.get("status") not in ("000", "013"):
            raise RuntimeError(f"DART 목록 오류 {j.get('status')} {j.get('message')}")
        out += j.get("list") or []
        if page >= int(j.get("total_page") or 1):
            break
    return sorted(out, key=lambda x: x.get("rcept_no", ""), reverse=True)


def document(key: str, rcept_no: str) -> list[str]:
    """원문 XML 목록(본문 먼저). data/dart/<접수번호>.zip 캐시."""
    DART_DIR.mkdir(parents=True, exist_ok=True)
    cf = DART_DIR / f"{rcept_no}.zip"
    if cf.exists():
        raw = cf.read_bytes()
    else:
        raw = _get(f"{API}/document.xml", {"crtfc_key": key, "rcept_no": rcept_no}, timeout=90)
        err = _api_error(raw)
        if err:
            raise RuntimeError(f"원문 {rcept_no}: {err}")
        cf.write_bytes(raw)
        time.sleep(0.3)
    zf = zipfile.ZipFile(io.BytesIO(raw))
    names = sorted(zf.namelist(), key=lambda n: (not n.startswith(rcept_no + "."), len(n), n))
    texts = []
    for n in names:
        b = zf.read(n)
        for enc in ("utf-8", "euc-kr", "cp949"):
            try:
                texts.append(b.decode(enc))
                break
            except UnicodeDecodeError:
                continue
        else:
            texts.append(b.decode("utf-8", "replace"))
    return texts


# ────────────────────────────────────────────────────────────── DART XML → 블록
CELL = {"td", "th", "te", "tu"}


class _Blocks(HTMLParser):
    """TITLE / P / TABLE 순서대로 블록 목록을 만든다. 표는 COLSPAN·ROWSPAN을 펼친 격자."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.blocks, self.buf, self.kind = [], [], "p"
        self.tdepth, self.rows, self.row, self.cell, self.span = 0, [], None, None, (1, 1)
        self.head_rows, self.group = 0, 0

    def _flush(self):
        t = re.sub(r"\s+", " ", "".join(self.buf)).strip()
        if t:
            self.blocks.append({"kind": self.kind, "text": t, "in_group": self.group > 0})
        self.buf, self.kind = [], "p"

    def handle_starttag(self, tag, attrs):
        a = {k.lower(): (v or "") for k, v in attrs}
        if tag == "table-group" and not self.tdepth:
            self._flush()
            self.group += 1
            return
        if tag == "table":
            if self.tdepth == 0:
                self._flush()
                self.rows, self.head_rows = [], 0
            self.tdepth += 1
        elif self.tdepth:
            if tag == "tr":
                self.row = []
            elif tag in CELL:
                self.cell = []
                try:
                    self.span = (max(1, int(a.get("colspan") or 1)), max(1, int(a.get("rowspan") or 1)))
                except ValueError:
                    self.span = (1, 1)
                self._th = tag == "th"
            elif tag in ("br", "p") and self.cell is not None:
                self.cell.append(" ")
        else:
            if tag == "title":
                self._flush()
                self.kind = "title"
            elif tag in ("p", "br", "div", "section-1", "section-2", "section-3", "library"):
                self._flush()

    def handle_endtag(self, tag):
        if tag == "table-group" and not self.tdepth:
            self._flush()
            self.group = max(0, self.group - 1)
            return
        if self.tdepth:
            if tag in CELL and self.cell is not None and self.row is not None:
                txt = re.sub(r"\s+", " ", "".join(self.cell)).strip()
                self.row.append((txt, *self.span, getattr(self, "_th", False)))
                self.cell = None
            elif tag == "tr" and self.row is not None:
                self.rows.append(self.row)
                self.row = None
            elif tag == "table":
                self.tdepth -= 1
                if self.tdepth == 0:
                    grid = _expand(self.rows)
                    if grid:
                        self.blocks.append({"kind": "table", "rows": grid, "th": [all(c[3] for c in r) for r in self.rows]})
        elif tag in ("title", "p"):
            self._flush()

    def handle_data(self, d):
        if self.tdepth:
            if self.cell is not None:
                self.cell.append(d)
        else:
            self.buf.append(d)


def _expand(rows):
    grid, pending = [], {}
    for r in rows:
        out, c = [], 0
        cells = list(r)
        while cells or pending.get(c):
            if pending.get(c):
                txt, left = pending[c]
                out.append(txt)
                pending[c] = (txt, left - 1) if left > 1 else None
                c += 1
                continue
            txt, cs, rs, _th = cells.pop(0)
            for k in range(cs):
                out.append(txt)
                if rs > 1:
                    pending[c + k] = (txt, rs - 1)
            c += cs
        grid.append(out)
    w = max((len(r) for r in grid), default=0)
    return [r + [""] * (w - len(r)) for r in grid]


def blocks_of(xml: str) -> list[dict]:
    p = _Blocks()
    p.feed(xml)
    p._flush()
    return p.blocks


# ────────────────────────────────────────────────────────────── 표 해석
AMT = re.compile(r"^\(?\s*[-△▲]?\s*[\d,]+(?:\.\d+)?\s*\)?$")
CUR = re.compile(r"\(당\)|당\s*\)?\s*(반|분)?\s*기")
PRI = re.compile(r"\(전\)|전\s*\)?\s*(반|분)?\s*기")
DATE = re.compile(r"(20\d{2})\s*[.\-년]\s*(\d{1,2})\s*[.\-월]\s*(\d{1,2})")
SKIP_ROW = re.compile(r"합\s*계|소\s*계|총\s*계|^\s*계\s*$|유동성|차감|할인|현재가치|상각|대체|가산|조정|이연")
TOTAL_ROW = re.compile(r"합\s*계|총\s*계|^\s*계\s*$")


def amount(s: str) -> float | None:
    s = (s or "").strip()
    if s in ("-", "–", "—"):
        return 0.0
    if not AMT.match(s):
        return None
    neg = s.startswith("(") or "△" in s or "-" in s or "▲" in s
    v = float(re.sub(r"[^\d.]", "", s) or 0)
    return -v if neg else v


def unit_of(text: str) -> float | None:
    m = re.search(r"단\s*위\s*[:：]?\s*(백만원|천원|억원|원)", text or "")
    return UNIT[m.group(1)] if m else None


def rate_of(s: str) -> float | None:
    """'4.50%' → 4.5, '4.2~5.1' → 4.65, 'CD+1.5' → None(가산금리만 있음), 'CD+1.5(4.67)' → 4.67."""
    s = (s or "").replace(" ", "")
    if not s:
        return None
    if re.search(r"CD|KOFR|COFIX|SOFR|LIBOR|금융채|변동|기준금리|\+", s, re.I):   # 변동금리 표기
        par = [float(x) for x in re.findall(r"\((\d+\.\d+)%?\)", s)]
        return par[0] if par and 0.5 < par[0] < 30 else None
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", s)]
    nums = [x for x in nums if 0 < x < 30]
    if not nums:
        return None
    if len(nums) >= 2 and re.search(r"[~∼～\-]", s):
        return (nums[0] + nums[1]) / 2
    return nums[0]


def spread_of(s: str):
    """변동금리 'CD(91일)+1.80' → ('CD', 1.8). 해석 못 하면 None."""
    s = (s or "").replace(" ", "")
    m = re.search(r"(CD|KOFR|COFIX|금융채)[^+]*\+\s*(\d+(?:\.\d+)?)", s, re.I)
    return (m.group(1).upper(), float(m.group(2))) if m else None


NUMLIKE = re.compile(r"^[\d.,\s%~∼～\-()]+$")


class T:
    """표 하나: 머리행 분리, 열 머리말 합성, 당기 열 판정."""

    def __init__(self, blk, context: str):
        self.rows = blk["rows"]
        self.context = context
        th = blk.get("th") or []
        if th and th[0]:                                   # THEAD(TH) 행이 머리행
            n = next((i for i, x in enumerate(th) if not x), len(th))
        else:
            n = 0
            for i, r in enumerate(self.rows):
                nums = [amount(c) for c in r[1:]]
                if any(v is not None and (abs(v) >= 100 or "," in r[j + 1]) for j, v in enumerate(nums)):
                    break
                n = i + 1
        self.nh = min(n, len(self.rows))
        w = len(self.rows[0]) if self.rows else 0
        self.head = []
        for j in range(w):
            parts = []
            for i in range(self.nh):
                t = self.rows[i][j]
                if t and t not in parts and not unit_of(t):
                    parts.append(t)
            self.head.append(" ".join(parts))
        self.body = self.rows[self.nh:]
        self.unit = unit_of(context) or next((unit_of(c) for r in self.rows[: self.nh + 1] for c in r if unit_of(c)), None)
        self.text = " ".join(" ".join(r) for r in self.rows)

    def cols(self, pat):
        return [j for j, h in enumerate(self.head) if re.search(pat, h)]

    def period_cols(self):
        """당기 열 목록. 머리말에 당기·전기가 없으면 날짜가 가장 늦은 열, 그것도 없으면 None(=표 전체 하나의 기간)."""
        cur = [j for j, h in enumerate(self.head) if CUR.search(h) and not PRI.search(h)]
        if cur:
            return cur
        dated = {j: DATE.search(h) for j, h in enumerate(self.head)}
        dated = {j: (int(m[1]), int(m[2]), int(m[3])) for j, m in dated.items() if m}
        if len(set(dated.values())) >= 2:
            last = max(dated.values())
            return [j for j, d in dated.items() if d == last]
        if any(PRI.search(h) for h in self.head):
            return [j for j, h in enumerate(self.head) if j > 0 and not PRI.search(h)]
        return None

    def is_prior_only(self):
        ctx = self.context[-120:]
        return bool(PRI.search(ctx)) and not CUR.search(ctx)

    def body_current(self):
        """행 묶음(첫 열에 '당기말/전기말')이 있으면 당기 묶음만."""
        out, grp = [], None
        for r in self.body:
            c0 = r[0]
            if re.match(r"^\s*(\(|<|\[)?\s*(당|전)\s*\)?\s*(반|분)?\s*기", c0 or ""):
                grp = "pri" if PRI.match(re.sub(r"^[\s(<\[]+", "", c0)) else "cur"
            if grp != "pri":
                out.append(r)
        return out

    def label(self, r, skip_cols=()):
        parts = [c for j, c in enumerate(r) if j not in skip_cols and c and amount(c) is None
                 and not NUMLIKE.match(c) and not DATE.search(c)]
        seen = []
        for p in parts:
            if p not in seen:
                seen.append(p)
        return " ".join(seen)

    @staticmethod
    def pick(r, cols, prefer=None):
        """여러 당기 열 중 값: '합계' 열 우선, 아니면 첫 번째 숫자."""
        order = sorted(cols, key=lambda j: 0 if (prefer and prefer(j)) else 1)
        for j in order:
            v = amount(r[j]) if j < len(r) else None
            if v is not None and r[j].strip() not in ("",):
                return v, j
        return None, None


def _context(blocks, i):
    """표 앞 설명: 문단 + 바로 앞 작은 표(XBRL 주석의 '당반기말 | (단위 : 천원)' 같은 캡션)."""
    prev = []
    for b in blocks[max(0, i - 4):i]:
        if b["kind"] != "table":
            prev.append(b["text"])
        elif len(b["rows"]) <= 3:
            prev.append(" | ".join(dict.fromkeys(c for r in b["rows"] for c in r if c)))
        else:
            prev = []                      # 큰 표를 지나면 앞 설명은 그 표의 것
    return " ".join(prev)


# ────────────────────────────────────────────────────────────── 주석 구역
HEAD = re.compile(r"^\s*(\d{1,2})\s*[.．]\s*(\S.*)$")
KEY_DEBT = re.compile(r"차입금|사채")
KEY_IP = re.compile(r"투자부동산")
KEY_EQ = re.compile(r"종속기업|관계기업|공동기업|지분법")


def _is_section_title(b):
    return b["kind"] == "title" and not b.get("in_group")


def notes_section(blocks):
    """별도 「재무제표 주석」 구역(다음 대단원 제목 전까지). 없으면 연결 주석, 그것도 없으면 None."""
    idx = [i for i, b in enumerate(blocks) if _is_section_title(b)]
    for want_sep in (True, False):
        for k, i in enumerate(idx):
            t = blocks[i]["text"]
            if re.search(r"재무제표\s*주석", t) and (("연결" not in t) == want_sep):
                end = idx[k + 1] if k + 1 < len(idx) else len(blocks)
                sec = blocks[i + 1:end]
                if sum(1 for b in sec if b["kind"] == "table") >= 3:
                    return sec, ("별도" if want_sep else "연결")
    return None, None


def statements_section(blocks, basis):
    idx = [i for i, b in enumerate(blocks) if _is_section_title(b)]
    for k, i in enumerate(idx):
        t = re.sub(r"\s+", "", blocks[i]["text"])
        ok = re.search(r"\d\.(연결)?재무제표$", t) and (("연결" in t) == (basis == "연결"))
        if ok:
            end = idx[k + 1] if k + 1 < len(idx) else len(blocks)
            return blocks[i + 1:end]
    return None


def split_notes(sec):
    """[{no, title, blocks}] — XBRL 주석은 주석마다 제목(TITLE)이 있고, 옛 양식은 '12. 차입금' 문단으로 나뉜다.
    문단 제목은 주석 번호가 커지는 순서로만 인정(본문 안의 '1. …' 목록 오인 방지)."""
    if any(b["kind"] == "title" for b in sec):
        notes, cur = [], None
        for b in sec:
            if b["kind"] == "title":
                m = re.match(r"^\s*(\d{1,2})", b["text"])
                cur = {"no": int(m.group(1)) if m else 0,
                       "title": re.sub(r"^\s*[\d\-]+\s*[.．]\s*", "", b["text"])[:45], "blocks": []}
                notes.append(cur)
            elif cur is not None:
                cur["blocks"].append(b)
        return notes
    notes, cur, last = [], None, 0
    for b in sec:
        m = HEAD.match(b["text"]) if b["kind"] == "p" else None
        if m:
            n, title = int(m.group(1)), m.group(2)
            short = len(title) <= 45 and not re.search(r"(습니다|입니다|한다|있다)\.?$", title)
            keyed = bool(KEY_DEBT.search(title[:25]) or KEY_IP.search(title[:25]) or KEY_EQ.search(title[:25]))
            if (short and last < n <= last + 3) or (keyed and last < n <= last + 15):
                cur = {"no": n, "title": title[:45], "blocks": []}
                notes.append(cur)
                last = n
                continue
        if cur is not None:
            cur["blocks"].append(b)
    return notes


# ────────────────────────────────────────────────────────────── 지표 추출
CONTRA = re.compile(r"할인발행차금|현재가치할인|사채할인|할증")
TOTAL_LVL = re.compile(r"^(?!.*이자율\s*유형\s*합계).*합\s*계")          # '이자율 유형 합계'(차입처 안 소계)만 예외
CONTRA_LVL = re.compile(r"현재가치할인|상환할증|전환권조정|사채할인|할인발행|할증")


def _pct(v):
    return None if v is None else (v * 100 if 0 < v < 1 else v)


def is_pivot(t):
    """XBRL 표준 주석(열 = 차입처·자산분류, 행 = 항목, 머리행의 첫 칸이 비어 있음)인지."""
    return t.nh >= 1 and all(not (t.rows[i][0] or "").strip() for i in range(t.nh))


def transpose(t):
    """행 = 차입처, 열 = 항목(이자율·금액)인 XBRL 표를 열 = 차입처 형태로 뒤집는다."""
    G, nh = t.rows, max(t.nh, 1)
    L = 0
    while L < len(G[0]) and not any((G[i][L] or "").strip() for i in range(nh)):
        L += 1
    if L == 0:
        return None
    data = G[nh:]
    rows = [[""] + [r[k] for r in data] for k in range(L)]
    for j in range(L, len(G[0])):
        rows.append([" ".join(dict.fromkeys(G[i][j] for i in range(nh) if G[i][j]))] + [r[j] for r in data])
    return T({"rows": rows, "th": [True] * L + [False] * (len(rows) - L)}, t.context)


def debt_pivot(t, unit, tr):
    """XBRL 차입금·사채 표 → ([{label, won, rate}], 장부금액 합계)."""
    w = len(t.head)
    paths = [[t.rows[i][j] for i in range(t.nh)] for j in range(w)]
    body = t.body
    rate_row = next((r for r in body if re.search(r"이자율|금리", r[0]) and not re.search(r"유형|기준|조정|가산|스프레드|범위", r[0])), None)
    base_row = next((r for r in body if re.search(r"기준\s*(이자율|금리)\s*$", r[0])), None)
    sprd_row = next((r for r in body if re.search(r"(기준\s*(이자율|금리)\s*조정|가산\s*(금리|이자율)|스프레드)", r[0])), None)
    amt_rows = []
    for r in body:
        lab = r[0] or ""
        if re.match(r"^\s*(차감\s*[:：]?\s*)?\(?\s*유동성\s*대체\s*\)?\s*$", lab):
            break                                   # 이 아래는 유동성 대체 후 금액(중복)
        if re.search(r"발행일|만기|이자율|금리|명칭|기술|정보|설명|일자|조건|담보|유동성\s*대체\s*부분|비유동성\s*부분|유동\s*부분", lab):
            continue
        if re.match(r"^\s*(유동성|비유동성?)\s", lab):      # '유동성 장기차입금'·'비유동성 장기차입금' = 위 합계를 나눈 행
            continue
        if any(amount(c) not in (None, 0.0) for c in r[1:]):
            amt_rows.append(r)
    all_rows = [r for r in body if not re.search(r"발행일|만기|이자율|금리|명칭|기술|정보|설명|일자|조건|담보|^\s*(차감\s*[:：]?\s*)?유동성\s*대체\s*$", r[0] or "")
                and any(amount(c) not in (None, 0.0) for c in r[1:])]
    tot_cols = [j for j in range(1, w) if paths[j] and re.search(r"합\s*계", paths[j][0])]
    book = None
    if tot_cols:
        jt = tot_cols[-1]
        vals = [amount(r[jt]) for r in amt_rows]
        if any(v is not None for v in vals):
            book = sum(v for v in vals if v is not None) * unit
    detail = [j for j in range(1, w) if not any(TOTAL_LVL.search(x) for x in paths[j])
              and not any(CONTRA_LVL.search(x) for x in paths[j])]
    glist = []                                      # 붙어 있는 열만 같은 대출(같은 은행의 다른 대출은 따로)
    for j in detail:
        key = tuple(x for x in paths[j] if x and not re.search(r"이자율\s*유형|고정|변동", x))
        member = any(re.search(r"고정|변동|이자율\s*유형\s*$", x) for x in paths[j])
        if glist and glist[-1][0] == key and glist[-1][2] == j - 1 and not (member and glist[-1][3]):
            glist[-1][1].append(j)
            glist[-1][2] = j
            glist[-1][3] = glist[-1][3] or member
        else:
            glist.append([key, [j], j, member])
    rows = []
    for key, cols, _last, _m in glist:
        amt = 0.0                                   # 같은 대출이 합계·유동·비유동 행에 반복되므로 최댓값
        for r in all_rows:
            v = next((amount(r[j]) for j in cols if amount(r[j]) is not None and r[j].strip()), None)
            amt = max(amt, v or 0)
        if not amt:
            continue
        rate = base = spread = None
        if rate_row:
            rate = next((_pct(amount(rate_row[j])) for j in cols if amount(rate_row[j])), None)
        if rate is None and base_row:
            base = next((base_row[j].strip() for j in cols if base_row[j].strip()), None)
            if sprd_row:
                spread = next((_pct(amount(sprd_row[j])) for j in cols if amount(sprd_row[j]) is not None and sprd_row[j].strip()), None)
            if base and re.search(r"CD", base, re.I):
                base = "CD"
        lab = key[-1] if key else "?"
        rows.append({"label": lab, "won": amt * unit, "rate": rate, "base": base, "spread": spread, "raw_rate": None})
        rr = f"{rate:.2f}" if rate is not None else (f"{base}+{spread:.2f}" if base and spread is not None else "[없음]")
        tr.append(f"    {lab[:40]:<40} {amt * unit / 1e8:>10,.1f}억  이자율 {rr}")
    if book is not None:
        tr.append(f"    → 장부금액 합계(현재가치할인차금 등 반영) {book / 1e8:,.1f}억")
    return rows, book


def pivot_value(t, row_pat, unit):
    """XBRL 표에서 row_pat 행의 값: 머리말에 '합계'가 있는 가장 오른쪽 열, 없으면 가장 큰 값."""
    w = len(t.head)
    for r in t.body:
        if not re.search(row_pat, r[0] or ""):
            continue
        tot = [j for j in range(1, w) if re.search(r"합\s*계", t.head[j]) and amount(r[j]) not in (None, 0.0)]
        if tot:
            return amount(r[tot[-1]]) * unit
        vals = [amount(c) for c in r[1:] if amount(c) not in (None, 0.0)]
        if vals:
            return max(vals) * unit
    return None


def debt_from_note(note, unit_default, tr):
    """[{label, won, rate, base, spread}] + 차감 계정(사채할인발행차금 등, 음수) 합."""
    rows, seen, contra = [], set(), 0.0
    bl = note["blocks"]
    for i, b in enumerate(bl):
        if b["kind"] != "table":
            continue
        t = T(b, _context(bl, i))
        if t.is_prior_only():
            tr.append(f"    표 건너뜀(전기): {t.context[-40:]}")
            continue
        tp = t if any(re.search(r"이자율", r[0]) for r in t.body) else None
        if tp is None and is_pivot(t) and any(re.search(r"이자율", h) for h in t.head):
            tp = transpose(t)
        if tp is not None and is_pivot(tp):
            prow, book = debt_pivot(tp, t.unit or unit_default, tr)
            rows += prow
            if book is not None:
                contra += book - sum(x["won"] for x in prow)
            continue
        rc = t.cols(r"이자율|금리|이율")
        if not rc:
            continue
        pc = t.period_cols()
        amt_cols = [j for j in (pc if pc is not None else range(1, len(t.head)))
                    if j not in rc and not re.search(r"이자율|금리|만기|발행|일자|기간|지분|상환|조건|담보|비고", t.head[j])]
        if not amt_cols:
            continue
        unit = t.unit or unit_default
        for r in t.body_current():
            lab = t.label(r, skip_cols=set(amt_cols) | set(rc))
            if not lab:
                continue
            v, j = T.pick(r, amt_cols, prefer=lambda j: "합계" in t.head[j])
            if not v:
                continue
            if CONTRA.search(lab):
                contra += -abs(v) * unit if "할증" not in lab else abs(v) * unit
                tr.append(f"    (차감) {lab[:36]:<36} {-abs(v) * unit / 1e8:>10,.1f}억")
                continue
            if SKIP_ROW.search(lab):
                continue
            rate, sp = _pct(rate_of(r[rc[0]])), spread_of(r[rc[0]])
            key = (re.sub(r"\s", "", lab), round(v))
            if key in seen:
                continue
            seen.add(key)
            rows.append({"label": lab, "won": v * unit, "rate": rate, "base": sp[0] if sp else None,
                         "spread": sp[1] if sp else None, "raw_rate": r[rc[0]]})
            rr = rate if rate is not None else (f"{sp[0]}+{sp[1]}" if sp else "[파싱불가]")
            tr.append(f"    {lab[:40]:<40} {v * unit / 1e8:>10,.1f}억  이자율 {rr}  ({t.head[j][:20]})")
    return rows, contra


def debt_summary(m: dict, cd: float | None) -> dict:
    """차입금·사채 합계(장부금액)와 가중평균 이자율. 변동금리는 CD(91일) 현재값 + 가산금리로 [추정]."""
    rows = m.get("debt_rows") or []
    if not rows:
        return {"debt_won": m.get("debt_won"), "rate": None, "rate_cover": None, "rate_est": False}
    tot = sum(r["won"] for r in rows) + (m.get("contra_won") or 0)
    num_, den, est = 0.0, 0.0, False
    for r in rows:
        rt = r["rate"]
        if rt is None and r.get("spread") is not None and r.get("base") == "CD" and cd is not None:
            rt, est = cd + r["spread"], True
        if rt is not None:
            num_ += r["won"] * rt
            den += r["won"]
    gross = max(sum(r["won"] for r in rows), m.get("bs_debt_won") or 0)   # 재무상태표 합계 대비 반영 비율
    return {"debt_won": tot, "rate": num_ / den if den else None, "rate_cover": den / gross if gross else None, "rate_est": est}


def value_from_table(t, unit_default, col_pat, tr, tag):
    """표에서 당기 값: 합계 행이 있으면 합계, 없으면 행 합(합계류 제외)."""
    pc = t.period_cols()
    base = pc if pc is not None else list(range(1, len(t.head)))
    cols = [j for j in base if re.search(col_pat, t.head[j])] or ([] if col_pat else base)
    if not cols:
        return None
    unit = t.unit or unit_default
    body = t.body_current()
    tot = [r for r in body if TOTAL_ROW.search(t.label(r, set(cols)))]
    if tot:
        v, j = T.pick(tot[-1], cols, prefer=lambda j: "합계" in t.head[j])
        if v:
            tr.append(f"    {tag}: 합계 행 {v * unit / 1e8:,.1f}억 ({t.head[j][:24]})")
            return v * unit
    s, n = 0.0, 0
    for r in body:
        lab = t.label(r, set(cols))
        if not lab or SKIP_ROW.search(lab):
            continue
        v, _ = T.pick(r, cols, prefer=lambda j: "합계" in t.head[j])
        if v:
            s += v
            n += 1
    if n:
        tr.append(f"    {tag}: {n}개 행 합 {s * unit / 1e8:,.1f}억")
        return s * unit
    return None


def ip_from_note(note, unit_default, tr):
    """투자부동산: 공정가치 서열체계 표의 측정금액 → 없으면 장부금액."""
    bl = note["blocks"]
    tables = [(i, T(b, _context(bl, i))) for i, b in enumerate(bl) if b["kind"] == "table"]
    for pat, basis in ((r"서열\s*체계", "공정가치(서열체계)"), (r"공정가치", "공정가치")):
        for i, t in tables:
            if re.search(pat, t.context) and is_pivot(t) and not t.is_prior_only():
                v = pivot_value(t, r"투자부동산", t.unit or unit_default)
                if v:
                    tr.append(f"    {basis}: {v / 1e8:,.1f}억 ({t.context[-50:]})")
                    return v, "공정가치"
    fv = []
    for i, t in tables:
        ctx = t.context + " " + " ".join(t.head)
        if re.search(r"서열|수준\s*[123]|Level", ctx) and not t.is_prior_only():
            fv.append(t)
    for t in fv:
        v = value_from_table(t, unit_default, r"합\s*계|공정가치|수준\s*3", tr, "공정가치(서열체계)")
        if v:
            return v, "공정가치"
    for i, t in tables:
        if t.is_prior_only() or not t.cols(r"공정가치"):
            continue
        v = value_from_table(t, unit_default, r"공정가치", tr, "공정가치")
        if v:
            return v, "공정가치"
    for i, t in tables:
        if t.is_prior_only():
            continue
        if is_pivot(t):
            v = pivot_value(t, r"^\s*투자부동산\s*$", t.unit or unit_default)
            if v:
                tr.append(f"    장부금액(XBRL): {v / 1e8:,.1f}억")
                return v, "장부금액"
            continue
        v = value_from_table(t, unit_default, r"장부|기말", tr, "장부금액")
        if v:
            return v, "장부금액"
    return None, None


def eq_from_note(note, unit_default, tr):
    bl = note["blocks"]
    for i, b in enumerate(bl):
        if b["kind"] != "table":
            continue
        t = T(b, _context(bl, i))
        if t.is_prior_only():
            continue
        for pat in (r"장부", r"^\s*기\s*말\s*$|[^당전반분]기말"):
            v = value_from_table(t, unit_default, pat, tr, "지분투자 장부금액")
            if v:
                return v
    return None


def deposit_from_notes(notes, unit_default, tr):
    """주석 표의 '임대보증금' 행(당기, 장부금액) 합. 처음 찾은 표만 쓴다."""
    for n in notes:
        bl = n["blocks"]
        for i, b in enumerate(bl):
            if b["kind"] != "table":
                continue
            t = T(b, _context(bl, i))
            if t.is_prior_only():
                continue
            hit = [r for r in t.body if re.search(r"임대\s*보증금", r[0] or "") and not re.search(r"설명|기술|할인|현재가치|담보|관련", r[0])]
            if not hit:
                continue
            unit = t.unit or unit_default
            tot = 0.0
            for r in hit:
                if is_pivot(t):
                    w = len(t.head)
                    cols = [j for j in range(1, w) if re.search(r"합\s*계", t.head[j]) and amount(r[j]) is not None]
                    v = amount(r[cols[-1]]) if cols else next((amount(c) for c in r[1:] if amount(c)), None)
                else:
                    pc = t.period_cols() or list(range(1, len(t.head)))
                    v, _ = T.pick(r, pc)
                tot += (v or 0) * unit
            if tot:
                tr.append(f"  임대보증금(주석 {n['no']}. {n['title'][:12]}): {tot / 1e8:,.1f}억 — " + ", ".join(r[0][:12] for r in hit))
                return tot
    return None


def statement_items(blocks, tr):
    """재무상태표 당기 값 {과목: 원}."""
    for b in blocks or []:
        if b["kind"] != "table":
            continue
        t = T(b, "")
        if not re.search(r"자\s*산\s*총\s*계", t.text):
            continue
        pc = t.period_cols()
        cols = pc or [j for j in range(1, len(t.head)) if not re.search(r"주\s*석", t.head[j])]
        unit = t.unit or unit_of(" ".join(x.get("text", "") for x in blocks[:5])) or 1
        items = {}
        for r in t.body:
            lab = re.sub(r"^[\sⅠ-Ⅻ\dIVX.()]+", "", t.label(r, set(cols))).replace(" ", "")
            v, _ = T.pick(r, cols)
            if lab and v is not None:
                items[lab] = items.get(lab, 0.0) + v * unit      # 유동·비유동에 같은 과목이 있으면 합산
        tr.append(f"  재무상태표: {len(items)}개 과목 (단위 {int(unit):,}원)")
        return items
    return {}


def report_metrics(xmls: list[str], tr: list[str]) -> dict:
    res = {"debt_won": None, "debt_rows": [], "contra_won": 0.0, "ip_won": None, "ip_basis": None, "eq_won": None,
           "deposit_won": None, "basis": None, "flags": []}
    for x_i, xml in enumerate(xmls):
        blocks = blocks_of(xml)
        sec, basis = notes_section(blocks)
        if not sec:
            continue
        res["basis"] = basis
        tr.append(f"  주석 구역: {basis} 재무제표 주석 (원문 파일 {x_i + 1}/{len(xmls)}, 블록 {len(sec)}개)")
        unit_default = next((unit_of(b.get("text", "")) for b in sec if unit_of(b.get("text", ""))), None) or 1
        notes = split_notes(sec)
        tr.append("  주석 목록: " + ", ".join(f"{n['no']}.{n['title'][:12]}" for n in notes))
        debt_rows, contra = [], 0.0
        for n in notes:
            title = n["title"][:30]
            if KEY_DEBT.search(title) and not re.search(r"금융상품|위험|리스|차입원가", title):
                tr.append(f"  [차입금·사채] {n['no']}. {title}")
                rows, c = debt_from_note(n, unit_default, tr)
                debt_rows += rows
                contra += c
            elif KEY_IP.search(title) and res["ip_won"] is None:
                tr.append(f"  [투자부동산] {n['no']}. {title}")
                res["ip_won"], res["ip_basis"] = ip_from_note(n, unit_default, tr)
            elif KEY_EQ.search(title) and res["eq_won"] is None:
                tr.append(f"  [지분투자] {n['no']}. {title}")
                res["eq_won"] = eq_from_note(n, unit_default, tr)
        if debt_rows:
            res["debt_rows"], res["contra_won"] = debt_rows, contra
            res["debt_won"] = sum(r["won"] for r in debt_rows) + contra
        st = statement_items(statements_section(blocks, basis), tr)
        bs_debt = [(k, v) for k, v in st.items() if re.search(r"차입금|사채|유동성장기부채", k)
                   and not re.search(r"할인|이자|원가|총계|합계|부채총계|^(유동|비유동)부채", k)]
        if bs_debt:
            bsv = sum(v for _, v in bs_debt)
            res["bs_debt_won"] = bsv
            note_v = res["debt_won"]
            tr.append(f"  차입금·사채 합계: 재무상태표 {bsv / 1e8:,.1f}억 ({', '.join(k for k, _ in bs_debt)})"
                      + (f" / 주석 대조 {note_v / 1e8:,.1f}억" if note_v is not None else "")
                      + (" ※ 차이" if note_v is not None and bsv and abs(note_v - bsv) / bsv > 0.02 else ""))
            res["contra_won"] = bsv - sum(r["won"] for r in res.get("debt_rows") or [])
            res["debt_won"] = bsv
        dep = [(k, v) for k, v in st.items() if "임대보증금" in k and not re.search(r"할인|현재가치", k)]
        if dep:
            res["deposit_won"] = sum(v for _, v in dep)
            tr.append("  임대보증금: " + ", ".join(f"{k} {v / 1e8:,.1f}억" for k, v in dep))
        if res["deposit_won"] is None:                      # 재무상태표에 별도 과목이 없으면(기타금융부채 등에 포함) 주석에서
            dep = deposit_from_notes(notes, unit_default, tr)
            if dep:
                res["deposit_won"] = dep
            elif st:
                res["deposit_won"] = 0.0
                tr.append("  임대보증금: 재무상태표·주석에 없음 → 0")
        # 대체: 주석에서 못 찾으면 재무상태표 과목
        if res["debt_won"] is None:
            d = [(k, v) for k, v in st.items() if re.search(r"차입금|사채", k) and not re.search(r"할인|이자|원가", k)]
            if d:
                res["debt_won"] = sum(v for _, v in d)
                res["flags"].append("차입금: 주석 파싱 실패 → 재무상태표 과목 합")
                tr.append("  차입금(재무상태표 대체): " + ", ".join(f"{k} {v / 1e8:,.1f}억" for k, v in d))
        if res["ip_won"] is None:
            d = [v for k, v in st.items() if k.startswith("투자부동산")]
            if d and d[0]:
                res["ip_won"], res["ip_basis"] = d[0], "재무상태표 장부금액"
                res["flags"].append("투자부동산 공정가치 주석 없음 → 재무상태표 장부금액")
        eq_bs = [(k, v) for k, v in st.items() if re.search(r"(종속|관계|공동)기업.*(투자|주식)", k) and not re.search(r"손상|처분", k)]
        if eq_bs:                                           # 기말 장부금액 = 재무상태표 금액(주석보다 확실)
            v = sum(x for _, x in eq_bs)
            if res["eq_won"] is not None and v and abs(res["eq_won"] - v) / v > 0.02:
                tr.append(f"  지분투자: 주석 {res['eq_won'] / 1e8:,.1f}억 대신 재무상태표 값 사용")
            res["eq_won"] = v
            tr.append("  지분투자(재무상태표): " + ", ".join(f"{k[:20]} {x / 1e8:,.1f}억" for k, x in eq_bs))
        if st and res["debt_won"] is None:
            res["debt_won"], res["contra_won"] = 0.0, 0.0
            tr.append("  차입금·사채: 재무상태표에 과목 없음 → 0")
        tot_assets = next((v for k, v in st.items() if re.match(r"자산총계", k)), None)
        core = (res["ip_won"] or 0) + (res["eq_won"] or 0)
        if tot_assets and core < 0.5 * tot_assets:          # 수익증권 등으로 부동산에 투자하는 리츠
            fv = [(k, v) for k, v in st.items() if re.search(r"비유동", k) and re.search(r"공정가치|당기손익", k) and v > 0]
            if fv:
                res["fvpl_won"] = sum(v for _, v in fv)
                res["flags"].append("자산가치에 비유동 공정가치측정 금융자산(수익증권 등) 포함")
                tr.append("  자산가치 보완(수익증권 등): " + ", ".join(f"{k[:24]} {v / 1e8:,.1f}억" for k, v in fv))
            elif core == 0:
                res["flags"].append("투자부동산·지분투자 없음(자산 매각 등)")
        res["total_assets_won"] = tot_assets
        break
    if res["basis"] is None:
        res["flags"].append("재무제표 주석 구역을 찾지 못함")
    return res


# ────────────────────────────────────────────────────────────── 배당
def dividend_info(xmls: list[str]) -> dict:
    text = " | ".join(b.get("text") or " | ".join(" | ".join(r) for r in b.get("rows", [])) for b in blocks_of(xmls[0]))
    flat = re.sub(r"\s+", " ", text)
    dps = None
    m = (re.search(r"1\s*주당\s*배당금.{0,80}?보통\s*주식?\s*\|\s*([\d,]+(?:\.\d+)?)", flat)
         or re.search(r"1\s*주당\s*배당금[^|]*\|\s*([\d,]+(?:\.\d+)?)", flat))
    if m:
        dps = float(m.group(1).replace(",", ""))
    base = None
    m = re.search(r"배당\s*기준일[^\d]{0,30}(\d{4})\s*[.\-년]\s*(\d{1,2})\s*[.\-월]\s*(\d{1,2})", flat)
    if m:
        base = dt.date(int(m[1]), int(m[2]), int(m[3]))
    months = None
    m = re.search(r"(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})\s*[~∼～]\s*(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})", flat)
    if m:
        a, b = dt.date(int(m[1]), int(m[2]), int(m[3])), dt.date(int(m[4]), int(m[5]), int(m[6]))
        months = round(((b - a).days + 1) / 30.44)
    yld = None
    m = (re.search(r"시가\s*배당\s*[율률].{0,60}?보통\s*주식?\s*\|\s*([\d.]+)", flat)
         or re.search(r"시가\s*배당\s*[율률][^|]*\|\s*([\d.]+)", flat))
    if m:
        yld = float(m.group(1))
    return {"dps": dps, "base": base.isoformat() if base else None, "period_months": months, "disc_yield": yld}


def is_div_decision(nm: str) -> bool:
    nm = nm.replace(" ", "")
    return ("배당결정" in nm or "배당 결정" in nm) and "주주명부" not in nm and "기준일" not in nm


def fiscal_months(regs) -> int | None:
    """정기보고서 이름의 기간(예: '사업보고서 (2026.06)')으로 결산기간 개월 수."""
    pts = []
    for f in regs:
        m = re.search(r"(사업|반기|분기)보고서\s*\((\d{4})\.(\d{2})\)", f.get("report_nm", ""))
        if m:
            pts.append((m.group(1), int(m.group(2)) * 12 + int(m.group(3))))
    biz = sorted({v for k, v in pts if k == "사업"}, reverse=True)
    if len(biz) >= 2 and 0 < biz[0] - biz[1] <= 12:
        return biz[0] - biz[1]
    if any(k == "분기" for k, _ in pts):
        return 12
    half = sorted({v for k, v in pts if k == "반기"}, reverse=True)
    if biz and half:
        gap = min(abs(biz[0] - h) for h in half)
        if 0 < gap <= 6:
            return gap * 2
    return None


def dividend_summary(divs, getdoc, override, close, fiscal=None):
    """최신 배당결정 1주당 배당금 + 결산기간(배당기준일 간격) → 연환산 시가배당수익률."""
    out = {"dps": None, "months": None, "yield": None, "yield_annualized": True, "flags": []}
    seen, recs = set(), []
    for f in divs[:6]:                                   # 최신순. 정정공시는 같은 기준일이면 최신 1건
        d = getdoc(f)
        key_ = d.get("base") or f["rcept_no"]
        if key_ in seen:
            continue
        seen.add(key_)
        recs.append((f, d))
        if len(recs) >= 2 and recs[0][1].get("base") and recs[1][1].get("base"):
            break
    f0, d0 = recs[0]
    out["dps"] = d0.get("dps")
    months, how = override.get("settle_months"), "설정값"
    if not months and d0.get("period_months"):
        months, how = _snap(d0["period_months"]), "공시 배당대상기간"
    gap = None
    if len(recs) >= 2 and d0.get("base") and recs[1][1].get("base"):
        g = (dt.date.fromisoformat(d0["base"]) - dt.date.fromisoformat(recs[1][1]["base"])).days / 30.44
        if abs(g - _snap(g)) <= 1.2:                    # 3·6·12개월 간격일 때만 신뢰
            gap = _snap(g)
    if not months and fiscal:
        months, how = fiscal, "정기보고서 결산기간"
        if gap and gap < fiscal:                         # 결산보다 잦은 분기·중간배당
            months, how = gap, f"배당기준일 간격(결산 {fiscal}개월보다 잦음)"
    if not months and gap:
        months, how = gap, f"배당기준일 간격 {recs[1][1]['base']}→{d0['base']}"
    if not months:
        n = len({d.get("base") for _, d in recs})
        months, how = 6, f"추정(배당결정 {n}건)"
        out["flags"].append("결산기간 6개월로 추정")
    out["months"] = months
    disc = d0.get("disc_yield")
    if out["dps"] and close:
        per = out["dps"] / close * 100
        if (disc or 0) >= 8 or per >= 8:                # 회차당 8% 이상 = 특별·청산 배당 가능성 → 연환산하지 않음
            out["yield"], out["yield_annualized"] = per, False
            out["flags"].append(f"일회성 배당 가능성(회차 {per:.1f}%, 연환산 안 함)")
        else:
            out["yield"] = per * 12 / months
    out["trace"] = (f"배당: {f0['report_nm'].strip()} {f0['rcept_dt']} 1주당 {out['dps']}원, 기준일 {d0.get('base')}, "
                    f"결산기간 {months}개월({how}), 공시 시가배당률 {disc}%")
    return out


def _snap(m):
    return min((3, 6, 12), key=lambda x: abs(x - m)) if m else None


# ────────────────────────────────────────────────────────────── 종목 1개
def reit_metrics(key: str, name: str, stock: str, close: float | None, mcap_eok: float | None,
                 corp: dict, cache: dict, override: dict, log, cd: float | None = None) -> dict:
    out = {"name": name, "report": None, "yield": None, "rate": None, "ltv": None, "nav_eok": None, "pnav": None,
           "flags": [], "dps": None, "months": None}
    info = corp.get(stock) or corp.get(str(override.get("dart_stock_code") or ""))
    if not info:                                   # 종목코드로 못 찾으면 회사명으로
        nm = lambda x: re.sub(r"\s|주식회사|\(주\)|㈜|(위탁관리|자기관리|기업구조조정)?부동산투자회사|리츠$", "", x).upper()  # noqa: E731
        want = nm(name)
        hits = [v for v in corp.values() if nm(v["corp_name"]) == want] or \
               [v for v in corp.values() if want and want in nm(v["corp_name"])]
        info = hits[0] if len(hits) == 1 else None
    if override.get("dart_corp_code"):
        info = {"corp_code": override["dart_corp_code"], "corp_name": name}
    if not info:
        out["flags"].append("DART 고유번호 없음")
        return out
    tr = [f"{name} ({stock}, DART {info['corp_code']} {info['corp_name']}) — {dt.datetime.now():%Y-%m-%d %H:%M}"]
    cc = info["corp_code"]
    # 정기보고서
    regs_all = [f for f in filings(key, cc, 800, "A") if re.search(r"사업보고서|반기보고서|분기보고서", f.get("report_nm", ""))]
    fiscal = fiscal_months(regs_all)
    cut = (dt.date.today() - dt.timedelta(days=400)).strftime("%Y%m%d")
    regs = [f for f in regs_all if re.search(r"사업보고서|반기보고서", f["report_nm"]) and f["rcept_dt"] >= cut]
    if regs:
        f = regs[0]
        out["report"] = f"{f['report_nm'].strip()} ({f['rcept_dt'][:4]}.{f['rcept_dt'][4:6]}.{f['rcept_dt'][6:]} 공시)"
        out["rcept_no"] = f["rcept_no"]
        tr.append(f"정기보고서: {f['report_nm']} 접수 {f['rcept_no']}")
        ck = f"{f['rcept_no']}#{PARSER_VERSION}"
        if ck in cache:
            m = cache[ck]["m"]
            tr += cache[ck]["trace"]
        else:
            t2: list[str] = []
            m = report_metrics(document(key, f["rcept_no"]), t2)
            cache[ck] = {"m": m, "trace": t2}
            tr += t2
    else:
        m = {"flags": ["최근 400일 사업·반기보고서 없음"]}
        tr.append("정기보고서 없음")
    m = {**m, **{k: v for k, v in override.items() if k.endswith("_won") or k == "rate"}}
    out["flags"] += m.get("flags", [])
    # 배당
    divs = [f for f in filings(key, cc, 800, "I") if is_div_decision(f.get("report_nm", ""))]
    if divs:
        def getdoc(f):
            k = f"div{PARSER_VERSION}#" + f["rcept_no"]
            if k not in cache:
                cache[k] = dividend_info(document(key, f["rcept_no"]))
            return cache[k]
        dv = dividend_summary(divs, getdoc, override, close, fiscal)
        out.update({k: dv[k] for k in ("dps", "months", "yield", "yield_annualized")})
        out["flags"] += dv["flags"]
        tr.append(dv["trace"])
    else:
        out["flags"].append("최근 배당결정 공시 없음")
    # 지표
    ds = debt_summary(m, cd)
    if "debt_won" in override:
        ds["debt_won"] = override["debt_won"]
    if "rate" in override:
        ds["rate"], ds["rate_cover"], ds["rate_est"] = override["rate"], 1.0, False
    debt, ip, eq, dep = ds["debt_won"], m.get("ip_won"), m.get("eq_won"), m.get("deposit_won")
    assets = (ip or 0) + (eq or 0) + (m.get("fvpl_won") or 0)
    out["rate"], out["rate_est"] = ds["rate"], ds["rate_est"]
    if ds["rate_cover"] is not None and ds["rate_cover"] < 0.95:
        out["flags"].append(f"금리 확인 차입금 {ds['rate_cover'] * 100:.0f}%만 반영")
    if ds["rate_est"]:
        out["flags"].append(f"변동금리는 CD {cd:.2f}%+가산금리로 추정")
    if debt is not None and assets:
        out["ltv"] = debt / assets * 100
    if assets and debt is not None:
        nav = assets - debt - (dep or 0)
        out["nav_eok"] = nav / 1e8
        if mcap_eok and nav > 0:
            out["pnav"] = mcap_eok / out["nav_eok"]
        if dep is None:
            out["flags"].append("임대보증금 미확인(0으로 계산)")
    out.update({"debt_eok": debt / 1e8 if debt is not None else None, "assets_eok": assets / 1e8 if assets else None,
                "ip_eok": ip / 1e8 if ip else None, "eq_eok": eq / 1e8 if eq else None,
                "deposit_eok": dep / 1e8 if dep is not None else None, "ip_basis": m.get("ip_basis"),
                "rate_cover": ds["rate_cover"]})
    tr.append(f"결과: 차입금·사채 {out['debt_eok']}억 / 투자부동산 {out['ip_eok']}억({out['ip_basis']}) / 지분투자 {out['eq_eok']}억 / "
              f"임대보증금 {out['deposit_eok']}억 / 가중금리 {out['rate']} / LTV {out['ltv']} / NAV {out['nav_eok']}억 / P/NAV {out['pnav']}")
    TRACE_DIR.mkdir(parents=True, exist_ok=True)
    (TRACE_DIR / f"{stock}_{name}.txt").write_text("\n".join(tr) + "\n", encoding="utf-8")
    return out


def collect_all(key: str, reits: list[dict], overrides: dict, log, cd: float | None = None) -> list[dict]:
    """reits: [{name, code, close, mcap_eok}] → 지표 목록. 오류는 종목 단위로 격리."""
    overrides = {k: v for k, v in overrides.items() if not k.startswith("_")}
    DART_DIR.mkdir(parents=True, exist_ok=True)
    cpath = DART_DIR / "metrics_cache.json"
    cache = json.loads(cpath.read_text(encoding="utf-8")) if cpath.exists() else {}
    cache = {k: v for k, v in cache.items() if f"#{PARSER_VERSION}" in k or k.startswith(f"div{PARSER_VERSION}#")}  # 옛 해석 결과 버림
    corp = corp_codes(key, log)
    res = []
    for r in reits:
        try:
            res.append(reit_metrics(key, r["name"], r["code"], r.get("close"), r.get("mcap_eok"), corp, cache,
                                    overrides.get(r["name"], {}), log, cd))
        except Exception as e:  # noqa: BLE001
            log(f"[DART] {r['name']}: {e}")
            res.append({"name": r["name"], "flags": [f"오류: {e}"]})
        time.sleep(0.2)
    cpath.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    return res
