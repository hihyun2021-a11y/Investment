# -*- coding: utf-8 -*-
"""OpenDART로 상장리츠 재무 지표 4종을 계산한다 (reits_daily.py에서 호출, 표준 라이브러리만 사용).

1. 시가배당수익률(연환산) = 최근 '부동산투자회사금전배당결정' 1주당 배당금 × (12 ÷ 결산기간 개월) ÷ 전일 종가
2. 가중평균 차입금리     = 최근 사업·반기보고서 「재무제표 주석」 차입금·사채 표의 당기말 장부금액 가중평균 이자율
3. LTV                 = 차입금·사채 당기말 합계 ÷ 자산가치
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
PARSER_VERSION = 4
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
        self.head_rows = 0

    def _flush(self):
        t = re.sub(r"\s+", " ", "".join(self.buf)).strip()
        if t:
            self.blocks.append({"kind": self.kind, "text": t})
        self.buf, self.kind = [], "p"

    def handle_starttag(self, tag, attrs):
        a = {k.lower(): (v or "") for k, v in attrs}
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
    prev = [b["text"] for b in blocks[max(0, i - 3):i] if b["kind"] != "table"]
    return " ".join(prev)


# ────────────────────────────────────────────────────────────── 주석 구역
HEAD = re.compile(r"^\s*(\d{1,2})\s*[.．]\s*(\S.*)$")
KEY_DEBT = re.compile(r"차입금|사채")
KEY_IP = re.compile(r"투자부동산")
KEY_EQ = re.compile(r"종속기업|관계기업|공동기업|지분법")


def notes_section(blocks):
    """별도 「재무제표 주석」 구역. 없으면 연결 주석, 그것도 없으면 None."""
    idx = [i for i, b in enumerate(blocks) if b["kind"] == "title"]
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
    idx = [i for i, b in enumerate(blocks) if b["kind"] == "title"]
    for k, i in enumerate(idx):
        t = re.sub(r"\s+", "", blocks[i]["text"])
        ok = re.search(r"\d\.(연결)?재무제표$", t) and (("연결" in t) == (basis == "연결"))
        if ok:
            end = idx[k + 1] if k + 1 < len(idx) else len(blocks)
            return blocks[i + 1:end]
    return None


def split_notes(sec):
    """[(번호, 제목, 블록목록)] — 주석 번호가 커지는 순서로만 제목을 인정(본문 안의 '1. …' 목록 오인 방지)."""
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


def debt_from_note(note, unit_default, tr):
    """[{label, won, rate, base, spread}] + 차감 계정(사채할인발행차금 등, 음수) 합."""
    rows, seen, contra = [], set(), 0.0
    bl = note["blocks"]
    for i, b in enumerate(bl):
        if b["kind"] != "table":
            continue
        t = T(b, _context(bl, i))
        if t.is_prior_only():
            tr.append(f"    표 건너뜀(전기): {t.text[:60]}")
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
            rate, sp = rate_of(r[rc[0]]), spread_of(r[rc[0]])
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
    gross = sum(r["won"] for r in rows)
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
        v = value_from_table(t, unit_default, r"장부|기말", tr, "지분투자 장부금액")
        if v:
            return v
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
                items.setdefault(lab, v * unit)
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
        dep = [(k, v) for k, v in st.items() if "임대보증금" in k and not re.search(r"할인|현재가치", k)]
        if dep:
            res["deposit_won"] = sum(v for _, v in dep)
            tr.append("  임대보증금: " + ", ".join(f"{k} {v / 1e8:,.1f}억" for k, v in dep))
        # 대체: 주석에서 못 찾으면 재무상태표 과목
        if res["debt_won"] is None:
            d = [(k, v) for k, v in st.items() if re.search(r"차입금|사채", k) and not re.search(r"할인|이자|원가", k)]
            if d:
                res["debt_won"] = sum(v for _, v in d)
                res["flags"].append("차입금: 주석 파싱 실패 → 재무상태표 과목 합")
                tr.append("  차입금(재무상태표 대체): " + ", ".join(f"{k} {v / 1e8:,.1f}억" for k, v in d))
        if res["ip_won"] is None:
            d = [v for k, v in st.items() if k.startswith("투자부동산")]
            if d:
                res["ip_won"], res["ip_basis"] = d[0], "재무상태표 장부금액"
                res["flags"].append("투자부동산: 주석 파싱 실패 → 재무상태표 장부금액")
        if res["eq_won"] is None:
            d = [v for k, v in st.items() if re.search(r"^(종속|관계|공동)기업", k)]
            if d:
                res["eq_won"] = sum(d)
                tr.append(f"  지분투자(재무상태표): {sum(d) / 1e8:,.1f}억")
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
    m = re.search(r"시가\s*배당\s*율[^|]*\|\s*(?:보통\s*주식?\s*\|\s*)?([\d.]+)", flat)
    if m:
        yld = float(m.group(1))
    return {"dps": dps, "base": base.isoformat() if base else None, "period_months": months, "disc_yield": yld}


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
    regs = [f for f in filings(key, cc, 400, "A") if re.search(r"사업보고서|반기보고서", f.get("report_nm", ""))]
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
    divs = [f for f in filings(key, cc, 800, "I") if "배당" in f.get("report_nm", "") and "결정" in f.get("report_nm", "")]
    if divs:
        d0 = cache.get("div#" + divs[0]["rcept_no"]) or dividend_info(document(key, divs[0]["rcept_no"]))
        cache["div#" + divs[0]["rcept_no"]] = d0
        out["dps"] = d0["dps"]
        months = override.get("settle_months") or d0.get("period_months")
        how = "설정값" if override.get("settle_months") else ("공시 기간" if months else None)
        if not months and len(divs) >= 2:
            d1 = cache.get("div#" + divs[1]["rcept_no"]) or dividend_info(document(key, divs[1]["rcept_no"]))
            cache["div#" + divs[1]["rcept_no"]] = d1
            if d0.get("base") and d1.get("base"):
                gap = (dt.date.fromisoformat(d0["base"]) - dt.date.fromisoformat(d1["base"])).days / 30.44
                months, how = _snap(gap), "직전 배당기준일 간격"
        if not months:
            n = sum(1 for f in divs if f["rcept_dt"] >= (dt.date.today() - dt.timedelta(days=365)).strftime("%Y%m%d"))
            months, how = ({4: 3, 3: 3, 2: 6, 1: 12}.get(min(n, 4)) or 6), f"최근 1년 배당결정 {n}건"
        out["months"] = _snap(months) if how != "설정값" else months
        tr.append(f"배당: {divs[0]['report_nm']} {divs[0]['rcept_dt']} 1주당 {d0['dps']}원, 기준일 {d0.get('base')}, "
                  f"결산기간 {out['months']}개월({how}), 공시 시가배당율 {d0.get('disc_yield')}%")
        if d0["dps"] and close:
            out["yield"] = d0["dps"] * 12 / out["months"] / close * 100
    else:
        out["flags"].append("최근 배당결정 공시 없음")
    # 지표
    ds = debt_summary(m, cd)
    if "debt_won" in override:
        ds["debt_won"] = override["debt_won"]
    if "rate" in override:
        ds["rate"], ds["rate_cover"], ds["rate_est"] = override["rate"], 1.0, False
    debt, ip, eq, dep = ds["debt_won"], m.get("ip_won"), m.get("eq_won"), m.get("deposit_won")
    assets = (ip or 0) + (eq or 0)
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
