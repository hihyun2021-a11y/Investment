# -*- coding: utf-8 -*-
"""상장리츠별 최신 사업·반기보고서 2건 + 배당결정 공시 3건 원문(document.xml zip)을 받아 저장한다 — 파서 보정용.

DART_API_KEY=... python3 platform/daily_report/tools/fetch_dart_samples.py <저장폴더>
키는 환경변수로만 받는다(파일에 쓰지 않음).
"""
import json
import os
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE))
import dart_metrics as D  # noqa: E402

key = os.environ["DART_API_KEY"]
out = Path(sys.argv[1]).resolve()
out.mkdir(parents=True, exist_ok=True)
D.DART_DIR = out / "_cache"
cfg = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
corp = D.corp_codes(key, print)
nm = lambda x: re.sub(r"\s|주식회사|\(주\)|㈜|(위탁관리|자기관리|기업구조조정)?부동산투자회사|리츠$", "", x).upper()  # noqa: E731
index = []
for r in cfg["reits"]:
    name, code = r["name"], r.get("code")
    info = corp.get(code) or next((v for v in corp.values() if nm(v["corp_name"]) == nm(name)), None)
    if not info:
        print("고유번호 없음:", name, code)
        index.append({"name": name, "code": code, "error": "corp_code not found"})
        continue
    item = {"name": name, "code": code, **info, "files": []}
    try:
        regs = [f for f in D.filings(key, info["corp_code"], 400, "A") if re.search(r"사업보고서|반기보고서|분기보고서", f["report_nm"])]
        divs = [f for f in D.filings(key, info["corp_code"], 800, "I") if "배당" in f["report_nm"] and "결정" in f["report_nm"]]
        for kind, lst, n in (("정기", regs, 2), ("배당", divs, 3)):
            for f in lst[:n]:
                raw_dir = D.DART_DIR
                D.document(key, f["rcept_no"])
                src = raw_dir / f"{f['rcept_no']}.zip"
                dst = out / f"{code}_{name}_{kind}_{f['rcept_no']}.zip"
                dst.write_bytes(src.read_bytes())
                item["files"].append({"kind": kind, "file": dst.name, "report_nm": f["report_nm"], "rcept_dt": f["rcept_dt"],
                                      "rcept_no": f["rcept_no"], "bytes": dst.stat().st_size})
                print(f"{name}: {kind} {f['report_nm']} {f['rcept_dt']} {dst.stat().st_size:,}B")
                time.sleep(0.3)
    except Exception as e:  # noqa: BLE001
        item["error"] = str(e)
        print("오류:", name, e)
    index.append(item)
(out / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
for p in (out / "_cache").glob("*"):
    p.unlink()
(out / "_cache").rmdir()
print("완료:", out)
