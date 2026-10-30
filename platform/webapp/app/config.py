# -*- coding: utf-8 -*-
"""앱 설정 — webapp/.env 파일(있으면)과 환경변수에서 읽는다. API 키는 코드에 넣지 않는다."""
import os
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1]          # platform/webapp
ROOT = APP_DIR.parents[1]                                # 저장소 루트(Investment)
DATA = APP_DIR / "data"                                  # 로컬 DB·업로드·결과물 (git 제외)


def _load_env(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env(APP_DIR / ".env")

MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-5-5")
EFFORT = os.environ.get("CLAUDE_EFFORT", "high")        # low | medium | high | xhigh | max
HOST = os.environ.get("APP_HOST", "127.0.0.1")          # 로컬 전용. 외부 접속은 README 2단계 참고
PORT = int(os.environ.get("APP_PORT", "8765"))
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")       # 설정하면 접속 시 비밀번호(ID 아무거나) 요구
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")
MAX_SAMPLE_CHARS = int(os.environ.get("MAX_SAMPLE_CHARS", "200000"))  # 샘플 1건당 텍스트 상한

for sub in ("samples", "cache", "outputs"):
    (DATA / sub).mkdir(parents=True, exist_ok=True)
