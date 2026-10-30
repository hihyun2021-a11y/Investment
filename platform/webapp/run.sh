#!/usr/bin/env bash
# 신규투자 워크벤치 실행 (macOS·Linux). 처음 한 번은 가상환경을 만들고 패키지를 설치한다.
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv && ./.venv/bin/pip install -q -r requirements.txt
fi
[ -f .env ] || { cp .env.example .env; echo ".env 를 만들었습니다. ANTHROPIC_API_KEY 를 넣고 다시 실행하세요."; }
HOST=$(grep -E '^APP_HOST=' .env | cut -d= -f2); PORT=$(grep -E '^APP_PORT=' .env | cut -d= -f2)
echo "브라우저에서 http://${HOST:-127.0.0.1}:${PORT:-8765} 을 여세요."
exec ./.venv/bin/python -m uvicorn app.main:app --host "${HOST:-127.0.0.1}" --port "${PORT:-8765}"
