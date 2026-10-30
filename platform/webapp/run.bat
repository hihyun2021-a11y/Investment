@echo off
chcp 65001 >nul
rem 신규투자 워크벤치 실행 (Windows). 처음 한 번은 가상환경을 만들고 패키지를 설치한다.
cd /d %~dp0
if not exist .venv (
  py -3 -m venv .venv
  .venv\Scripts\pip install -q -r requirements.txt
)
if not exist .env (
  copy .env.example .env >nul
  echo .env 파일을 만들었습니다. ANTHROPIC_API_KEY 를 넣고 다시 실행하세요.
  notepad .env
)
echo 브라우저에서 http://127.0.0.1:8765 을 여세요.
start "" http://127.0.0.1:8765
.venv\Scripts\python -m uvicorn app.main:app --host 127.0.0.1 --port 8765
