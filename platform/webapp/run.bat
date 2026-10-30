@echo off
chcp 65001 >nul
rem 신규투자 워크벤치 실행 (Windows). 처음 한 번은 가상환경을 만들고 패키지를 설치한다.
cd /d "%~dp0"

rem 1) Python 확인 (py 런처 우선)
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY goto :nopython

rem 2) 가상환경 만들기
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] 처음 실행 준비: 가상환경을 만드는 중입니다...
  %PY% -m venv .venv
  if errorlevel 1 goto :fail
)

rem 3) 패키지 설치 (설치가 끝까지 된 경우에만 표시 파일을 남김)
if not exist ".venv\installed.ok" (
  echo [2/3] 필요한 패키지를 설치하는 중입니다. 1~3분 걸립니다...
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :pipfail
  echo ok> ".venv\installed.ok"
)

rem 4) 설정 파일
if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo .env 파일을 만들었습니다. ANTHROPIC_API_KEY= 뒤에 키를 붙여넣고 저장한 뒤 메모장을 닫아 주세요.
  notepad ".env"
)

rem 5) 서버 시작 - 서버가 뜰 시간을 준 뒤 브라우저를 연다
echo [3/3] 워크벤치를 시작합니다. 이 검은 창은 사용하는 동안 닫지 마세요.
start "" cmd /c "timeout /t 8 >nul & start http://127.0.0.1:8765"
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8765
echo.
echo 서버가 멈췄습니다. 위에 나온 오류 메시지를 확인해 주세요.
pause
exit /b

:nopython
echo [오류] Python을 찾지 못했습니다.
echo python.org 에서 Python 3.11 이상을 설치하세요. 설치 첫 화면에서 "Add python.exe to PATH" 를 꼭 체크합니다.
pause
exit /b 1

:fail
echo [오류] 가상환경을 만들지 못했습니다. 위 메시지를 확인해 주세요.
pause
exit /b 1

:pipfail
echo [오류] 패키지 설치에 실패했습니다.
echo 회사망(프록시·보안 프로그램)에서 인터넷 설치가 막혔을 수 있습니다. 위 메시지를 IT팀에 전달해 주세요.
pause
exit /b 1
