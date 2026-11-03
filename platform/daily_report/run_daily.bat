@echo off
rem REITs daily report launcher. Double-click: menu. Task Scheduler: run_daily.bat auto
rem (ASCII only on purpose: cmd.exe misreads some UTF-8 Korean lines)
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
set "PY="
for %%V in (3.13 3.12 3.11 3.10 3) do (
  if not defined PY (
    py -%%V -c "import sys" >nul 2>nul && set "PY=py -%%V"
  )
)
if not defined PY (
  python -c "import sys" >nul 2>nul && set "PY=python"
)
if not defined PY goto :nopython
if /i "%~1"=="auto" (
  echo ===== %date% %time% >> run.log
  %PY% reits_daily.py >> run.log 2>&1
  goto :eof
)
%PY% reits_daily.py --menu
goto :eof

:nopython
echo.
echo [!] Python not found. Install Python 3.12 or 3.13 from https://www.python.org/downloads/
echo     and check "Add python.exe to PATH" on the first screen.
pause
