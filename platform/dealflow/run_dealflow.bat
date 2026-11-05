@echo off
rem Dealflow launcher (stage 1). ASCII only on purpose (cmd.exe UTF-8 issue). Menu is in Python.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
set "PY="
for %%V in (3.13 3.12 3.11) do (
  if not defined PY (
    py -%%V -c "import sys" >nul 2>nul && set "PY=py -%%V"
  )
)
if not defined PY goto :nopython
if not exist ".venv\Scripts\python.exe" (
  echo [setup 1/2] Creating virtual environment with %PY% ...
  %PY% -m venv .venv
  if errorlevel 1 goto :fail
)
if not exist ".venv\installed_v1.ok" (
  echo [setup 2/2] Installing packages. This takes 1-2 minutes, only once ...
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :pipfail
  echo ok> ".venv\installed_v1.ok"
)
".venv\Scripts\python.exe" dealflow.py %*
pause
goto :eof

:nopython
echo.
echo [!] Python 3.11 - 3.13 not found.
echo     Open a command prompt and run:   py install 3.13
echo     or install Python 3.13 from https://www.python.org/downloads/  then run this file again.
pause
goto :eof

:pipfail
echo.
echo [!] Package install failed. If the company network blocks pip, ask IT to allow pypi.org,
echo     or capture this window and send it to Claude.
pause
goto :eof

:fail
echo [!] Failed to create the virtual environment. Capture this window and send it to Claude.
pause
