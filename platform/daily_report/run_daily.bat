@echo off
chcp 65001 > nul
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
if "%1"=="check" (%PY% reits_daily.py --check & pause & goto :eof)
if "%1"=="preview" (%PY% reits_daily.py --preview & start "" report_preview.html & goto :eof)
%PY% reits_daily.py >> run.log 2>&1
