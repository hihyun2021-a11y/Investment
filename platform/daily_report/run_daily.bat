@echo off
chcp 65001 >nul
rem 상장리츠 공시·금리 일일 보고. 더블클릭하면 메뉴가 뜬다. 작업 스케줄러는 "run_daily.bat auto"로 실행한다.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

rem 1) Python 찾기
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

rem 2) 자동 실행(작업 스케줄러용): 화면 없이 run.log에 기록
if /i "%~1"=="auto" (
  echo ===== %date% %time% >> run.log
  %PY% reits_daily.py >> run.log 2>&1
  goto :eof
)
if /i "%~1"=="check" goto :check
if /i "%~1"=="preview" goto :preview

rem 3) 설정 파일 확인
if not exist "env.txt" (
  if exist "env.example.txt" copy "env.example.txt" "env.txt" >nul
)

:menu
echo.
echo  ==========================================
echo   상장리츠 공시·금리 일일 보고   (Python: %PY%)
echo  ==========================================
echo   1. 점검 (접속·수집 확인, 메일 안 보냄)
echo   2. 미리보기 (오늘 메일을 브라우저로 열기)
echo   3. 지금 메일 보내기
echo   4. 평일 07:50 자동 발송 등록
echo   5. 설정 파일(env.txt) 열기
echo   0. 끝내기
echo.
set "SEL="
set /p SEL=번호를 입력하고 Enter: 
if "%SEL%"=="1" goto :check
if "%SEL%"=="2" goto :preview
if "%SEL%"=="3" goto :send
if "%SEL%"=="4" goto :register
if "%SEL%"=="5" (notepad "env.txt" & goto :menu)
if "%SEL%"=="0" goto :eof
goto :menu

:check
echo.
echo [점검] 1~3분 걸릴 수 있습니다...
%PY% reits_daily.py --check
echo.
echo 위 결과에 경고가 있으면 이 창을 캡처해 보내 주세요.
pause
goto :menu

:preview
echo.
echo [미리보기] 23개 종목을 수집합니다. 2~5분 걸릴 수 있습니다...
%PY% reits_daily.py --preview
if exist "report_preview.html" start "" "report_preview.html"
pause
goto :menu

:send
echo.
echo [발송] 수집 후 메일을 보냅니다. 2~5분 걸릴 수 있습니다...
%PY% reits_daily.py
pause
goto :menu

:register
schtasks /create /tn ReitsDailyReport /tr "\"%~dp0run_daily.bat\" auto" /sc weekly /d MON,TUE,WED,THU,FRI /st 07:50 /f
if errorlevel 1 (echo 등록에 실패했습니다. 이 창을 캡처해 보내 주세요.) else (echo 등록 완료: 평일 07:50 자동 발송. 해제: schtasks /delete /tn ReitsDailyReport /f)
pause
goto :menu

:nopython
echo.
echo [안내] Python을 찾지 못했습니다.
echo  https://www.python.org/downloads/ 에서 Python 3.12 또는 3.13을 설치하세요.
echo  설치 첫 화면에서 "Add python.exe to PATH"를 꼭 체크하세요.
pause
