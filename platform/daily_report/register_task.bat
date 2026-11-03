@echo off
chcp 65001 >nul
rem 평일 07:50 자동 발송 등록(한 번만 실행). 해제: schtasks /delete /tn ReitsDailyReport /f
schtasks /create /tn ReitsDailyReport /tr "\"%~dp0run_daily.bat\" auto" /sc weekly /d MON,TUE,WED,THU,FRI /st 07:50 /f
if errorlevel 1 (echo 등록에 실패했습니다.) else (echo 등록 완료: 평일 07:50 ReitsDailyReport)
pause
