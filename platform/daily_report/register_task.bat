@echo off
chcp 65001 > nul
rem 평일 07:50 자동 실행 등록 (한 번만 실행). 해제: schtasks /delete /tn ReitsDailyReport /f
schtasks /create /tn ReitsDailyReport /tr "\"%~dp0run_daily.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 07:50 /f
echo 등록 완료: 평일 07:50 ReitsDailyReport
pause
