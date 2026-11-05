@echo off
rem Register weekday 09:10 auto run (once). Remove: schtasks /delete /tn ReitsDailyReport /f
schtasks /create /tn ReitsDailyReport /tr "\"%~dp0run_daily.bat\" auto" /sc weekly /d MON,TUE,WED,THU,FRI /st 09:10 /f
pause
