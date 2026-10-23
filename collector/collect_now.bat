@echo off
REM Collect new Outlook mails now (read-only). Then tell Claude: "process new mails".
cd /d "%~dp0"
python mail_collector.py
echo.
powershell -NoProfile -Command "Get-Content '%~dp0run_log.txt' -Tail 1"
echo.
echo Done. Now tell Claude: process new mails
pause
