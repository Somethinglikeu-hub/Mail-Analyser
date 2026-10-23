@echo off
REM Collect new Outlook mails now (read-only) and push them to GitHub.
REM GitHub Actions then analyses them -> reports/ and DASHBOARD.md in the repo.
cd /d "%~dp0"
python mail_collector.py
echo.
powershell -NoProfile -Command "Get-Content '%~dp0run_log.txt' -Tail 2"
echo.
echo Done. Analysis runs on GitHub in ~1-2 min: open DASHBOARD.md in the repo.
pause
