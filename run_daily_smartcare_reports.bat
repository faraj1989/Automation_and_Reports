@echo off
rem Usage: run_daily_smartcare_reports.bat [cem ^| penetration]
rem   cem          - SmartCare CEM export + analysis   (daily task, 05:00)
rem   penetration  - Weekly Device Penetration export  (weekly task, Monday 01:00 - export then covers the full Mon-Sun week)
rem   (no argument) - both, one after the other
cd /d "C:\Users\user\Desktop\python\Libyana Daily Report - Copy"
set PY="C:\Users\user\Desktop\python\Libyana Daily Report - Copy\.venv\Scripts\python.exe"
set PYTHONIOENCODING=utf-8
set MODE=%~1
if "%MODE%"=="" set MODE=all
if not exist logs mkdir logs
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set STAMP=%%I
set LOG=logs\smartcare_reports_%STAMP%_%MODE%.log

rem The CEM and Device Penetration scrapers share ONE SmartCare account and its
rem Async Export task list, so they must never run in parallel. Each scraper
rem only downloads its own task name prefix, and the two tasks are scheduled
rem hours apart.
echo ============================================================ >> %LOG%
echo  SmartCare Reports (%MODE%) - %date% %time% >> %LOG%
echo ============================================================ >> %LOG%

if /i "%MODE%"=="penetration" goto penetration

echo  Step: SmartCare CEM export (download only) >> %LOG%
%PY% "scrapers\smartcare_cem_scraper.py" --skip-analysis >> %LOG% 2>&1
echo. >> %LOG%
echo  Step: SmartCare CEM analysis (folds new exports into history) >> %LOG%
%PY% "reports\run_smartcare_analysis_task.py" >> %LOG% 2>&1
if /i "%MODE%"=="cem" goto done

:penetration
echo. >> %LOG%
echo  Step: Device Penetration export + history >> %LOG%
%PY% "scrapers\weekly_device_penetration_scraper.py" >> %LOG% 2>&1

:done
echo. >> %LOG%
echo  Finished %date% %time% >> %LOG%
