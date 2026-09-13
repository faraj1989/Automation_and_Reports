@echo off
cd /d "C:\Users\user\Desktop\python\Libyana Daily Report - Copy"
echo ============================================================
echo  Daily SmartCare Reports (CEM + Device Penetration)
echo  Step 1/3: SmartCare CEM export + analysis
echo  Step 2/3: Weekly Device Penetration export
echo ============================================================
"C:\Users\user\Desktop\python\Libyana Daily Report - Copy\.venv\Scripts\python.exe" "C:\Users\user\Desktop\python\Libyana Daily Report - Copy\scrapers\smartcare_cem_scraper.py"
"C:\Users\user\Desktop\python\Libyana Daily Report - Copy\.venv\Scripts\python.exe" "C:\Users\user\Desktop\python\Libyana Daily Report - Copy\reports\run_smartcare_analysis_task.py"
"C:\Users\user\Desktop\python\Libyana Daily Report - Copy\.venv\Scripts\python.exe" "C:\Users\user\Desktop\python\Libyana Daily Report - Copy\scrapers\weekly_device_penetration_scraper.py"
