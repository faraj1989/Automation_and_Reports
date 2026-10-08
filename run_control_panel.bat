@echo off
rem Private control panel - listens on this PC only (127.0.0.1), never the LAN.
cd /d "C:\Users\user\Desktop\python\Libyana Daily Report - Copy"
".venv\Scripts\python.exe" -m streamlit run control_panel\app.py --server.address 127.0.0.1 --server.port 8502 --browser.gatherUsageStats false --server.headless true
