@echo off
cd /d "C:\Users\user\Desktop\python\Libyana Daily Report - Copy"

set ETH_IP=
for /f "usebackq tokens=1,* delims=:" %%A in (`powershell -NoProfile -Command "(Get-NetIPAddress -InterfaceAlias 'Ethernet' -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress"`) do set ETH_IP=%%A

set WIFI_IP=
for /f "usebackq tokens=1,* delims=:" %%A in (`powershell -NoProfile -Command "(Get-NetIPAddress -InterfaceAlias 'Wi-Fi' -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress"`) do set WIFI_IP=%%A

echo ============================================================
echo  Libyana Network Dashboard
echo.
if not "%ETH_IP%"=="" (
    echo  Ethernet / Intranet - fixed IP, use this for RF/NOC on the LAN:
    echo    http://%ETH_IP%:8501
    echo.
)
if not "%WIFI_IP%"=="" (
    echo  Wi-Fi:
    echo    http://%WIFI_IP%:8501
    echo.
)
if "%ETH_IP%%WIFI_IP%"=="" (
    echo  [!] Could not detect an IP on Ethernet or Wi-Fi - falling back to 0.0.0.0
    echo      Dashboard will still be reachable at whatever IP this PC has on the network.
)
echo  Press Ctrl+C to stop the dashboard.
echo ============================================================
echo.

".venv\Scripts\python.exe" -m streamlit run streamlit_dashboard.py --server.address 0.0.0.0 --server.port 8501 --browser.gatherUsageStats false
