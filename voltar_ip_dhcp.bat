@echo off
:: ============================================================
:: REVERSAO: volta a interface Ethernet para DHCP (IP automatico).
:: Use somente se a fixacao de IP (fixar_ip_dash.bat) causar problema.
:: >>> EXECUTAR COMO ADMINISTRADOR <<<
:: ============================================================
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo  ERRO: execute como ADMINISTRADOR (botao direito ^> Executar como administrador).
    pause
    exit /b 1
)

echo Revertendo Ethernet para DHCP (IP e DNS automaticos)...
netsh interface ip set address name="Ethernet" dhcp
netsh interface ip set dns name="Ethernet" dhcp

echo.
echo === IPv4 atual: ===
ipconfig | findstr /i "IPv4"
echo.
pause
