@echo off
:: ============================================================
:: Fixa o IP desta maquina (servidor do Dash de Cotas) em
:: 192.168.3.83 (ESTATICO). Evita que o DHCP troque o IP num
:: reboot e derrube o acesso do time (inclusive nas ferias da RI).
::
:: >>> EXECUTAR COMO ADMINISTRADOR <<<
:: (botao direito neste arquivo > "Executar como administrador")
::
:: Para reverter: voltar_ip_dhcp.bat (tambem como administrador).
:: ============================================================
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo.
    echo  ERRO: este script precisa ser executado como ADMINISTRADOR.
    echo  Feche esta janela e clique com o botao direito ^>
    echo  "Executar como administrador".
    echo.
    pause
    exit /b 1
)

echo Fixando IP 192.168.3.83 / mascara 255.255.255.0 / gateway 192.168.3.249 (Ethernet)...
netsh interface ip set address name="Ethernet" static 192.168.3.83 255.255.255.0 192.168.3.249

echo Configurando DNS (10.1.0.75 / 192.168.3.249 / 192.168.3.58)...
netsh interface ip set dns name="Ethernet" static 10.1.0.75 primary
netsh interface ip add dns name="Ethernet" 192.168.3.249 index=2
netsh interface ip add dns name="Ethernet" 192.168.3.58 index=3

echo.
echo === Configuracao aplicada. IPv4 atual: ===
ipconfig | findstr /i "IPv4"
echo.
echo  Pronto. O IP esta FIXO em 192.168.3.83 - nao muda mais em reboot.
echo  Dash: http://192.168.3.83:8502
echo.
pause
