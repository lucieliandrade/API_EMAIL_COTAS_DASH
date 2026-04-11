@echo off
cd /d "%~dp0"

echo Iniciando dashboard de cotas...
echo Acesse: http://192.168.3.78:8502
echo Login: RI
echo.
echo Para fechar, pressione Ctrl+C ou feche esta janela.
echo.

streamlit run status_mailers_v2.py --server.port 8502 --server.address 0.0.0.0
pause
