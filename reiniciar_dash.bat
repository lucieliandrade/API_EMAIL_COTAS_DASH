@echo off
cd /d "%~dp0"

echo Iniciando dashboard de cotas...
echo Acesse: http://RI011:8502  (nome fixo - nao muda com o IP)
echo Login: RI
echo.
echo Para fechar, pressione Ctrl+C ou feche esta janela.
echo.

streamlit run status_mailers_v2.py --server.port 8502 --server.address 0.0.0.0
pause
