@echo off
:: Verifica se o Streamlit ja esta rodando na porta 8502
tasklist /FI "IMAGENAME eq streamlit.exe" 2>NUL | find /I "streamlit.exe" >NUL
if %ERRORLEVEL%==0 (
    exit /b
)

:: Inicia o dash
cd /d "%~dp0"
start "" streamlit run status_mailers_v2.py --server.port 8502 --server.address 0.0.0.0

:: Aguarda 5 segundos e abre no Chrome
timeout /t 5 /nobreak >NUL
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" "http://192.168.3.78:8502"
