@echo off
:: Dash de Cotas - VERSAO NOVA (v2) para testes, na porta 8504.
:: NAO substitui o dash atual (status_mailers_v2.py na porta 8502).
cd /d "%~dp0"
start "" streamlit run status_mailers_v3.py --server.port 8504 --server.address 0.0.0.0
timeout /t 5 /nobreak >NUL
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" "http://localhost:8504"
