@echo off
:: ============================================================
:: Watchdog do Dash de Cotas (Streamlit, porta 8502)
:: Mantem o dash SEMPRE no ar, todos os dias: se o Streamlit cair
:: por qualquer motivo, reinicia automaticamente em 15s. Loop infinito.
::
:: Colocar na pasta Startup do Windows. NAO usar junto com o
:: auto_dash.bat (os dois subiriam o Streamlit e brigariam pela porta
:: 8502) - por isso o auto_dash.bat do Startup foi desativado (.disabled).
:: ============================================================
cd /d "C:\Users\lucieli.andrade\OneDrive - Capitania S.A\DASH_2026\API_EMAIL_COTAS_DASH"

set "STREAMLIT=C:\Users\lucieli.andrade\AppData\Local\Programs\Python\Python314\Scripts\streamlit.exe"
set "LOG=dash_watchdog_log.txt"

:: Abre o dash no Chrome uma vez, ao subir o watchdog
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" "http://192.168.3.78:8502"

:loop
echo [%date% %time%] Iniciando Streamlit (porta 8502)...>> "%LOG%"
"%STREAMLIT%" run "status_mailers_v2.py" --server.port 8502 --server.address 0.0.0.0
echo [%date% %time%] Streamlit encerrou (exit %ERRORLEVEL%). Reiniciando em 15s...>> "%LOG%"
timeout /t 15 /nobreak >NUL
goto loop
