@echo off
:: ============================================================
:: Watchdog do Dash de ROTINAS DIARIAS (Streamlit, porta 8503)
:: Mantem o dash SEMPRE no ar: se o Streamlit cair por qualquer
:: motivo, reinicia automaticamente em 15s. Loop infinito.
::
:: Independente do dash de cotas (porta 8502) - um nao derruba o
:: outro. Colocar na pasta Startup do Windows.
:: ============================================================
cd /d "C:\Users\lucieli.andrade\OneDrive - Capitania S.A\DASH_2026\API_EMAIL_COTAS_DASH"

set "STREAMLIT=C:\Users\lucieli.andrade\AppData\Local\Programs\Python\Python314\Scripts\streamlit.exe"
set "LOG=rotinas_watchdog_log.txt"

:loop
echo [%date% %time%] Iniciando Streamlit Rotinas (porta 8503)...>> "%LOG%"
"%STREAMLIT%" run "dash_rotinas.py" --server.port 8503 --server.address 0.0.0.0
echo [%date% %time%] Streamlit encerrou (exit %ERRORLEVEL%). Reiniciando em 15s...>> "%LOG%"
timeout /t 15 /nobreak >NUL
goto loop
