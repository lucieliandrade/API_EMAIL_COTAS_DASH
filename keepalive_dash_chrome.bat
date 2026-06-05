@echo off
:: ============================================================
:: Keepalive da janela do Dash no Chrome.
:: Mantem o keepalive_dash_chrome.ps1 rodando 24/7 (qualquer horario).
:: Se a janela do dash for fechada, reabre em ate 60s.
:: Colocar na pasta Startup do Windows (sobe junto com o watchdog_dash.bat).
:: ============================================================
set "BASE=C:\Users\lucieli.andrade\OneDrive - Capitania S.A\DASH_2026\API_EMAIL_COTAS_DASH"

:loop
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%BASE%\keepalive_dash_chrome.ps1"
echo [%date% %time%] keepalive_dash_chrome.ps1 encerrou (exit %ERRORLEVEL%). Reiniciando em 30s...>> "%BASE%\dash_watchdog_log.txt"
timeout /t 30 /nobreak >NUL
goto loop
