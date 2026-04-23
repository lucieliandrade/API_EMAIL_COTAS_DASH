@echo off
:: Watchdog do robo de mailers - monitora robo_log.txt e alerta se ficar parado.
:: Roda em loop: se o Python morrer por qualquer motivo, reinicia em 30s.
:: Colocar este .bat na pasta Startup do Windows junto com mailer_robo.bat.
:loop
"C:\Users\lucieli.andrade\AppData\Local\Programs\Python\Python314\python.exe" "C:\Users\lucieli.andrade\OneDrive - Capitania S.A\DASH_2026\API_EMAIL_COTAS_DASH\watchdog_robo.py"
echo [%date% %time%] Watchdog encerrou. Reiniciando em 30s...
timeout /t 30 /nobreak >NUL
goto loop
