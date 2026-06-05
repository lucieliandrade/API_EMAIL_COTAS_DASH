# keepalive_dash_chrome.ps1
# Mantem a janela do Dash de Cotas SEMPRE aberta no Chrome.
# A cada 60s verifica se a janela dedicada (modo --app, perfil proprio) existe;
# se foi fechada, reabre. Deteccao confiavel via --user-data-dir dedicado
# (a janela --app roda como instancia propria, com cmdline identificavel).
$url    = 'http://192.168.3.83:8502'
$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$prof   = Join-Path $env:LOCALAPPDATA 'DashCotasChrome'

while ($true) {
    try {
        $aberto = Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" -ErrorAction Stop |
                  Where-Object { $_.CommandLine -like '*DashCotasChrome*' }
        if (-not $aberto) {
            Start-Process $chrome -ArgumentList "--app=$url", "--user-data-dir=`"$prof`""
        }
    } catch { }
    Start-Sleep -Seconds 60
}
