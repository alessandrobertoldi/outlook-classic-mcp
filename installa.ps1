# installa.ps1 - installa o aggiorna outlook-classic-mcp sul PC e lo registra in Claude desktop.
#
# COME SI LANCIA (Windows PowerShell, come utente NORMALE, mai "come amministratore"):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\installa.ps1"
# dalla cartella in cui hai scaricato il repository. Outlook Classic deve essere aperto.
#
# COSA FA, IN ORDINE:
#   1. controlla che esista config.json (copiato da config.example.json e compilato);
#   2. crea un ambiente Python isolato (.venv) nella cartella del connettore, con uv se c'e', altrimenti con Python;
#   3. installa le due librerie necessarie: mcp (protocollo) e pywin32 (COM di Windows);
#   4. prova il collegamento COM con Outlook (deve essere aperto, non elevato);
#   5. prova ad avviare il server e stampa l'elenco degli strumenti;
#   6. registra il server nel file claude_desktop_config.json di Claude desktop, conservando gli altri server gia' presenti.
# Non tocca Outlook, non salva credenziali, non modifica altro.
#
# NOTA SUL FILE: va salvato in UTF-8 con BOM, altrimenti Windows PowerShell 5.1 legge male le lettere accentate.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host "== outlook-classic-mcp: installazione in $here"

# 1. configurazione personale
if (-not (Test-Path "$here\config.json")) {
    Write-Host ''
    Write-Host 'Manca config.json. Copia config.example.json in config.json, compila i valori fra < > e rilancia.' -ForegroundColor Yellow
    Write-Host "  Copy-Item `"$here\config.example.json`" `"$here\config.json`""
    exit 1
}
try {
    $cfgTest = (Get-Content "$here\config.json" -Raw) -replace '^\xEF\xBB\xBF', '' | ConvertFrom-Json
} catch {
    throw "config.json non è un JSON valido: $($_.Exception.Message)"
}
if (($cfgTest | ConvertTo-Json -Depth 10) -match '<[^>]+>') {
    Write-Host 'ATTENZIONE: in config.json ci sono ancora valori fra < > da sostituire.' -ForegroundColor Yellow
}

# 2. ambiente Python isolato: uv se disponibile (più veloce), altrimenti il launcher "py" o python.exe
$py = "$here\.venv\Scripts\python.exe"
$uv = Get-Command uv -ErrorAction SilentlyContinue
if (-not (Test-Path $py)) {
    if ($uv) {
        Write-Host '-- creo .venv con uv (Python 3.11)'
        & uv venv "$here\.venv" --python 3.11
    } else {
        $pyBase = $null
        foreach ($cand in @('py -3.12', 'py -3.11', 'python')) {
            try {
                $v = & cmd /c "$cand -c `"import sys;print(sys.version_info[:2]>=(3,11))`"" 2>$null
                if ($v -eq 'True') { $pyBase = $cand; break }
            } catch { }
        }
        if (-not $pyBase) { throw 'Serve Python 3.11 o più recente (https://www.python.org/downloads/windows/), oppure uv (winget install astral-sh.uv).' }
        Write-Host "-- creo .venv con $pyBase"
        & cmd /c "$pyBase -m venv `"$here\.venv`""
    }
}

# 3. librerie
Write-Host '-- installo/aggiorno mcp e pywin32'
if ($uv) {
    & uv pip install --python $py --upgrade "mcp>=2,<3" "pywin32>=306"
} else {
    & $py -m pip install --upgrade pip | Out-Null
    & $py -m pip install --upgrade "mcp>=2,<3" "pywin32>=306"
}
if ($LASTEXITCODE -ne 0) { throw 'Installazione delle librerie fallita: vedere errore sopra.' }

# 4. prova di collegamento COM (Outlook deve essere aperto)
Write-Host '-- prova COM con Outlook'
$test = @'
import win32com.client, sys
try:
    app = win32com.client.GetActiveObject("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    print("OK Outlook", app.Version, "profilo", ns.CurrentProfileName, "archivi:", ns.Stores.Count)
    for s in ns.Stores:
        print("   -", s.DisplayName)
except Exception as e:
    print("ERRORE:", e); sys.exit(1)
'@
$test | & $py -
if ($LASTEXITCODE -ne 0) { throw 'Collegamento COM fallito: Outlook Classic è aperto, come utente normale?' }
Write-Host '   I nomi qui sopra sono quelli da usare in config.json alla voce "caselle".'

# 5. prova di avvio del server (elenco strumenti)
Write-Host '-- prova di avvio del server MCP'
$probe = @'
import asyncio, sys
sys.path.insert(0, r"__HERE__")
import server
tools = asyncio.run(server.mcp.list_tools())
print("OK server, fase", server.FASE, "-", len(tools), "strumenti:", ", ".join(t.name for t in tools))
'@
$probe.Replace('__HERE__', $here) | & $py -
if ($LASTEXITCODE -ne 0) { throw 'Il server non parte: vedere errore sopra (di solito config.json).' }

# 6. registrazione in Claude desktop
# Il file di configurazione dei server MCP di Claude desktop è claude_desktop_config.json. Di solito sta in
# %APPDATA%\Claude; alcune installazioni usano un profilo proprio in %LOCALAPPDATA%\Claude-Data\<profilo>\.
# Per sapere quale usa il tuo: Claude desktop > Impostazioni > Sviluppatore > Modifica configurazione.
$candidati = @()
$std = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
if (Test-Path $std) { $candidati += $std }
$profili = Join-Path $env:LOCALAPPDATA 'Claude-Data'
if (Test-Path $profili) {
    Get-ChildItem $profili -Directory | ForEach-Object {
        $p = Join-Path $_.FullName 'claude_desktop_config.json'
        if (Test-Path $p) { $candidati += $p }
    }
}
if ($candidati.Count -eq 0) {
    $cfg = $std
    New-Item -ItemType Directory -Path (Split-Path $cfg) -Force | Out-Null
    Write-Host "-- nessuna configurazione esistente: creo $cfg"
} elseif ($candidati.Count -eq 1) {
    $cfg = $candidati[0]
} else {
    Write-Host 'Trovati più file di configurazione di Claude desktop:'
    for ($i = 0; $i -lt $candidati.Count; $i++) { Write-Host "  [$($i+1)] $($candidati[$i])" }
    $scelta = Read-Host 'Quale usare? (numero; verifica in Claude desktop > Impostazioni > Sviluppatore)'
    $cfg = $candidati[[int]$scelta - 1]
}

$json = $null
if (Test-Path $cfg) {
    Copy-Item $cfg "$cfg.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"   # copia di sicurezza
    $raw = (Get-Content $cfg -Raw) -replace '^\xEF\xBB\xBF', ''
    if ($raw.Trim()) { $json = $raw | ConvertFrom-Json }
}
if ($null -eq $json) { $json = New-Object PSObject }
if (-not ($json.PSObject.Properties.Name -contains 'mcpServers') -or $null -eq $json.mcpServers) {
    $json | Add-Member -NotePropertyName mcpServers -NotePropertyValue (New-Object PSObject) -Force
}
# La voce "outlook": comando = il Python dell'ambiente isolato, argomento = server.py. Le variabili d'ambiente
# forzano l'UTF-8 sui flussi standard, così gli accenti passano intatti fra server e Claude desktop.
$entry = [PSCustomObject]@{
    command = $py
    args    = [string[]]@("$here\server.py")
    env     = [PSCustomObject]@{ PYTHONIOENCODING = 'utf-8'; PYTHONUTF8 = '1' }
}
$json.mcpServers | Add-Member -NotePropertyName outlook -NotePropertyValue $entry -Force
$text = $json | ConvertTo-Json -Depth 10
# UTF-8 SENZA BOM: Claude desktop non legge un JSON con il BOM.
[IO.File]::WriteAllText($cfg, $text, (New-Object System.Text.UTF8Encoding $false))
Write-Host "-- registrato in $cfg"

Write-Host ''
Write-Host 'FATTO. Ora chiudi completamente Claude desktop (icona nell''area di notifica > Esci) e riaprilo.' -ForegroundColor Green
Write-Host 'Il server "outlook" comparirà fra i connettori locali. Per verificare, chiedi a Claude: "controlla lo stato di Outlook".'
Write-Host 'Prova completa delle funzioni di scrittura (consigliata): .\.venv\Scripts\python.exe .\prova_accenti.py'
