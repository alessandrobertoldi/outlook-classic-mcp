# install.ps1 - installs or updates outlook-classic-mcp on this PC and registers it in Claude desktop.
#
# HOW TO RUN (Windows PowerShell, as a NORMAL user, never "as administrator"):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
# from the folder where you downloaded the repository. Classic Outlook must be open.
#
# WHAT IT DOES, IN ORDER:
#   1. checks that config.json exists (copied from config.example.json and filled in);
#   2. creates an isolated Python environment (.venv) in the server folder, with uv if present, otherwise with Python;
#   3. installs the two required libraries: mcp (the protocol) and pywin32 (Windows COM);
#   4. tests the COM connection to Outlook (it must be open, not elevated) and prints the store names for config.json;
#   5. test-starts the server and prints the list of tools;
#   6. registers the server in Claude desktop's claude_desktop_config.json, keeping any other server already there.
# It does not touch Outlook, does not store credentials, changes nothing else.

$ErrorActionPreference = 'Stop'
# Text piped to Python must travel as UTF-8, otherwise an install path with accented letters becomes '?'.
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONUTF8 = '1'   # and Python must read that pipe as UTF-8 too (stdin from a pipe is not covered by PEP 528)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host "== outlook-classic-mcp: installing in $here"

# 1. personal configuration
if (-not (Test-Path "$here\config.json")) {
    Write-Host ''
    Write-Host 'config.json is missing. Copy config.example.json to config.json, fill in the values in < > and run again.' -ForegroundColor Yellow
    Write-Host "  Copy-Item `"$here\config.example.json`" `"$here\config.json`""
    exit 1
}
try {
    $cfgTest = (Get-Content "$here\config.json" -Raw -Encoding UTF8) | ConvertFrom-Json
} catch {
    throw "config.json is not valid JSON: $($_.Exception.Message)"
}
if (($cfgTest | ConvertTo-Json -Depth 10) -match '<[^>]+>') {
    Write-Host 'WARNING: config.json still contains values in < > to be replaced.' -ForegroundColor Yellow
}

# 2. isolated Python environment: uv when available (faster), otherwise the "py" launcher or python.exe
$py = "$here\.venv\Scripts\python.exe"
$uv = Get-Command uv -ErrorAction SilentlyContinue
if (-not (Test-Path $py)) {
    if ($uv) {
        Write-Host '-- creating .venv with uv (Python 3.11)'
        & uv venv "$here\.venv" --python 3.11
    } else {
        $pyBase = $null
        foreach ($cand in @('py -3.12', 'py -3.11', 'python')) {
            try {
                $v = & cmd /c "$cand -c `"import sys;print(sys.version_info[:2]>=(3,11))`"" 2>$null
                if ($v -eq 'True') { $pyBase = $cand; break }
            } catch { }
        }
        if (-not $pyBase) { throw 'Python 3.11 or newer is required (https://www.python.org/downloads/windows/), or uv (winget install astral-sh.uv).' }
        Write-Host "-- creating .venv with $pyBase"
        & cmd /c "$pyBase -m venv `"$here\.venv`""
    }
}

# 3. libraries, at the exact versions listed in requirements.txt (the ones the server is tested with)
Write-Host '-- installing mcp and pywin32 (versions pinned in requirements.txt)'
if ($uv) {
    & uv pip install --python $py -r "$here\requirements.txt"
} else {
    & $py -m pip install -r "$here\requirements.txt"
}
if ($LASTEXITCODE -ne 0) { throw 'Library installation failed: see the error above.' }

# 4. COM connection test (Outlook must be open)
Write-Host '-- COM test with Outlook'
$test = @'
import win32com.client, sys
try:
    app = win32com.client.GetActiveObject("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    print("OK Outlook", app.Version, "profile", ns.CurrentProfileName, "stores:", ns.Stores.Count)
    for s in ns.Stores:
        print("   -", s.DisplayName)
except Exception as e:
    print("ERROR:", e); sys.exit(1)
'@
$test | & $py -
if ($LASTEXITCODE -ne 0) { throw 'COM connection failed: is classic Outlook open, as a normal user?' }
Write-Host '   The names above are the ones to use in config.json under "mailboxes".'

# 5. server start test (tool list)
Write-Host '-- MCP server start test'
$probe = @'
import asyncio, sys
sys.path.insert(0, r"__HERE__")
import server
tools = asyncio.run(server.mcp.list_tools())
print("OK server, phase", server.PHASE, "-", len(tools), "tools:", ", ".join(t.name for t in tools))
'@
$probe.Replace('__HERE__', $here) | & $py -
if ($LASTEXITCODE -ne 0) { throw 'The server does not start: see the error above (usually config.json).' }

# 6. registration in Claude desktop
# Claude desktop's MCP server configuration file is claude_desktop_config.json. It usually lives in %APPDATA%\Claude;
# some installations use their own profile under %LOCALAPPDATA%\Claude-Data\<profile>\.
# To find out which one yours uses: Claude desktop > Settings > Developer > Edit configuration.
$candidates = @()
$std = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
if (Test-Path $std) { $candidates += $std }
$profiles = Join-Path $env:LOCALAPPDATA 'Claude-Data'
if (Test-Path $profiles) {
    Get-ChildItem $profiles -Directory | ForEach-Object {
        $p = Join-Path $_.FullName 'claude_desktop_config.json'
        if (Test-Path $p) { $candidates += $p }
    }
}
if ($candidates.Count -eq 0) {
    $cfg = $std
    New-Item -ItemType Directory -Path (Split-Path $cfg) -Force | Out-Null
    Write-Host "-- no existing configuration: creating $cfg"
} elseif ($candidates.Count -eq 1) {
    $cfg = $candidates[0]
} else {
    Write-Host 'More than one Claude desktop configuration file found:'
    for ($i = 0; $i -lt $candidates.Count; $i++) { Write-Host "  [$($i+1)] $($candidates[$i])" }
    $choice = Read-Host 'Which one? (number; check in Claude desktop > Settings > Developer)'
    $n = 0
    if (-not [int]::TryParse($choice, [ref]$n) -or $n -lt 1 -or $n -gt $candidates.Count) {
        throw "Invalid choice '$choice': run the script again and type a number between 1 and $($candidates.Count)."
    }
    $cfg = $candidates[$n - 1]
}

$json = $null
if (Test-Path $cfg) {
    Copy-Item $cfg "$cfg.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"   # safety copy
    # -Encoding UTF8: Windows PowerShell 5.1 would otherwise read the file as ANSI and corrupt accented paths of the
    # OTHER servers configured there when writing it back.
    $raw = Get-Content $cfg -Raw -Encoding UTF8
    if ($raw -and $raw.Trim()) { $json = $raw | ConvertFrom-Json }
}
if ($null -eq $json) { $json = New-Object PSObject }
if (-not ($json.PSObject.Properties.Name -contains 'mcpServers') -or $null -eq $json.mcpServers) {
    $json | Add-Member -NotePropertyName mcpServers -NotePropertyValue (New-Object PSObject) -Force
}
# The "outlook" entry: command = the Python of the isolated environment, argument = server.py. The environment
# variables force UTF-8 on the standard streams, so accented letters pass intact between server and Claude desktop.
$entry = [PSCustomObject]@{
    command = $py
    args    = [string[]]@("$here\server.py")
    env     = [PSCustomObject]@{ PYTHONIOENCODING = 'utf-8'; PYTHONUTF8 = '1' }
}
if ($json.mcpServers.PSObject.Properties.Name -contains 'outlook') {
    $old = $json.mcpServers.outlook
    Write-Host "NOTE: an 'outlook' server is already registered ($($old.command) $($old.args -join ' '))." -ForegroundColor Yellow
    $ans = Read-Host 'Replace it with this installation? (y/n)'
    if ($ans -notmatch '^(y|yes|s|si)$') { Write-Host 'Nothing changed in the Claude desktop configuration.'; exit 0 }
}
$json.mcpServers | Add-Member -NotePropertyName outlook -NotePropertyValue $entry -Force
$text = $json | ConvertTo-Json -Depth 10
# UTF-8 WITHOUT BOM: Claude desktop does not read a JSON file with a BOM.
[IO.File]::WriteAllText($cfg, $text, (New-Object System.Text.UTF8Encoding $false))
Write-Host "-- registered in $cfg"

Write-Host ''
Write-Host 'DONE. Now quit Claude desktop completely (tray icon > Quit) and open it again.' -ForegroundColor Green
Write-Host 'The "outlook" server will appear among the local connectors. To check, ask Claude: "check the Outlook status".'
Write-Host 'Recommended full test of the writing tools: .\.venv\Scripts\python.exe .\test_accents.py'
