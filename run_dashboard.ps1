$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$venvActivate = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $venvActivate) {
    . $venvActivate
}

python (Join-Path $PSScriptRoot "dashboard_server.py") --port 8111
exit $LASTEXITCODE

