$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$venvActivate = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $venvActivate) {
    . $venvActivate
}

python -u (Join-Path $PSScriptRoot "sync_dingtalk_to_feishu.py") @args
exit $LASTEXITCODE
