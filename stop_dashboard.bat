@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
set "PORT=8111"
set "SYNC_ROOT=%ROOT%"

cd /d "%ROOT%"

echo Stopping DingTalk to Feishu sync dashboard on port %PORT%...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$port=%PORT%;" ^
  "$pids=@();" ^
  "$pids += Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique;" ^
  "$root=(Resolve-Path $env:SYNC_ROOT).Path;" ^
  "$escaped=[regex]::Escape($root);" ^
  "$projectProcs=Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -and $_.CommandLine -match $escaped -and ($_.CommandLine -match 'dashboard_server\.py' -or $_.CommandLine -match 'sync_dingtalk_to_feishu\.py') } | Select-Object -ExpandProperty ProcessId;" ^
  "$pids += $projectProcs;" ^
  "$pids | Where-Object { $_ } | Sort-Object -Unique | ForEach-Object { try { Stop-Process -Id $_ -Force -ErrorAction Stop; Write-Host ('Stopped PID ' + $_) } catch {} }"

echo Closing automation browser processes for this project profile...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$profile=(Join-Path (Resolve-Path $env:SYNC_ROOT).Path '.browser\dingtalk');" ^
  "$escaped=[regex]::Escape($profile);" ^
  "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -and $_.CommandLine -match $escaped -and $_.Name -in @('msedge.exe','chrome.exe','msedgewebview2.exe') } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop; Write-Host ('Stopped browser PID ' + $_.ProcessId) } catch {} }"

echo Cleaning safe cache and temporary files...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$root=(Resolve-Path $env:SYNC_ROOT).Path;" ^
  "$removed=0;" ^
  "function Remove-PathSafe($path) { if (-not $path) { return }; try { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop; $script:removed++ } } catch {} }" ^
  "Get-ChildItem -LiteralPath $root -Directory -Recurse -Force -Filter '__pycache__' -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\(\.venv|\.browser)(\\|$)' } | ForEach-Object { Remove-PathSafe $_.FullName };" ^
  "Get-ChildItem -LiteralPath $root -Directory -Recurse -Force -Filter '.pytest_cache' -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\(\.venv|\.browser)(\\|$)' } | ForEach-Object { Remove-PathSafe $_.FullName };" ^
  "Get-ChildItem -LiteralPath $root -File -Force -Filter 'debug_*.png' -ErrorAction SilentlyContinue | ForEach-Object { Remove-PathSafe $_.FullName };" ^
  "Get-ChildItem -LiteralPath $root -File -Force -Filter '.env.*.tmp' -ErrorAction SilentlyContinue | ForEach-Object { Remove-PathSafe $_.FullName };" ^
  "$downloads=Join-Path $root 'downloads'; if (Test-Path -LiteralPath $downloads) { Get-ChildItem -LiteralPath $downloads -File -Force -Filter '*.candidate*.xlsx' -ErrorAction SilentlyContinue | ForEach-Object { Remove-PathSafe $_.FullName }; Get-ChildItem -LiteralPath $downloads -File -Force -Filter '*_probe.*' -ErrorAction SilentlyContinue | ForEach-Object { Remove-PathSafe $_.FullName }; Get-ChildItem -LiteralPath $downloads -File -Force -Filter '*failure*.png' -ErrorAction SilentlyContinue | ForEach-Object { Remove-PathSafe $_.FullName } };" ^
  "$logs=Join-Path $root 'logs'; if (Test-Path -LiteralPath $logs) { Get-ChildItem -LiteralPath $logs -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Length -eq 0 } | ForEach-Object { Remove-PathSafe $_.FullName } };" ^
  "Write-Host ('Removed cache/temp item(s): ' + $removed)"

echo Done.
exit /b 0
