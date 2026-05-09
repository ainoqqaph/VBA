Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
Set-Location $RepoRoot

$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    Write-Host "Local virtual environment was not found. Running setup first..."
    & (Join-Path $ScriptDir "setup_windows.ps1")
}

$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
Write-Host "Starting Streamlit at http://localhost:8501"
& $VenvPython -m streamlit run app.py --server.address localhost --server.port 8501
