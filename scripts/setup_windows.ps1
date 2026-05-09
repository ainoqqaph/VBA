Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
Set-Location $RepoRoot

function Resolve-PythonExe {
    $candidates = @("py", "python", "python3")

    foreach ($candidate in $candidates) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if (-not $command) {
            continue
        }

        if ($candidate -eq "py") {
            $pythonExe = & py -3 -c "import sys; print(sys.executable)" 2>$null
        } else {
            $pythonExe = & $candidate -c "import sys; print(sys.executable)" 2>$null
        }

        if ($LASTEXITCODE -eq 0 -and $pythonExe) {
            return ($pythonExe | Select-Object -Last 1).Trim()
        }
    }

    throw "Python 3 was not found. Install Python 3.10 or newer, then run this script again."
}

$PythonExe = Resolve-PythonExe
Write-Host "Using Python: $PythonExe"

& $PythonExe -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
if ($LASTEXITCODE -ne 0) {
    throw "Python 3.10 or newer is required."
}

if (-not (Test-Path ".venv")) {
    Write-Host "Creating local virtual environment..."
    & $PythonExe -m venv .venv
}

$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "Virtual environment Python was not created correctly."
}

Write-Host "Upgrading pip..."
& $VenvPython -m pip install --upgrade pip

Write-Host "Installing project dependencies..."
& $VenvPython -m pip install -r requirements.txt

Write-Host "Checking imports..."
& $VenvPython -c "import streamlit, pandas, openpyxl, xlrd, pdfplumber, docx; print('Environment check passed')"

Write-Host ""
Write-Host "Setup complete."
Write-Host "Run the app with:"
Write-Host ".\scripts\run_streamlit.ps1"
