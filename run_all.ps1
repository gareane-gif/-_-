<#
One-shot runner for Windows PowerShell.
Usage: open PowerShell in the project folder and run:
  .\run_all.ps1

It will:
- create a virtualenv `.venv` if missing
- install `requirements.txt` into the venv (or system python if venv creation fails)
- run `run_local_check.py` and `find_student.py 232029`
- save logs: `pip_install_log.txt`, `run_check_output.txt`, `find_output.txt`
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host "Running project diagnostics..."
$cwd = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Push-Location $cwd

function Run-CommandToFile($exe, $args, $outfile) {
    Write-Host "Running: $exe $args"
    try {
        & $exe $args 2>&1 | Tee-Object -FilePath $outfile
    } catch {
        "ERROR running $exe $args : $_" | Tee-Object -FilePath $outfile
    }
}

# 1) Try to create venv
if (-not (Test-Path -Path .\.venv)) {
    Write-Host "Creating virtual environment .venv..."
    try {
        python -m venv .venv 2>&1 | Tee-Object -FilePath pip_install_log.txt
    } catch {
        Write-Host "Failed to create venv; will fallback to system python. Error: $_"
    }
} else {
    Write-Host ".venv already exists."
}

# Determine python and pip executables
$venvPython = Join-Path $cwd ".venv\Scripts\python.exe"
$venvPip = Join-Path $cwd ".venv\Scripts\pip.exe"
$useVenv = $false
if (Test-Path $venvPython) { $useVenv = $true }

if ($useVenv) {
    Write-Host "Using venv python: $venvPython"
    Run-CommandToFile $venvPython "-m pip install -r requirements.txt" "pip_install_log.txt"
    Run-CommandToFile $venvPython "run_local_check.py" "run_check_output.txt"
    Run-CommandToFile $venvPython "find_student.py 232029" "find_output.txt"
} else {
    Write-Host "Using system python"
    Run-CommandToFile python "-m pip install -r requirements.txt" "pip_install_log.txt"
    Run-CommandToFile python "run_local_check.py" "run_check_output.txt"
    Run-CommandToFile python "find_student.py 232029" "find_output.txt"
}

Write-Host "Done. Logs:
 - pip_install_log.txt
 - run_check_output.txt
 - find_output.txt"

Pop-Location
