# run.ps1 - run the Computer Resource Monitor once from PowerShell.
#
# Usage:
#   .\run.ps1          real run, samples every 5 minutes for 24 hours
#   .\run.ps1 -Demo    quick demonstration, about 30 seconds
#
# Python 3 must be installed and on the PATH. The script checks and installs the
# Python packages it needs on first run.

param([switch]$Demo)

Set-Location -Path $PSScriptRoot

if ($Demo) {
    python monitor.py --demo
} else {
    python monitor.py
}
