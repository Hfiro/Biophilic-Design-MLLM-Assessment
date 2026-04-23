# Author: Hoda Alem
# Date: November 2025
# Context: Part of PhD research for a PhD in Environmental Design and Planning
#          at Myers-Lawson School of Construction, Virginia Tech, Blacksburg, VA.

$ErrorActionPreference = 'Stop'

param(
  [string]$Python = 'python',
  [string]$Excel = 'BD GPT Prompts.xlsx',
  [switch]$SkipPlots,
  [switch]$NoMplAgg,
  [switch]$NoRun
)

Write-Host '==> Checking Python...'
try {
  & $Python --version | Out-Host
} catch {
  Write-Error "Python is not available on PATH. Please install Python 3.10+ and rerun."
  exit 1
}

Write-Host '==> Ensuring virtual environment (.venv)'
if (-not (Test-Path .venv\Scripts\Activate.ps1)) {
  & $Python -m venv .venv
}

Write-Host '==> Activating virtual environment'
. .\.venv\Scripts\Activate.ps1

Write-Host '==> Upgrading pip'
python -m pip install --upgrade pip

Write-Host '==> Installing Python dependencies (requirements.txt)'
pip install -r requirements.txt

# Optional: configure headless plotting backend and writable MPL config dir
if (-not $NoMplAgg.IsPresent) {
  if (-not (Test-Path '.cache-mpl')) { New-Item -ItemType Directory '.cache-mpl' | Out-Null }
  $env:MPLBACKEND = 'Agg'
  $env:MPLCONFIGDIR = (Resolve-Path '.cache-mpl').Path
}

# Prompt for API key if not set
if (-not $env:OPENAI_API_KEY) {
  $api = Read-Host 'Enter OPENAI_API_KEY (sk-...) or press Enter to skip for now'
  if ($api) { $env:OPENAI_API_KEY = $api }
}

if (-not $NoRun.IsPresent) {
  Write-Host '==> Running evaluator (biophilic_eval.py)'
  python biophilic_eval.py --xlsx $Excel

  if ($LASTEXITCODE -ne 0) {
    Write-Warning 'Evaluator exited with a non-zero code. You can rerun: `python biophilic_eval.py --xlsx "BD GPT Prompts.xlsx"`'
  }
}

if ((-not $SkipPlots.IsPresent) -and (-not $NoRun.IsPresent)) {
  Write-Host '==> Consolidating results for plotting/reporting (combined_summary.xlsx)'
  python results_analyzer.py --consolidated-xlsx combined_summary.xlsx --xlsx $Excel --glob 'results_*.xlsx'

  if ($LASTEXITCODE -eq 0) {
    Write-Host '==> Generating plots'
    python plot_summary.py --xlsx combined_summary.xlsx --grid --out grid.png
    python plot_accuracy_progression.py --glob 'results_*.xlsx' --xlsx $Excel --out accuracy_progression.png
    python plot_dispersion_progression.py --glob 'results_*.xlsx' --out dispersion_progression.png

    Write-Host '==> Building final report table (last iteration)'
    python build_report_table.py --summary combined_summary.xlsx --expert $Excel --out final_report_table.xlsx
  }
}

Write-Host '==> Done.'
