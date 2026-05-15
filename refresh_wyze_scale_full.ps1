$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($env:WYZE_OUTPUT_DIR) {
    $PrivateBase = $env:WYZE_OUTPUT_DIR
} else {
    $PrivateBase = Join-Path $ProjectDir "output"
}

$RawDir = Join-Path $ProjectDir "raw"
$CleanDir = Join-Path $PrivateBase "clean"
$AnalysisDir = Join-Path $PrivateBase "analysis"
$GarminDir = Join-Path $PrivateBase "garmin"

$ExportScript = Join-Path $ProjectDir "export_wyze_scale_all.py"
$CleanScript = Join-Path $ProjectDir "scripts\wyze_clean_scale.py"
$AnalysisScript = Join-Path $ProjectDir "scripts\wyze_build_analysis.py"
$GarminScript = Join-Path $ProjectDir "scripts\wyze_export_garmin_fitbit_body_full_strict.py"

$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"

if (Test-Path $VenvPython) {
    $PythonExe = $VenvPython
} else {
    $PythonExe = "python"
}

Set-Location $ProjectDir

Write-Host ""
Write-Host "Using Python:" $PythonExe -ForegroundColor Cyan

if (!(Test-Path $ExportScript)) {
    Write-Host "Missing export script: $ExportScript" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $CleanScript)) {
    Write-Host "Missing cleaning script: $CleanScript" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $AnalysisScript)) {
    Write-Host "Missing analysis script: $AnalysisScript" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $GarminScript)) {
    Write-Host "Missing Garmin export script: $GarminScript" -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $RawDir -Force | Out-Null
New-Item -ItemType Directory -Path $CleanDir -Force | Out-Null
New-Item -ItemType Directory -Path $AnalysisDir -Force | Out-Null
New-Item -ItemType Directory -Path $GarminDir -Force | Out-Null

Write-Host ""
Write-Host "Step 1 of 5 - Running Wyze scale export..." -ForegroundColor Cyan
& $PythonExe $ExportScript

if ($LASTEXITCODE -ne 0) {
    Write-Host "Wyze export failed. Stopping." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Step 2 of 5 - Moving raw export files into raw folder..." -ForegroundColor Cyan

$RootJson = Join-Path $ProjectDir "Wyze_Scale_Raw_All.json"
$RootXlsx = Join-Path $ProjectDir "Wyze_Scale_Raw_All.xlsx"

$RawJson = Join-Path $RawDir "Wyze_Scale_Raw_All.json"
$RawXlsx = Join-Path $RawDir "Wyze_Scale_Raw_All.xlsx"

if (Test-Path $RootJson) {
    Move-Item $RootJson $RawJson -Force
    Write-Host "Moved JSON raw export to: $RawJson"
} else {
    Write-Host "No new root JSON export found." -ForegroundColor Yellow
}

if (Test-Path $RootXlsx) {
    Move-Item $RootXlsx $RawXlsx -Force
    Write-Host "Moved Excel raw export to: $RawXlsx"
} else {
    Write-Host "No new root Excel export found." -ForegroundColor Yellow
}

if (!(Test-Path $RawJson)) {
    Write-Host "Missing required raw JSON file: $RawJson" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Step 3 of 5 - Running cleaning script..." -ForegroundColor Cyan
& $PythonExe $CleanScript

if ($LASTEXITCODE -ne 0) {
    Write-Host "Cleaning script failed. Stopping." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Step 4 of 5 - Building trend analysis workbook..." -ForegroundColor Cyan
& $PythonExe $AnalysisScript

if ($LASTEXITCODE -ne 0) {
    Write-Host "Analysis script failed. Stopping." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Step 5 of 5 - Building Garmin Fitbit Body CSV..." -ForegroundColor Cyan
& $PythonExe $GarminScript

if ($LASTEXITCODE -ne 0) {
    Write-Host "Garmin CSV export failed. Stopping." -ForegroundColor Red
    exit $LASTEXITCODE
}

$LatestGarminCsv = Get-ChildItem -Path $GarminDir -Filter "Wyze_To_Garmin_FITBIT_BODY_*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $LatestGarminCsv) {
    Write-Host "No Garmin timestamped CSV found in: $GarminDir" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Wyze scale full refresh complete." -ForegroundColor Green
Write-Host ""
Write-Host "Raw JSON:"
Write-Host $RawJson
Write-Host ""
Write-Host "Clean workbook:"
Write-Host (Join-Path $CleanDir "Wyze_Scale_Clean_Master.xlsx")
Write-Host ""
Write-Host "Analysis workbook:"
Write-Host (Join-Path $AnalysisDir "Wyze_Scale_Trend_Analysis.xlsx")
Write-Host ""
Write-Host "Latest Garmin import CSV:"
Write-Host $LatestGarminCsv.FullName -ForegroundColor Green
Write-Host ""
Write-Host "Garmin manual import settings:"
Write-Host "Language: English"
Write-Host "Length Units: Feet, Yards, Miles"
Write-Host "Weight Units: Pounds"
Write-Host "Date Format: 31-12-2026"
Write-Host "Number Format: 1,234.56"
