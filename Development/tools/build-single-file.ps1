param(
  [string]$Root = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"
$src = Join-Path $Root "src\dashboard"
$dist = Join-Path $Root "dist"
New-Item -ItemType Directory -Force -Path $dist | Out-Null

$indexPath = Join-Path $src "index.html"
$cssPath = Join-Path $src "styles.css"
$appPath = Join-Path $src "app.js"
$dataPath = Join-Path $src "data\leads.js"
$logoPath = Join-Path $src "assets\logo.jpg"

foreach ($p in @($indexPath, $cssPath, $appPath, $dataPath, $logoPath)) {
  if (!(Test-Path $p)) { throw "Missing required file: $p" }
}

$integrityScript = Join-Path $Root "tools\check_db_integrity.py"
if (Test-Path $integrityScript) {
  $pythonCmd = $null
  if (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
  } elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py -3"
  }
  if ($pythonCmd) {
    Write-Host "Running integrity check..."
    & cmd /c "$pythonCmd `"$integrityScript`""
    if ($LASTEXITCODE -ne 0) { throw "Integrity check failed. Fix data issues before build." }
  } else {
    Write-Warning "Python not found; skipping integrity check."
  }
}

$html = Get-Content -Raw -Path $indexPath -Encoding UTF8
$css = Get-Content -Raw -Path $cssPath -Encoding UTF8
$app = Get-Content -Raw -Path $appPath -Encoding UTF8
$data = Get-Content -Raw -Path $dataPath -Encoding UTF8
$logoB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($logoPath))

$html = $html -replace '<link rel="stylesheet" href="\./styles\.css">', ("<style>`r`n" + $css + "`r`n</style>")
$html = $html -replace '<img src="\./assets/logo\.jpg"', ('<img src="data:image/jpeg;base64,' + $logoB64 + '"')
$html = $html -replace '<script src="\./data/leads\.js"></script>\s*<script src="\./app\.js"></script>', ("<script>`r`n" + $data + "`r`n" + $app + "`r`n</script>")

$dashboardFile = "APEIRON_BRASIL_-_Opportunities_Management_single_file.html"
$outPath = Join-Path $dist $dashboardFile
if (Test-Path $outPath) {
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $backupPath = Join-Path $dist ("APEIRON_BRASIL_-_Opportunities_Management_single_file.backup.{0}.html" -f $ts)
  Copy-Item -Path $outPath -Destination $backupPath -Force

  $backups = Get-ChildItem -Path $dist -Filter "APEIRON_BRASIL_-_Opportunities_Management_single_file.backup.*.html" |
    Sort-Object Name -Descending
  $oldBackups = $backups | Select-Object -Skip 2
  foreach ($f in $oldBackups) {
    Remove-Item -Path $f.FullName -Force
  }
}

Set-Content -Path $outPath -Value $html -Encoding UTF8
Write-Host "Built:" $outPath
