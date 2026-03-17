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

$html = Get-Content -Raw -Path $indexPath
$css = Get-Content -Raw -Path $cssPath
$app = Get-Content -Raw -Path $appPath
$data = Get-Content -Raw -Path $dataPath
$logoB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($logoPath))

$html = $html -replace '<link rel="stylesheet" href="\./styles\.css">', ("<style>`r`n" + $css + "`r`n</style>")
$html = $html -replace '<img src="\./assets/logo\.jpg"', ('<img src="data:image/jpeg;base64,' + $logoB64 + '"')
$html = $html -replace '<script src="\./data/leads\.js"></script>\s*<script src="\./app\.js"></script>', ("<script>`r`n" + $data + "`r`n" + $app + "`r`n</script>")

$outPath = Join-Path $dist "Prospecting_Dashboard_single_file.html"
Set-Content -Path $outPath -Value $html -Encoding UTF8
Write-Host "Built:" $outPath
