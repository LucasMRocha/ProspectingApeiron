$base = Split-Path -Parent $PSScriptRoot
if (!(Test-Path -LiteralPath (Join-Path $base "Apeiron_BR_Gestao_Comercial.xlsx"))) {
  $base = Split-Path -Parent $base
}
$support = Join-Path $base "Support Files"
$bakDir = Join-Path $support "_backup"
New-Item -ItemType Directory -Force -Path $bakDir | Out-Null

$src = Join-Path $base "Apeiron_BR_Gestao_Comercial.xlsx"
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$bak = Join-Path $bakDir ("Apeiron_BR_Gestao_Comercial_backup_{0}.xlsx" -f $ts)

Copy-Item -LiteralPath $src -Destination $bak -Force

$all = Get-ChildItem -LiteralPath $bakDir -Filter "Apeiron_BR_Gestao_Comercial_backup_*.xlsx" | Sort-Object Name -Descending
$toDelete = $all | Select-Object -Skip 2
foreach ($f in $toDelete) {
  Remove-Item -LiteralPath $f.FullName -Force
  Write-Output ("Deleted old backup: " + $f.Name)
}

Write-Output ("Created backup: " + [IO.Path]::GetFileName($bak))
Get-ChildItem -LiteralPath $bakDir -Filter "Apeiron_BR_Gestao_Comercial_backup_*.xlsx" |
  Sort-Object LastWriteTime -Descending |
  Select-Object Name, LastWriteTime |
  Format-Table -AutoSize
