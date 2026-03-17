$base = "C:\Users\LucasMartinsRocha\OneDrive - Apeiron Pte Ltd\Prospecting"
$src = Join-Path $base "Apeiron_BR_Gestao_Comercial.xlsx"
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$bak = Join-Path $base ("Apeiron_BR_Gestao_Comercial_backup_{0}.xlsx" -f $ts)

Copy-Item -LiteralPath $src -Destination $bak -Force

$all = Get-ChildItem -LiteralPath $base -Filter "Apeiron_BR_Gestao_Comercial_backup_*.xlsx" | Sort-Object Name -Descending
$toDelete = $all | Select-Object -Skip 2
foreach ($f in $toDelete) {
  Remove-Item -LiteralPath $f.FullName -Force
  Write-Output ("Deleted old backup: " + $f.Name)
}

Write-Output ("Created backup: " + [IO.Path]::GetFileName($bak))
Get-ChildItem -LiteralPath $base -Filter "Apeiron_BR_Gestao_Comercial_backup_*.xlsx" |
  Sort-Object LastWriteTime -Descending |
  Select-Object Name, LastWriteTime |
  Format-Table -AutoSize
