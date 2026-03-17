# get-zoho-token.ps1
# Exchange a Zoho grant code for a refresh token and save local env file.
# Usage:
#   .\tools\get-zoho-token.ps1 -Code "1000.xxxxx.xxxxx"
# Optional:
#   -ClientId "...", -ClientSecret "...", -AccountsDomain "accounts.zoho.com"

param(
    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = $env:ZOHO_CLIENT_ID,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = $env:ZOHO_CLIENT_SECRET,

    [Parameter(Mandatory = $false)]
    [string]$AccountsDomain = "accounts.zoho.com"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw "Missing ClientId. Pass -ClientId or set `$env:ZOHO_CLIENT_ID before running."
}
if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
    throw "Missing ClientSecret. Pass -ClientSecret or set `$env:ZOHO_CLIENT_SECRET before running."
}

Write-Host "Exchanging grant code for refresh token..." -ForegroundColor Cyan

$body = @{
    grant_type    = "authorization_code"
    client_id     = $ClientId
    client_secret = $ClientSecret
    redirect_uri  = ""
    code          = $Code
}

$resp = Invoke-RestMethod -Uri "https://$AccountsDomain/oauth/v2/token" `
                          -Method POST `
                          -Body $body `
                          -ContentType "application/x-www-form-urlencoded"

if (-not $resp.refresh_token) {
    Write-Host "Zoho API error response:" -ForegroundColor Red
    $resp | ConvertTo-Json
    exit 1
}

Write-Host "Refresh token obtained successfully." -ForegroundColor Green

$envPath = Join-Path $PSScriptRoot "zoho-env.ps1"
@"
# Zoho CRM OAuth credentials - generated locally
`$env:ZOHO_CLIENT_ID     = "$ClientId"
`$env:ZOHO_CLIENT_SECRET = "$ClientSecret"
`$env:ZOHO_REFRESH_TOKEN = "$($resp.refresh_token)"
`$env:ZOHO_ACCOUNTS_DOMAIN = "$AccountsDomain"
"@ | Set-Content -Path $envPath -Encoding UTF8

Write-Host "Credentials saved to: $envPath" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  . .\tools\zoho-env.ps1" -ForegroundColor White
Write-Host "  python .\tools\sync-zoho-to-db.py --dry-run --skip-build" -ForegroundColor White
