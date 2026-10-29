# Apply winget upgrades. Default is -WhatIf style (no install) unless -Apply is set.
param(
    [switch]$Apply
)
$ErrorActionPreference = "Continue"
$outDir = Join-Path $PSScriptRoot "evidence"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

if (-not $Apply) {
    $msg = "Dry run only. Re-run with -Apply to install upgrades on this lab machine."
    Write-Host $msg -ForegroundColor Yellow
    $payload = [pscustomobject]@{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
        Mode         = "DryRun"
        Message      = $msg
    }
} else {
    Write-Host "Applying winget upgrades on $env:COMPUTERNAME ..." -ForegroundColor Cyan
    $output = winget upgrade --all --include-unknown --accept-package-agreements --accept-source-agreements 2>&1 | Out-String
    $payload = [pscustomobject]@{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
        Mode         = "Apply"
        Host         = $env:COMPUTERNAME
        Output       = $output
    }
    Write-Host $output
}

$file = Join-Path $outDir ("{0}_UpgradeRun.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$payload | ConvertTo-Json -Depth 4 | Set-Content -Path $file -Encoding UTF8
Write-Host "Wrote $file"
