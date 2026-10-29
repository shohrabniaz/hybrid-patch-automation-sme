# List packages winget can upgrade (lab machine only).
param()
$ErrorActionPreference = "Continue"
$outDir = Join-Path $PSScriptRoot "evidence"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$upgradeText = ""
try {
    $upgradeText = winget upgrade --include-unknown --accept-source-agreements 2>&1 | Out-String
} catch {
    $upgradeText = "winget upgrade failed: $($_.Exception.Message)"
}

$payload = [pscustomobject]@{
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
    Host         = $env:COMPUTERNAME
    Tool         = "winget upgrade --include-unknown"
    Output       = $upgradeText
}
$file = Join-Path $outDir ("{0}_OutdatedPackages.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$payload | ConvertTo-Json -Depth 4 | Set-Content -Path $file -Encoding UTF8
Write-Host "Wrote $file"
Write-Host $upgradeText
