# Summarise evidence for Essential Eight-style patch reporting (lab model).
param()
$outDir = Join-Path $PSScriptRoot "evidence"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$inv = Get-ChildItem $outDir -Filter "*_SoftwareInventory.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$out = Get-ChildItem $outDir -Filter "*_OutdatedPackages.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

$invCount = 0
if ($inv) {
    $data = Get-Content $inv.FullName -Raw | ConvertFrom-Json
    $invCount = [int]$data.RegistryAppCount
}

$hasOutdated = $false
$outdatedNote = "No outdated scan yet"
if ($out) {
    $odata = Get-Content $out.FullName -Raw | ConvertFrom-Json
    $outdatedNote = if ($odata.Output -match "No installed package") { "No upgrades listed" } else { "Upgrade candidates listed — see JSON" }
    $hasOutdated = $odata.Output -notmatch "No installed package"
}

# Essential Eight patch applications / OS — model maturity statement for the lab host only
$report = [pscustomobject]@{
    TimestampUtc              = (Get-Date).ToUniversalTime().ToString("o")
    Host                      = $env:COMPUTERNAME
    InventoriedApps           = $invCount
    OutdatedScan              = $outdatedNote
    E8_PatchApps_LabClaim     = "Pilot uses winget for third-party apps on one lab host. Not a fleet maturity attestation."
    E8_PatchOS_LabClaim       = "OS updates remain on Windows Update / managed channel. This pilot does not replace WSUS/Intune."
    RecommendedCycle          = "Critical/exploited: target 48 hours after test. Routine third-party: weekly winget ring. Monthly OS quality: after pilot ring."
    RollbackNote              = "Keep a restore point or previous package version before -Apply on anything business-critical."
}
$file = Join-Path $outDir ("{0}_EssentialEight_PatchReport.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$report | ConvertTo-Json -Depth 4 | Set-Content -Path $file -Encoding UTF8
$report | Format-List
Write-Host "Wrote $file"
