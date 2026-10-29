# Inventory installed software from registry and winget (lab machine only).
param()
$ErrorActionPreference = "Continue"
$outDir = Join-Path $PSScriptRoot "evidence"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$apps = @()
$paths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
foreach ($path in $paths) {
    Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.DisplayName) {
            $apps += [pscustomobject]@{
                Name            = $_.DisplayName
                Version         = $_.DisplayVersion
                Publisher       = $_.Publisher
                InstallDate     = $_.InstallDate
                Source          = "Registry"
            }
        }
    }
}

$winget = @()
try {
    $raw = winget list --accept-source-agreements 2>$null | Out-String
    $winget += [pscustomobject]@{ Note = "winget list captured"; Length = $raw.Length }
} catch {
    $winget += [pscustomobject]@{ Note = "winget not available"; Error = $_.Exception.Message }
}

$payload = [pscustomobject]@{
    TimestampUtc     = (Get-Date).ToUniversalTime().ToString("o")
    Host             = $env:COMPUTERNAME
    RegistryAppCount = $apps.Count
    SampleApps       = @($apps | Select-Object -First 40)
    Winget           = $winget
}
$file = Join-Path $outDir ("{0}_SoftwareInventory.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$payload | ConvertTo-Json -Depth 6 | Set-Content -Path $file -Encoding UTF8
Write-Host "Wrote $file ($($apps.Count) registry apps)"
$apps | Select-Object -First 15 | Format-Table -AutoSize
