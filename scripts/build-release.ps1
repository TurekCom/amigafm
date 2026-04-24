param(
    [string]$Version = "0.1.0",
    [string]$Configuration = "release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        $iscc = $command.Source
    }
}
if (-not $iscc) {
    throw "Nie znaleziono ISCC.exe. Zainstaluj Inno Setup 6."
}

Push-Location $repoRoot
try {
    cargo build --release
    & $iscc "/DMyAppVersion=$Version" "installer\amigafm.iss"
    $installer = Join-Path $repoRoot "installer\output\AmigaFM-Setup-$Version.exe"
    if (-not (Test-Path $installer)) {
        throw "Nie utworzono instalatora: $installer"
    }
    Write-Host "Utworzono instalator: $installer"
}
finally {
    Pop-Location
}
