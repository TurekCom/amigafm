param(
    [string]$Version = "0.1.1",
    [string]$Configuration = "release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$isccCandidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
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
    if ([string]::IsNullOrWhiteSpace($env:AMIGAFM_GOOGLE_DRIVE_CLIENT_ID)) {
        Write-Warning "Brak AMIGAFM_GOOGLE_DRIVE_CLIENT_ID. Google Drive nie zaloguje użytkownika w tym buildzie."
    }
    if ([string]::IsNullOrWhiteSpace($env:AMIGAFM_DROPBOX_APP_KEY)) {
        Write-Warning "Brak AMIGAFM_DROPBOX_APP_KEY. Dropbox nie zaloguje użytkownika w tym buildzie."
    }
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
