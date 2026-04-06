param(
    [switch]$BuildInstaller
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314\python.exe"
$PyInstaller = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314\Scripts\pyinstaller.exe"
$Iscc = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
$AppName = "JLJ-invoice Rider"
$DistAppDir = Join-Path (Join-Path $Root "dist") $AppName
$InstallerScript = Join-Path $Root "installer\JLJ-invoice Rider.iss"
$InstallerOutputDir = Join-Path $Root "installer_output"
$InstallerOutput = Join-Path $InstallerOutputDir "$AppName Setup.exe"
$LegacyInstallerOutput = Join-Path $InstallerOutputDir "JLJInvoiceStudioSetup.exe"
$Icon = Join-Path $Root "assets\jlj_invoice.ico"
$TesseractBundleDir = Join-Path $Root "installer\downloads"
$TesseractBundle = Join-Path $TesseractBundleDir "tesseract-ocr-installer.exe"
$FallbackTesseractUrl = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.4.0.20240606.exe"

if (-not (Test-Path $Python)) {
    throw "Python 3.14 was not found at $Python"
}

function Get-TesseractInstallerUrl {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        try {
            $output = & $winget.Path show --source winget --accept-source-agreements UB-Mannheim.TesseractOCR 2>$null | Out-String
            $match = [regex]::Match($output, "Installer Url:\s+(https?://\S+)")
            if ($match.Success) {
                return $match.Groups[1].Value.Trim()
            }
        }
        catch {
        }
    }

    return $FallbackTesseractUrl
}

function Ensure-TesseractBundle {
    New-Item -ItemType Directory -Force -Path $TesseractBundleDir | Out-Null
    $url = Get-TesseractInstallerUrl
    Write-Host "Downloading bundled Tesseract installer from $url"
    & curl.exe -L -A "Mozilla/5.0" $url -o $TesseractBundle
    if (-not (Test-Path $TesseractBundle)) {
        throw "Failed to download bundled Tesseract installer."
    }
}

Push-Location $Root
try {
    if (-not (Test-Path $PyInstaller)) {
        & $Python -m pip install -r .\requirements-desktop.txt
    }

    Remove-Item -Recurse -Force .\build, .\dist -ErrorAction SilentlyContinue

    $PyArgs = @(
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onedir",
        "--name", $AppName,
        "--collect-all", "fitz",
        "--collect-all", "pytesseract",
        "--collect-all", "PIL"
    )

    if (Test-Path $Icon) {
        $PyArgs += @("--icon", $Icon)
    }

    $PyArgs += ".\jlj_invoice_desktop.py"

    & $PyInstaller @PyArgs

    Copy-Item .\installer\install_tesseract.ps1 (Join-Path $DistAppDir "install_tesseract.ps1") -Force

    if ($BuildInstaller) {
        if (-not (Test-Path $Iscc)) {
            throw "Inno Setup compiler not found at $Iscc"
        }
        Ensure-TesseractBundle
        Remove-Item -LiteralPath $LegacyInstallerOutput, $InstallerOutput -Force -ErrorAction SilentlyContinue
        & $Iscc $InstallerScript
    }
}
finally {
    Pop-Location
}
