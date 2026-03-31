param(
    [switch]$BuildInstaller
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314\python.exe"
$PyInstaller = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314\Scripts\pyinstaller.exe"
$Iscc = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
$Icon = Join-Path $Root "assets\jlj_invoice.ico"

if (-not (Test-Path $Python)) {
    throw "Python 3.14 was not found at $Python"
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
        "--name", "JLJInvoiceStudio",
        "--collect-all", "fitz",
        "--collect-all", "pytesseract",
        "--collect-all", "PIL"
    )

    if (Test-Path $Icon) {
        $PyArgs += @("--icon", $Icon)
    }

    $PyArgs += ".\jlj_invoice_desktop.py"

    & $PyInstaller @PyArgs

    Copy-Item .\installer\install_tesseract.ps1 .\dist\JLJInvoiceStudio\install_tesseract.ps1 -Force

    if ($BuildInstaller) {
        if (-not (Test-Path $Iscc)) {
            throw "Inno Setup compiler not found at $Iscc"
        }
        & $Iscc ".\installer\JLJInvoiceStudio.iss"
    }
}
finally {
    Pop-Location
}
