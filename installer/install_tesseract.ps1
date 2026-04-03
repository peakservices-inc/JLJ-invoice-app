$ErrorActionPreference = "Stop"

$logPath = Join-Path $env:TEMP "JLJInvoiceStudio_TesseractInstall.log"
$preferredInstallDir = Join-Path ${env:ProgramFiles} "Tesseract-OCR"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] $Message"
}

function Find-Tesseract {
    $existing = @(
        "C:\Program Files\Tesseract-OCR\tesseract.exe",
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe"
    )

    foreach ($candidate in $existing) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $resolved = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($resolved -and $resolved.Path) {
        return $resolved.Path
    }

    return $null
}

function Install-BundledTesseract {
    param([string]$InstallerPath)

    if (-not (Test-Path $InstallerPath)) {
        return $false
    }

    Write-Log "Running bundled Tesseract installer: $InstallerPath"
    Write-Log "Preferred install directory: $preferredInstallDir"

    $arguments = @(
        "/S",
        "/D=$preferredInstallDir"
    )

    $process = Start-Process -FilePath $InstallerPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    Write-Log "Bundled installer exit code: $($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        Write-Log "Bundled installer returned a non-zero exit code."
        return $false
    }

    return [bool](Find-Tesseract)
}

Write-Log "Starting Tesseract dependency check."
$installedPath = Find-Tesseract
if ($installedPath) {
    Write-Log "Tesseract already present at $installedPath"
    exit 0
}

$bundledInstaller = Join-Path $PSScriptRoot "tesseract-ocr-installer.exe"
if (Install-BundledTesseract -InstallerPath $bundledInstaller) {
    $installedPath = Find-Tesseract
    Write-Log "Tesseract installed successfully at $installedPath"
    exit 0
}

$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget) {
    $ids = @(
        "UB-Mannheim.TesseractOCR",
        "tesseract-ocr.tesseract"
    )

    foreach ($packageId in $ids) {
        Write-Log "Trying winget package $packageId"

        try {
            & $winget.Path install --source winget --exact --id $packageId --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
            Write-Log "winget machine-scope exit code for $packageId : $LASTEXITCODE"
        }
        catch {
            Write-Log "winget machine-scope failed for $packageId : $($_.Exception.Message)"
        }

        $installedPath = Find-Tesseract
        if ($installedPath) {
            Write-Log "Tesseract installed successfully at $installedPath"
            exit 0
        }

        try {
            & $winget.Path install --source winget --exact --id $packageId --scope user --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
            Write-Log "winget user-scope exit code for $packageId : $LASTEXITCODE"
        }
        catch {
            Write-Log "winget user-scope failed for $packageId : $($_.Exception.Message)"
        }

        $installedPath = Find-Tesseract
        if ($installedPath) {
            Write-Log "Tesseract installed successfully at $installedPath"
            exit 0
        }
    }
}

Write-Log "Tesseract installation failed."
throw "Tesseract OCR could not be installed automatically. See log: $logPath"
