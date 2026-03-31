$ErrorActionPreference = "Stop"

$existing = @(
    "C:\Program Files\Tesseract-OCR\tesseract.exe",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe"
)

foreach ($candidate in $existing) {
    if (Test-Path $candidate) {
        exit 0
    }
}

$winget = Get-Command winget -ErrorAction SilentlyContinue
if (-not $winget) {
    exit 0
}

& $winget.Source install -e --id tesseract-ocr.tesseract --accept-package-agreements --accept-source-agreements
exit $LASTEXITCODE
