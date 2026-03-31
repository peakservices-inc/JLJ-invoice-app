# JLJ Invoice Rule Studio

Property of JLJ IV Enterprises Inc.

JLJ Invoice Rule Studio is a Windows-based OCR workflow for scanned invoice PDFs. It can:

- detect an invoice date using OCR
- add a calculated due date based on the invoice date
- add a bottom note to each page
- let office users manage rules from a desktop app instead of editing code
- build a Windows installer for distribution inside the business

The project currently includes two ways to use the logic:

- `annotate_invoice_due_dates.py`: direct script usage
- `jlj_invoice_desktop.py`: desktop application with a configurable UI

## Project Files

- `annotate_invoice_due_dates.py`: OCR and PDF annotation engine
- `jlj_invoice_desktop.py`: PySide6 desktop app
- `build_app.ps1`: build script for the packaged app and installer
- `requirements-script.txt`: minimal dependencies for script-only usage
- `requirements-desktop.txt`: dependencies for the desktop app and packaging
- `installer/JLJInvoiceStudio.iss`: Inno Setup installer definition
- `installer/install_tesseract.ps1`: optional Tesseract bootstrap run by the installer

## Requirements

### Script-only requirements

- Windows
- Python 3.14 or another recent Python 3 version
- Tesseract OCR installed
- Python packages:
  - `pymupdf`
  - `pillow`
  - `pytesseract`

### Desktop app requirements

- Windows
- Python 3.14
- Tesseract OCR installed
- Python packages from `requirements-desktop.txt`
- For building the installer: Inno Setup 6

## Install Tesseract

The OCR pipeline depends on Tesseract.

Recommended install path:

```powershell
winget install -e --id tesseract-ocr.tesseract
```

Typical executable location after install:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If Tesseract is not on `PATH`, both the script and the desktop app let you provide the full executable path.

## Run The Script Only

Install the minimal dependencies:

```powershell
py -m pip install -r .\requirements-script.txt
```

Run against a single PDF:

```powershell
py .\annotate_invoice_due_dates.py `
  --input "C:\path\to\invoice.pdf" `
  --output "C:\path\to\invoice_processed.pdf" `
  --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

Run in batch mode using the default `INPUT` and `OUTPUT` folders:

```powershell
py .\annotate_invoice_due_dates.py `
  --input-dir .\INPUT `
  --output-dir .\OUTPUT `
  --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

If no `--input` is provided, the script scans `INPUT\*.pdf` and writes processed files to `OUTPUT`.

### Useful script flags

- `--input`: single input PDF
- `--output`: single output PDF
- `--input-dir`: folder to scan for PDFs
- `--output-dir`: folder for processed PDFs
- `--dpi`: OCR render DPI, default `300`
- `--jpeg-quality`: output image quality, default `95`
- `--tesseract-cmd`: explicit path to `tesseract.exe`

## Run The Desktop App From Source

Install the desktop dependencies:

```powershell
py -m pip install -r .\requirements-desktop.txt
```

Start the app:

```powershell
py .\jlj_invoice_desktop.py
```

### How the desktop app works

1. Add one or more PDF files.
2. Choose the output folder.
3. Confirm the Tesseract path.
4. Review or edit the rules.
5. Click `Process Files`.
6. Open the processed PDFs from the results pane.

### Current starter rules

- `Due Date`: finds the invoice date and adds a due date based on offset days
- `Bottom Note`: adds a configurable note near the bottom of the page

The UI is rule-driven, so new rule types can be added later without redesigning the main workflow.

### App data

The desktop app stores user settings and logs here:

```text
Settings: C:\Users\<YourUser>\AppData\Roaming\JLJ IV Enterprises Inc\Invoice Rule Studio\settings.json
Log file: C:\Users\<YourUser>\AppData\Roaming\JLJ IV Enterprises Inc\Invoice Rule Studio\app.log
```

## Build The Packaged Windows App

The project uses PyInstaller to create a one-folder Windows app build.

Install build dependencies:

```powershell
py -m pip install -r .\requirements-desktop.txt
```

Build the desktop app only:

```powershell
.\build_app.ps1
```

Build the desktop app and the installer:

```powershell
.\build_app.ps1 -BuildInstaller
```

### What `build_app.ps1` does

- checks for Python at `%LOCALAPPDATA%\Programs\Python\Python314\python.exe`
- installs packaging dependencies if needed
- removes old `build` and `dist` folders
- runs PyInstaller in `--onedir` mode
- copies `install_tesseract.ps1` into the packaged app
- optionally compiles the Inno Setup installer

### Build outputs

- packaged app folder: `dist\JLJInvoiceStudio\`
- installer exe: `installer_output\JLJInvoiceStudioSetup.exe`

## Installer Notes

The installer is defined in `installer/JLJInvoiceStudio.iss` and currently provides:

- modern Inno Setup wizard
- desktop shortcut option
- optional Tesseract installation task
- launch-after-install option

The installer copies the packaged app from `dist\JLJInvoiceStudio\` into the installation folder and can run `installer/install_tesseract.ps1` during setup.

## Rule Configuration Notes

The OCR engine currently supports:

- month-style top dates such as `March 19, 2026`
- labeled dates such as `Invoice Date`
- auto-detection mode

Users can configure rule appearance from the app, including:

- font family
- text size adjustments
- spacing
- alignment
- x/y offsets
- label text
- note text

Unsafe symbol fonts are automatically rejected for rule text so processed PDFs do not render boxes instead of letters.

## Troubleshooting

### Due date shows boxes instead of letters

This usually means a symbol-style font was selected. The app now sanitizes unsafe fonts automatically and falls back to a safe text font such as Arial.

### Processing fails

Check:

- Tesseract is installed
- the Tesseract path is correct
- the input PDF exists
- the app log at `AppData\Roaming\JLJ IV Enterprises Inc\Invoice Rule Studio\app.log`

### Installer build fails

Check:

- Python 3.14 is installed in `%LOCALAPPDATA%\Programs\Python\Python314`
- Inno Setup 6 is installed in `%LOCALAPPDATA%\Programs\Inno Setup 6`
- all Python build dependencies are installed

## Development Notes

- This repository is Windows-focused.
- Font detection currently relies on Windows font files under `C:\Windows\Fonts`.
- OCR accuracy depends on scan quality and Tesseract results.

