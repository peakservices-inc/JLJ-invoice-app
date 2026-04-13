# JLJ-invoice Rider

Property of JLJ IV Enterprises Inc.

JLJ-invoice Rider is a Windows-based OCR workflow for scanned invoice PDFs. It can:

- detect an invoice date using OCR
- add a calculated due date based on the invoice date
- add a bottom note to each page
- let office users manage rules from a desktop app instead of editing code
- build a Windows installer for distribution inside the business

The project currently includes two ways to use the logic:

- `annotate_invoice_due_dates.py`: direct script usage
- `jlj_invoice_desktop.py`: desktop application with a configurable UI

## Stack Overview

This project is a Windows desktop OCR application built on a small Python stack:

- `Python 3.14`: application runtime
- `PyMuPDF (fitz)`: opens PDFs, renders pages to images, and rebuilds the processed PDF
- `Pillow`: image drawing and text rendering
- `pytesseract`: Python bridge to Tesseract OCR
- `Tesseract OCR`: extracts text from scanned invoice images
- `PySide6`: desktop application UI
- `PyInstaller`: packages the Python app into a Windows app folder
- `Inno Setup`: builds the Windows installer

## How The Stack Works

At a high level, the project has four layers:

1. `OCR and annotation engine`
   `annotate_invoice_due_dates.py` handles PDF rendering, OCR, invoice date detection, due-date calculation, note placement, and saving the finished PDF.

2. `Rule configuration layer`
   The due-date and note rules are defined as dataclasses in `annotate_invoice_due_dates.py`. The desktop app edits these rule objects and passes them to the backend before processing.

3. `Desktop UI layer`
   `jlj_invoice_desktop.py` gives office users a simple front end. The main window focuses on:
   - adding files
   - choosing the output folder
   - clicking `Proceed`
   - viewing results

   Advanced options such as OCR settings and rule editing are moved into the Settings dialog.

4. `Packaging and installer layer`
   `build_app.ps1` runs PyInstaller to create the packaged app and optionally runs Inno Setup to create the installer.

### Processing flow

When a user clicks `Proceed`, the app does this:

1. validates the selected files, output folder, enabled rules, and Tesseract path
2. starts a background worker so the UI does not freeze
3. converts the selected rules into a backend processing config
4. renders each PDF page to an image with PyMuPDF
5. runs OCR on that image with Tesseract
6. finds the invoice date using the configured detection mode
7. calculates the due date
8. draws the due date and note onto the page image
9. rebuilds a new PDF from the annotated page images
10. saves the finished files and shows them in the Results panel

### Key code entry points

- Backend rule definitions: `annotate_invoice_due_dates.py`
- Date detection: `find_invoice_date_match`
- Due-date drawing: `draw_due_date_block`
- Note drawing: `draw_note_block`
- PDF pipeline: `process_pdf`
- Desktop worker thread: `ProcessWorker` in `jlj_invoice_desktop.py`
- Main app window: `MainWindow` in `jlj_invoice_desktop.py`
- Settings dialog UI: `_build_settings_dialog` in `jlj_invoice_desktop.py`

## Project Files

- `annotate_invoice_due_dates.py`: OCR and PDF annotation engine
- `jlj_invoice_desktop.py`: PySide6 desktop app
- `build_app.ps1`: build script for the packaged app and installer
- `requirements-script.txt`: minimal dependencies for script-only usage
- `requirements-desktop.txt`: dependencies for the desktop app and packaging
- `installer/JLJ-invoice Rider.iss`: Inno Setup installer definition
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
3. Click `Proceed`.
4. Watch processing messages in the `Processing Details` pane.
5. Open the processed PDFs from the results pane.

If needed, click `Settings` to:

- click `?` beside a setting to see what it does
- change the Tesseract path
- change OCR DPI and JPEG quality
- edit rules
- change font family, size, color, and bold styling
- change due-date spacing and placement
- change note spacing, alignment, and placement

### Current starter rules

- `Due Date`: finds the invoice date and adds a due date based on offset days
- `Bottom Note`: adds a configurable note near the bottom of the page

The UI is rule-driven, so new rule types can be added later without redesigning the main workflow.

### Current due-date defaults

The default due-date rule currently uses:

- `30` days after the invoice date
- default horizontal adjustment of `-30`
- default line spacing of `1 space` between `Due Date` and the date below it

Users can adjust these values from `Settings > Rules`.

### Current text and layout controls

The app now exposes a fuller set of text-related controls so office users can tune the output without editing code.

Shared styling options available across rules:

- font family
- text size adjustment, including very small text sizes
- bold on or off
- text color
- horizontal position
- vertical position

Due Date rule options:

- days after invoice date
- invoice-date detection mode
- custom invoice-date label keywords
- line spacing between the due-date label and the due-date value
- mirrored right-side spacing based on the source invoice date position

Bottom Note rule options:

- full note text
- left, center, or right alignment
- line spacing between wrapped note lines
- body-margin alignment

### App data

The desktop app stores user settings and logs here:

```text
Settings: C:\Users\<YourUser>\AppData\Roaming\JLJ IV Enterprises Inc\JLJ-invoice Rider\settings.json
Log file: C:\Users\<YourUser>\AppData\Roaming\JLJ IV Enterprises Inc\JLJ-invoice Rider\app.log
```

On first launch after the rename, the app copies any existing settings and log file from the older `Invoice Rule Studio` AppData folder into the new `JLJ-invoice Rider` folder.

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
- downloads and bundles the Tesseract installer into the setup build
- optionally compiles the Inno Setup installer

### Build outputs

- packaged app folder: `dist\JLJ-invoice Rider\`
- installer exe: `installer_output\JLJ-invoice Rider Setup.exe`

## Installer Notes

The installer is defined in `installer/JLJ-invoice Rider.iss` and currently provides:

- modern Inno Setup wizard
- desktop shortcut option
- bundled Tesseract installer so end users do not have to fetch OCR dependencies themselves
- optional Tesseract installation task during setup
- launch-after-install option

The installer copies the packaged app from `dist\JLJ-invoice Rider\` into the installation folder and can run `installer/install_tesseract.ps1` during setup. The bundled OCR installer targets the standard machine-wide Tesseract location under `C:\Program Files\Tesseract-OCR\`.

## Rule Configuration Notes

The OCR engine currently supports:

- month-style top dates such as `March 19, 2026`
- labeled dates such as `Invoice Date`
- auto-detection mode

Users can configure rule appearance from the app, including:

- font family
- text size adjustments, including smaller-than-before text
- bold styling
- text color
- line spacing between the due-date label and value
- line spacing between note lines
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
- the app log at `AppData\Roaming\JLJ IV Enterprises Inc\JLJ-invoice Rider\app.log`

### Installer build fails

Check:

- Python 3.14 is installed in `%LOCALAPPDATA%\Programs\Python\Python314`
- Inno Setup 6 is installed in `%LOCALAPPDATA%\Programs\Inno Setup 6`
- all Python build dependencies are installed

## Development Notes

- This repository is Windows-focused.
- Font detection currently relies on Windows font files under `C:\Windows\Fonts`.
- OCR accuracy depends on scan quality and Tesseract results.
