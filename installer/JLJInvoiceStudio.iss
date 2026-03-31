#define MyAppName "JLJ Invoice Rule Studio"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "JLJ IV Enterprises Inc."
#define MyAppExeName "JLJInvoiceStudio.exe"

[Setup]
AppId={{A740D509-DAF9-4D9D-B95D-8D69FDD7A8F9}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://jlj.example.invalid
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
WizardStyle=modern
PrivilegesRequired=lowest
Compression=lzma2/max
SolidCompression=yes
OutputDir=..\installer_output
OutputBaseFilename=JLJInvoiceStudioSetup
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=..\assets\jlj_invoice.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Invoice automation desktop studio
VersionInfoProductName={#MyAppName}

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"
Name: "installtesseract"; Description: "Install Tesseract OCR dependency (recommended)"; Flags: checkedonce

[Files]
Source: "..\dist\JLJInvoiceStudio\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\installer\install_tesseract.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\install_tesseract.ps1"""; Flags: runhidden waituntilterminated; Tasks: installtesseract
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
