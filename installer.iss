; Inno Setup script — turns dist\QuoteGenerator.exe into a normal Windows
; installer: QuoteGenerator-Setup.exe. Staff double-click it, click
; "Install", and get a "Quote Generator" icon on the Desktop and Start
; menu. No choices, no technical steps.

[Setup]
AppName=Quote Generator
AppVersion=1.0
AppPublisher=Your Agency
DefaultDirName={autopf}\Quote Generator
DisableProgramGroupPage=yes
DisableDirPage=yes
DisableReadyPage=yes
OutputDir=installer-output
OutputBaseFilename=QuoteGenerator-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\QuoteGenerator.exe

[Files]
Source: "dist\QuoteGenerator.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\Quote Generator"; Filename: "{app}\QuoteGenerator.exe"
Name: "{autoprograms}\Quote Generator"; Filename: "{app}\QuoteGenerator.exe"

[Run]
Filename: "{app}\QuoteGenerator.exe"; Description: "Open Quote Generator now"; Flags: nowait postinstall skipifsilent
