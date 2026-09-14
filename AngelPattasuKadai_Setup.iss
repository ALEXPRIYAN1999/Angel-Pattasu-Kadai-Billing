; =========================================================
; ANGEL PATTASU KADAI - Inno Setup Installer
; Builds an installer from PyInstaller ONEDIR output.
;
; Expected build output:
;   dist\AngelPattasuKadai\
;       AngelPattasuKadai.exe
;       assets\
;           angel_logo.png
;           angel_logo.jpeg
;           logo_round.png
; =========================================================

#define MyAppName "Angel Pattasu Kadai"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Angel Pattasu Kadai"
#define MyAppExeName "AngelPattasuKadai.exe"

[Setup]
AppId={{8B6C4F0B-8C5D-4D55-A4D5-ANGELPATTASUKADAI}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Angel Pattasu Kadai
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=AngelPattasuKadai_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayName={#MyAppName}
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=angel_pattasu_kadai.ico
UninstallIconFile=angel_pattasu_kadai.ico
ChangesAssociations=no
AllowNoIcons=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Copy the complete PyInstaller ONEDIR folder, including all dependencies/assets.
Source: "dist\AngelPattasuKadai\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Angel Pattasu Kadai"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\Angel Pattasu Kadai"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Angel Pattasu Kadai"; Flags: nowait postinstall skipifsilent

[Dirs]
; Application data is intentionally NOT stored in Program Files.
; The Python application itself creates:
;   %USERPROFILE%\angel_data
; e.g. C:\Users\LENOVO\angel_data
Name: "{userappdata}\AngelPattasuKadai"; Flags: uninsneveruninstall
