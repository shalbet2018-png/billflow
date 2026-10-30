; Built automatically on GitHub (.github/workflows/windows-installer.yml).
; By hand on Windows: run the two pyinstaller commands from that workflow, then ISCC installer\billflow.iss
[Setup]
AppName=billflow
AppVersion=0.2.0
AppPublisher=billflow
DefaultDirName={autopf}\billflow
DefaultGroupName=billflow
LicenseFile=..\LICENSE.txt
OutputDir=..\installer-out
OutputBaseFilename=billflow-setup-0.2.0
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
Compression=lzma2
SolidCompression=yes
ChangesEnvironment=yes

[Files]
Source: "..\dist\billflow\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\dist\billflow-gui\*"; DestDir: "{app}\gui"; Flags: recursesubdirs ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\docs\GUIDE.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\BILLFLOW-Business-Guide.docx"; DestDir: "{app}\docs"; Flags: ignoreversion

[Icons]
Name: "{group}\BILLFLOW"; Filename: "{app}\gui\billflow-gui.exe"; WorkingDir: "{app}\gui"
Name: "{autodesktop}\BILLFLOW"; Filename: "{app}\gui\billflow-gui.exe"; WorkingDir: "{app}\gui"; Tasks: desktopicon
Name: "{group}\billflow Business Guide"; Filename: "{app}\docs\BILLFLOW-Business-Guide.docx"
Name: "{group}\billflow README"; Filename: "{app}\README.md"
Name: "{group}\Uninstall billflow"; Filename: "{uninstallexe}"

[Tasks]
Name: desktopicon; Description: "Put a BILLFLOW shortcut on the desktop"
Name: addpath; Description: "Add billflow to PATH (use it from any terminal)"

[Registry]
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Tasks: addpath; Check: NeedsPath('{app}')

[Run]
Filename: "{app}\gui\billflow-gui.exe"; Description: "Open BILLFLOW now"; Flags: nowait postinstall skipifsilent

[Code]
function NeedsPath(Param: string): Boolean;
var P: string;
begin
  if not RegQueryStringValue(HKCU, 'Environment', 'Path', P) then begin Result := True; exit; end;
  Result := Pos(';' + Uppercase(ExpandConstant(Param)) + ';', ';' + Uppercase(P) + ';') = 0;
end;

