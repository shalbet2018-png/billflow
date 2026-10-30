; Build: ISCC installer\billflow.iss  (after: pyinstaller ... -> dist\billflow)
[Setup]
AppName=billflow
AppVersion=0.1.0
AppPublisher=billflow
DefaultDirName={autopf}\billflow
DefaultGroupName=billflow
LicenseFile=..\LICENSE.txt
OutputDir=..\installer-out
OutputBaseFilename=billflow-setup-0.1.0
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
Compression=lzma2
SolidCompression=yes
ChangesEnvironment=yes

[Files]
Source: "..\dist\billflow\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\billflow README"; Filename: "{app}\README.md"
Name: "{group}\Uninstall billflow"; Filename: "{uninstallexe}"

[Tasks]
Name: addpath; Description: "Add billflow to PATH (use it from any terminal)"

[Registry]
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Tasks: addpath; Check: NeedsPath('{app}')

[Code]
function NeedsPath(Param: string): Boolean;
var P: string;
begin
  if not RegQueryStringValue(HKCU, 'Environment', 'Path', P) then begin Result := True; exit; end;
  Result := Pos(';' + Uppercase(ExpandConstant(Param)) + ';', ';' + Uppercase(P) + ';') = 0;
end;
