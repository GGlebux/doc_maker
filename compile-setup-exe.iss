[Setup]
AppId={{A1408E5E-A591-416A-94DF-1B6F64F0B4FE}
AppName=приЁмка
AppVersion=2.0
AppPublisher=SASK, Inc.
AppPublisherURL=https://www.example.com/
AppSupportURL=https://www.example.com/
AppUpdatesURL=https://www.example.com/
DefaultDirName={autopf}\doc_maker
UninstallDisplayIcon={app}\main.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=mysetup
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Используем {src} для указания папки, где лежит .iss файл
Source: "output\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\rejng\OneDrive\Рабочий стол\resources\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKA; Subkey: "Software\Classes\.exe\OpenWithProgids"; ValueType: string; ValueName: "приЁмкаFile.exe"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\приЁмкаFile.exe"; ValueType: string; ValueName: ""; ValueData: "приЁмка File"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\приЁмкаFile.exe\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\main.exe,0"
Root: HKA; Subkey: "Software\Classes\приЁмкаFile.exe\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\main.exe"" ""%1"""

[Icons]
Name: "{autoprograms}\приЁмка"; Filename: "{app}\main.exe"
Name: "{autodesktop}\приЁмка"; Filename: "{app}\main.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\main.exe"; Description: "{cm:LaunchProgram,приЁмка}"; Flags: nowait postinstall skipifsilent