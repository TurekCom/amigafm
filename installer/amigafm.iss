#define MyAppName "Amiga FM"
#ifndef MyAppVersion
#define MyAppVersion "0.1.0"
#endif
#define MyAppPublisher "TurekCom"
#define MyAppExeName "amiga_fm.exe"

[Setup]
AppId={{2B841B92-8F79-4D12-93FA-9D08F334D4D6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\Amiga FM
DefaultGroupName=Amiga FM
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=AmigaFM-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=lowest
SetupLogging=yes

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Files]
Source: "..\target\release\amiga_fm.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\target\release\nvdaControllerClient.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\readme.md"; DestDir: "{app}"; DestName: "README.md"; Flags: ignoreversion
Source: "..\CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\NOTICE.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\license.txt"; DestDir: "{app}"; DestName: "LICENSE-NVDA-CONTROLLER.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Amiga FM"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\Amiga FM"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Uruchom Amiga FM"; Flags: nowait postinstall skipifsilent
