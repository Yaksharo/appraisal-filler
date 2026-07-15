; Inno Setup script for Advisee Document Filler
; Installs to C:\Program Files\Yaksharo Solutions\Advisee Document Filler

#ifndef MyAppVersion
  #define MyAppVersion "1.0"
#endif

[Setup]
AppId={{7E2B9C41-5A83-4D1F-9B67-A2BAD5F11E01}
AppName=Advisee Document Filler
AppVersion={#MyAppVersion}
AppPublisher=Yaksharo Solutions
DefaultDirName={autopf}\Yaksharo Solutions\Advisee Document Filler
DefaultGroupName=Yaksharo Solutions
DisableProgramGroupPage=yes
OutputBaseFilename=AdviseeDocFiller-Setup-{#MyAppVersion}
SetupIconFile=assets\logo.ico
UninstallDisplayIcon={app}\AdviseeDocFiller.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\AdviseeDocFiller\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\Advisee Document Filler"; Filename: "{app}\AdviseeDocFiller.exe"
Name: "{autodesktop}\Advisee Document Filler"; Filename: "{app}\AdviseeDocFiller.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\AdviseeDocFiller.exe"; Description: "{cm:LaunchProgram,Advisee Document Filler}"; Flags: nowait postinstall skipifsilent
