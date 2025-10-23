; setup.iss
#define Version "1.0.0"

#ifndef ROOT
  #define ROOT "."
#endif

[Setup]
AppId={{b4ed3766-a828-4f4a-8f71-eec090de8894}}
AppName=Pub-Xel
AppVersion={#Version}
AppPublisher=Pub-Xel Project
AppVerName=Pub-Xel {#Version}
PrivilegesRequired=lowest
DefaultDirName={userappdata}\pubxel
DefaultGroupName=Pub-Xel
OutputBaseFilename=Pub-Xel_Installer_v{#Version}
OutputDir=Output
ChangesAssociations=no
DisableDirPage=no
DisableProgramGroupPage=no
Compression=lzma
SolidCompression=yes
LicenseFile={#ROOT}\LICENSE.txt

[Files]
Source: "{#ROOT}\dist\Pub-Xel.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Pub-Xel"; Filename: "{app}\Pub-Xel.exe"; Tasks: "startmenuicon"
Name: "{userdesktop}\Pub-Xel"; Filename: "{app}\Pub-Xel.exe"; Tasks: "desktopicon"
Name: "{group}\Uninstall Pub-Xel"; Filename: "{uninstallexe}"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startmenuicon"; Description: "Create a &Start Menu shortcut"; GroupDescription: "Additional icons:"

[Registry]
Root: HKCU; Subkey: "Software\Pub-Xel"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[Run]
Filename: "{app}\Pub-Xel.exe"; Description: "Launch Pub-Xel"; Flags: nowait postinstall skipifsilent