; setup.iss
#ifndef Version
  #define Version "1.0.0"
#endif

#ifndef ROOT
  #define ROOT "."
#endif

[Setup]
AppId={{b4ed3766-a828-4f4a-8f71-eec090de8894}}
AppName=Pub-Xel
AppVersion={#Version}
AppVerName=Pub-Xel {#Version}
AppPublisher=Pub-Xel Project
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
UsePreviousTasks=no
AppModifyPath={app}\Pub-Xel.exe
UsePreviousAppDir=yes

[Code]
var
  IsUpdateInstall: Boolean;

function GetPrevInstallDir: string;
var
  Prev: string;
begin
  if RegQueryStringValue(HKCU, 'Software\Pub-Xel', 'InstallPath', Prev) and DirExists(Prev) then
    Result := Prev
  else
    Result := '';
end;

procedure InitializeWizard;
var
  Prev: string;
begin
  Prev := GetPrevInstallDir;
  IsUpdateInstall := Prev <> '';

  if IsUpdateInstall then
  begin
    WizardForm.Caption := 'Pub-Xel Update';
    WizardForm.DirEdit.Text := Prev;
    WizardForm.DirBrowseButton.Visible := False;
  end;
end;

function ShouldTaskBeChecked(const TaskName: String): Boolean;
begin
  if IsUpdateInstall then
  begin
    if (TaskName = 'autostart') then
      Result := True
    else if (TaskName = 'desktopicon') or (TaskName = 'startmenuicon') then
      Result := False
    else
      Result := False;
  end
  else
    Result := True;  // First install: all checked
end;

[Files]
Source: "{#ROOT}\dist\Pub-Xel.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Pub-Xel"; Filename: "{app}\Pub-Xel.exe"; Tasks: "startmenuicon"
Name: "{userdesktop}\Pub-Xel"; Filename: "{app}\Pub-Xel.exe"; Tasks: "desktopicon"
Name: "{group}\Uninstall Pub-Xel"; Filename: "{uninstallexe}"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startmenuicon"; Description: "Create a &Start Menu shortcut"; GroupDescription: "Additional icons:"
Name: "autostart"; Description: "Run Pub-Xel on startup"; GroupDescription: "Startup options:"; Flags: checkedonce 

[Registry]
Root: HKCU; Subkey: "Software\Pub-Xel"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "Pub-Xel"; ValueData: """{app}\Pub-Xel.exe"""; \
    Flags: uninsdeletevalue; Tasks: autostart

[Run]
Filename: "{app}\Pub-Xel.exe"; Description: "Launch Pub-Xel"; Flags: nowait postinstall skipifsilent