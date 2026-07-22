[Setup]
AppName=Excel Toolbox
AppVersion=1.4
AppPublisher=Nguyen Hieu
DefaultDirName={autopf}\Excel Toolbox
DefaultGroupName=Excel Toolbox
UninstallDisplayIcon={app}\Excel_Toolbox.exe
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=F:\Python\10.Ghep_Sheet_File\Output
OutputBaseFilename=Excel_Toolbox_Setup_v1.4
Compression=lzma
SolidCompression=yes
SetupIconFile=F:\Python\10.Ghep_Sheet_File\app_icon.ico
DisableProgramGroupPage=no

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "F:\Python\10.Ghep_Sheet_File\dist_v3\Excel_Toolbox.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Excel Toolbox"; Filename: "{app}\Excel_Toolbox.exe"
Name: "{autodesktop}\Excel Toolbox"; Filename: "{app}\Excel_Toolbox.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Excel_Toolbox.exe"; Description: "{cm:LaunchProgram,Excel Toolbox}"; Flags: nowait postinstall skipifsilent
