[Setup]
AppName=Excel Toolbox
AppVersion=1.2
DefaultDirName={autopf}\Excel Toolbox
DefaultGroupName=Excel Toolbox
OutputDir=F:\Python\10.Ghep_Sheet_File\Output
OutputBaseFilename=Excel_Toolbox_Setup_v1.2
Compression=lzma
SolidCompression=yes
SetupIconFile=F:\Python\10.Ghep_Sheet_File\app_icon.ico

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "F:\Python\10.Ghep_Sheet_File\dist\Excel_Toolbox.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Excel Toolbox"; Filename: "{app}\Excel_Toolbox.exe"
Name: "{autodesktop}\Excel Toolbox"; Filename: "{app}\Excel_Toolbox.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Excel_Toolbox.exe"; Description: "{cm:LaunchProgram,Excel Toolbox}"; Flags: nowait postinstall skipifsilent
