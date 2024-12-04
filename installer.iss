[Setup]
AppName=Doc Converter
AppVersion=1.0
DefaultDirName={pf}\DocConverter
DefaultGroupName=Doc Converter
OutputBaseFilename=DocConverterInstaller
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
AlwaysRestart=no
AllowNoIcons=yes
AlwaysShowDirPage=yes
DisableDirPage=no
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\doc_converter_gui.exe
Uninstallable=yes
CreateUninstallRegKey=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes

[Files]
Source: "dist\doc_converter_gui.exe"; DestDir: "{app}"; Flags: ignoreversion replacesameversion

[Icons]
Name: "{group}\Doc Converter"; Filename: "{app}\doc_converter_gui.exe"
Name: "{commondesktop}\Doc Converter"; Filename: "{app}\doc_converter_gui.exe"

[Run]
Filename: "{app}\doc_converter_gui.exe"; Description: "Launch Doc Converter"; Flags: postinstall nowait skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\*"
Type: dirifempty; Name: "{app}"
