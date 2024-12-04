[Setup]
AppName=Doc Converter
AppVersion=2024.12.-170613
DefaultDirName={pf}\DocConverter
DefaultGroupName=Doc Converter
OutputBaseFilename=DocConverterInstaller_v2024.12.-170613
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
AlwaysRestart=no
AllowNoIcons=yes
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\doc_converter_gui.exe
Uninstallable=yes
CreateUninstallRegKey=yes

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
