[Setup]
AppId={{8B7C1E65-0A11-4C3B-9C14-2F5D8A6E9D01}
AppName=Asset Register
AppVersion=1.0
AppPublisher=Yabesh
DefaultDirName={autopf}\Asset Register
DefaultGroupName=Asset Register
DisableProgramGroupPage=auto
OutputDir=D:\YABESH\04-CODING TOOLS\AssetRegisterApp\AssetRegisterApp\Output
OutputBaseFilename=AssetRegisterAppSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible

SetupIconFile=D:\YABESH\04-CODING TOOLS\AssetRegisterApp\AssetRegisterApp\LOGO.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "D:\YABESH\04-CODING TOOLS\AssetRegisterApp\AssetRegisterApp\LOGO.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\YABESH\04-CODING TOOLS\AssetRegisterApp\AssetRegisterApp\dist\win-unpacked\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Notice the space in "Asset Register.exe" below! 
Name: "{group}\Asset Register"; Filename: "{app}\Asset Register.exe"; IconFilename: "{app}\LOGO.ico"
Name: "{autodesktop}\Asset Register"; Filename: "{app}\Asset Register.exe"; IconFilename: "{app}\LOGO.ico"; Tasks: desktopicon

[Run]
; Notice the space in "Asset Register.exe" below! 
Filename: "{app}\Asset Register.exe"; Description: "{cm:LaunchProgram,Asset Register}"; Flags: nowait postinstall skipifsilent