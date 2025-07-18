[Setup]
AppName=거래처별 엑셀 변환기
AppVersion=1.1
AppPublisher=EtcOrder
AppPublisherURL=https://github.com/your-repo
AppSupportURL=https://github.com/your-repo
AppUpdatesURL=https://github.com/your-repo
DefaultDirName={pf}\EtcOrder
DefaultGroupName=거래처별 엑셀 변환기
UninstallDisplayIcon={app}\icon.ico
OutputDir=.
OutputBaseFilename=EtcOrderSetup_v1.1
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\거래처별 엑셀 변환기"; Filename: "{app}\etc-order.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\거래처별 엑셀 변환기"; Filename: "{app}\etc-order.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\etc-order.exe"; Description: "{cm:LaunchProgram,거래처별 엑셀 변환기}"; Flags: nowait postinstall skipifsilent