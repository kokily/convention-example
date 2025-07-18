[Setup]
AppName=삼성웰스토리 발주서 전처리기
AppVersion=1.0
AppPublisher=삼성웰스토리
DefaultDirName={pf}\삼성웰스토리발주서전처리기
DefaultGroupName=삼성웰스토리발주서전처리기
OutputDir=.
OutputBaseFilename=삼성웰스토리_발주서_전처리기_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\samsung-order.exe
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "bin\Release\net9.0-windows\win-x64\publish\samsung-order.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net9.0-windows\win-x64\publish\samsung-order.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\삼성웰스토리 발주서 전처리기"; Filename: "{app}\samsung-order.exe"; IconFilename: "{app}\icon.ico"
Name: "{commondesktop}\삼성웰스토리 발주서 전처리기"; Filename: "{app}\samsung-order.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\samsung-order.exe"; Description: "삼성웰스토리 발주서 전처리기 실행"; Flags: postinstall nowait

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

function IsDotNetDetected(version: string; release: cardinal): boolean;
var
  success: boolean;
  release45: cardinal;
  keyValue: string;
begin
  success := false;
  release45 := 0;

  // .NET 4.5 이상 체크
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\', 'Release', release45) then
  begin
    if (release45 >= release) then
    begin
      success := true;
    end;
  end;

  // .NET Core/5+ 체크
  if RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{92FB6C44-E685-45AD-9B20-CADF4CABA132}  - ' + version) then
  begin
    success := true;
  end;

  if RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{92FB6C44-E685-45AD-9B20-CADF4CABA132}  - ' + version) then
  begin
    success := true;
  end;

  Result := success;
end; 