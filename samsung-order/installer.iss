[Setup]
AppName=삼성웰스토리 엑셀 전처리기
AppVersion=1.0
AppPublisher=삼성웰스토리
DefaultDirName={pf}\삼성웰스토리엑셀전처리기
DefaultGroupName=삼성웰스토리엑셀전처리기
OutputDir=.
OutputBaseFilename=삼성웰스토리_발주서_전처리기_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\samsung-order.exe

[Files]
Source: "bin\Release\net9.0-windows\win-x64\publish\samsung-order.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net9.0-windows\win-x64\publish\samsung-order.pdb"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\삼성웰스토리 엑셀 전처리기"; Filename: "{app}\samsung-order.exe"
Name: "{commondesktop}\삼성웰스토리 엑셀 전처리기"; Filename: "{app}\samsung-order.exe"

[Run]
Filename: "{app}\samsung-order.exe"; Description: "삼성웰스토리 엑셀 전처리기 실행"; Flags: postinstall nowait 