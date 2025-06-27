[Setup]
AppName=삼성웰스토리 결산서 전처리기
AppVersion=1.0
DefaultDirName={pf}\SamsungWelstoryStatement
DefaultGroupName=삼성웰스토리 결산서 전처리기
OutputDir=.
OutputBaseFilename=SamsungWelstoryStatementSetup
Compression=lzma
SolidCompression=yes
SetupIconFile=icon.ico

[Files]
Source: "bin\Release\net9.0-windows\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\삼성웰스토리 결산서 전처리기"; Filename: "{app}\samsung-statements.exe"
Name: "{group}\프로그램 제거"; Filename: "{uninstallexe}" 