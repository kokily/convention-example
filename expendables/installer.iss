; Inno Setup Script for expendables-excel-converter
[Setup]
AppName=컨벤션 소모품 결산
AppVersion=1.0
AppPublisher=dnkdream
DefaultDirName={pf}\컨벤션 소모품 결산
DefaultGroupName=컨벤션 소모품 결산
OutputDir=.
OutputBaseFilename=convention-expendables-setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes

[Files]
Source: "bin/Debug/net9.0-windows/*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\컨벤션 소모품 결산"; Filename: "{app}\expendables-excel-converter.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\컨벤션 소모품 결산"; Filename: "{app}\expendables-excel-converter.exe"; WorkingDir: "{app}"

[Run]
Filename: "{app}\expendables-excel-converter.exe"; Description: "프로그램 실행"; Flags: nowait postinstall skipifsilent 