[Setup]
AppName=거래처별 엑셀 변환기
AppVersion=1.0
DefaultDirName={pf}\EtcOrder
DefaultGroupName=거래처별 엑셀 변환기
UninstallDisplayIcon={app}\icon.ico
OutputDir=.
OutputBaseFilename=EtcOrderSetup
SetupIconFile=icon.ico

[Files]
Source: "publish\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\거래처별 엑셀 변환기"; Filename: "{app}\etc-order.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\거래처별 엑셀 변환기"; Filename: "{app}\etc-order.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 생성"; GroupDescription: "추가 아이콘:"