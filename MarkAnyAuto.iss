; MarkAny 복호화 자동화 설치 스크립트 (Inno Setup 6)
; Build-MarkAnyAuto.ps1 이 ISCC "/DExePath=..." "/O<출력폴더>" 로 호출한다.
; 이 파일은 UTF-8 BOM으로 저장한다. BOM이 없으면 Inno가 한글을 깨뜨린다.

#ifndef ExePath
  #define ExePath "dist\MarkAnyAuto.exe"
#endif

[Setup]
AppId={{8F3C2A41-9B77-4E2D-9C1A-6B5D0E7A4C10}
AppName=MarkAny 복호화 자동화
AppVersion=1.0
DefaultDirName={autopf}\MarkAnyAuto
DefaultGroupName=MarkAny 복호화 자동화
PrivilegesRequired=lowest
OutputBaseFilename=MarkAnyAuto-Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "{#ExePath}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\MarkAny 복호화 자동화"; Filename: "{app}\MarkAnyAuto.exe"
Name: "{autodesktop}\MarkAny 복호화 자동화"; Filename: "{app}\MarkAnyAuto.exe"

[Run]
Filename: "{app}\MarkAnyAuto.exe"; Description: "지금 실행"; Flags: nowait postinstall skipifsilent
