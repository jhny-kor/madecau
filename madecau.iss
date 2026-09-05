; madecau 설치 스크립트 (Inno Setup 6)
; Build-madecau.ps1 이 ISCC "/DExePath=..." "/O<출력폴더>" 로 호출한다.
; 이 파일은 UTF-8 BOM으로 저장한다. BOM이 없으면 Inno가 한글을 깨뜨린다.

#ifndef ExePath
  #define ExePath "dist\madecau.exe"
#endif

[Setup]
AppId={{8F3C2A41-9B77-4E2D-9C1A-6B5D0E7A4C10}
AppName=madecau
AppPublisher=김지현
AppVersion=1.0
DefaultDirName={autopf}\madecau
DefaultGroupName=madecau
PrivilegesRequired=lowest
OutputBaseFilename=madecau-Setup
SetupIconFile=madecau.ico
UninstallDisplayIcon={app}\madecau.exe
Compression=lzma2
SolidCompression=yes

[InstallDelete]
Type: files; Name: "{app}\MarkAnyAuto.exe"
Type: files; Name: "{autodesktop}\MarkAny 복호화 자동화.lnk"
Type: filesandordirs; Name: "{autoprograms}\MarkAny 복호화 자동화"

[Files]
Source: "{#ExePath}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\madecau"; Filename: "{app}\madecau.exe"
Name: "{autodesktop}\madecau"; Filename: "{app}\madecau.exe"

[Run]
Filename: "{app}\madecau.exe"; Description: "지금 실행"; Flags: nowait postinstall skipifsilent
