; ETA LIMS — Inno Setup 설치 스크립트
; 사용법(Windows):
;   1) iscc.exe Build\installer.iss  또는  Inno Setup Compiler 에서 열어 [Compile]
;   2) 결과물: Build\Output\ETA-Setup-1.2.0.exe
;
; 동작:
;  - 기본 설치 경로: C:\Program Files\ETA  (관리자 권한 요청)
;  - 바탕화면 + 시작메뉴 바로가기 생성 (아이콘: Assets/icons/ETA.ico)
;  - 기존 버전 자동 삭제(덮어쓰기) — AppId 고정으로 업그레이드 처리
;  - 설치 후 앱 실행 옵션 제공

#define MyAppName         "ETA"
#define MyAppNameFull     "ETA LIMS"
#define MyAppVersion      "1.4.1"
#define MyAppPublisher    "ETA"
#define MyAppExeName      "ETA.exe"
#define MyPublishDir      "publish"
; 아이콘은 publish 복사본을 기준으로 참조 (Build 폴더만 옮겨도 동작)
#define MyIconFile        MyPublishDir + "\Assets\icons\ETA.ico"

[Setup]
; AppId 는 고정 GUID — 기존 버전 인식·업그레이드용
AppId={{7E2A4C18-4D1E-4E8B-9F5A-ETA-0001-LIMS}}
AppName={#MyAppNameFull}
AppVersion={#MyAppVersion}
AppVerName={#MyAppNameFull} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=ETA-Setup-{#MyAppVersion}
SetupIconFile={#MyIconFile}
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; 관리자 권한으로 Program Files 쓰기
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; 업그레이드 시 기존 버전 자동 제거
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 생성"; GroupDescription: "추가 아이콘:"; Flags: checkedonce

[Files]
; publish 폴더 전체를 {app} 에 복사
Source: "{#MyPublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppNameFull}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\icons\ETA.ico"
Name: "{group}\{#MyAppNameFull} 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppNameFull}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\icons\ETA.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "ETA 실행"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 사용자 설정 파일은 삭제하지 않음 — Data/template_paths.json, appsettings.json 등 유지 의도면 여기에 항목 추가 안 함
Type: filesandordirs; Name: "{app}\Logs"
