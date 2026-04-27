# ETA 설치형 배포 가이드

Windows 10/11 x64 설치 프로그램을 **GitHub Actions로 자동 생성**합니다.
- 설치 경로: `C:\Program Files\ETA` (관리자 권한)
- 바로가기: 바탕화면 + 시작 메뉴 (아이콘 `Assets/icons/ETA.ico`)
- 기존 버전 자동 덮어쓰기 (AppId 고정)
- .NET 8 런타임 내장 → 사용자 PC에 .NET 설치 불필요

## 🚀 기본 배포 절차 — 태그 하나로 자동 릴리스

### 1. 버전 올리기
`ETA.csproj` 의 `<Version>` 을 올리고 커밋:
```bash
git commit -am "bump v1.3.0"
```

### 2. 태그 푸시
```bash
git tag v1.3.0
git push origin v1.3.0
```

이 푸시가 [.github/workflows/release.yml](../.github/workflows/release.yml) 을 트리거합니다:
1. Windows 러너에서 .NET 8 SDK·Inno Setup 6 자동 설치
2. `dotnet publish` — self-contained single-file win-x64
3. Data/Templates, Assets, appsettings.json 스테이징
4. `installer.iss` 컴파일 → `ETA-Setup-1.3.0.exe`
5. GitHub **Releases** 에 태그 v1.3.0 이름으로 업로드

3~5분 후 `https://github.com/1110x/ETA/releases` 에 `.exe` 가 올라옵니다.

### 3. 사용자에게 배포
Releases 페이지의 `ETA-Setup-1.3.0.exe` 링크 공유.
사용자는 더블클릭 → 설치 마법사 따라가기 → 바탕화면 아이콘으로 실행.

### 4. (선택) Microsoft To Do 자동 알림
태그 푸시 직후 `Build/notify-release-todo.sh` 를 백그라운드로 실행해두면,
빌드가 끝나는 시점에 **"ETA 개발"** 리스트에 작업이 자동 등록됩니다.
작업 메모에는 인스톨러 직접 다운로드 URL · 파일명 · 크기, 릴리즈/워크플로 링크가 포함됩니다.

```bash
# 최신 태그 자동 사용
nohup ./Build/notify-release-todo.sh > /tmp/eta-todo.log 2>&1 &

# 특정 태그 지정
nohup ./Build/notify-release-todo.sh v1.4.3 > /tmp/eta-todo.log 2>&1 &
```

자격증명은 `Services/SERVICE1/TodoService.cs` 의 `ClientId` / `RefreshToken` 을 자동 추출합니다.
환경변수 `ETA_TODO_CLIENT_ID` / `ETA_TODO_REFRESH_TOKEN` 으로 덮어쓸 수도 있습니다.

## 🧪 수동 실행 (태그 없이 테스트 빌드)
GitHub 리포지토리 → Actions 탭 → **"Release Windows Installer"** → **Run workflow** → version 입력 (예: `1.3.0-rc1`).
이 경우 Release 는 생성되지 않고, **Artifacts** 에만 Setup.exe 가 올라갑니다 (30일 보존). 테스트용으로 다운로드만.

## 🔢 버전 관리
- `ETA.csproj` 의 `<Version>` 이 **단일 원본(single source of truth)**
- 태그 이름(`v1.3.0`) 에서 `v` 제외한 문자열이 설치 파일명·Inno Setup 버전으로 자동 주입 (`/DMyAppVersion=` 플래그)
- `installer.iss` 의 `#define MyAppVersion` 기본값은 개발·로컬 컴파일용 — CI 에서는 덮어쓰기됨

## 🔒 AppId
`installer.iss` 의 `AppId` 는 **절대 변경 금지**. 변경하면 Windows 가 별개 프로그램으로 인식해 기존 버전이 남습니다. 업그레이드 시 자동으로 기존 파일 덮어쓰기·프로세스 종료가 동작하려면 동일 GUID 필수.

## 💾 로컬 빌드 (선택사항)
Windows PC 에서 직접 빌드하려면:
```powershell
dotnet publish ETA.sln -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o Build\publish
# Inno Setup 6 설치 후
& 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' Build\installer.iss
```

macOS 에서 게시만(설치 파일 생성은 Windows 필요):
```bash
./Build/publish-win.sh
```

## 📦 설치 동작 요약
| 동작 | 대상 |
|------|------|
| 설치 경로 | `C:\Program Files\ETA` |
| 바로가기 | 바탕화면(옵션), 시작 메뉴 |
| 기존 버전 | 자동 감지 → 프로세스 종료 → 덮어쓰기 |
| 제거 시 삭제 | 프로그램 파일, `Logs/` |
| 제거 시 유지 | `Data/template_paths.json`, `appsettings.json` 등 사용자 설정 |

## 📝 체크리스트 (첫 릴리스)
- [ ] `ETA.csproj` `<Version>` 확인
- [ ] `appsettings.json` 의 MariaDB 연결 정보가 배포 환경용인지 확인
- [ ] `Assets/icons/ETA.ico` 해상도 확인 (권장: 16/32/48/256 포함)
- [ ] GitHub 리포지토리 Settings → Actions → General → Workflow permissions 가 **Read and write** 인지 확인 (Release 업로드 권한)
- [ ] 태그 푸시 → Actions 로그에서 5분 내 성공 확인
- [ ] Releases 페이지에서 Setup.exe 다운로드 → 테스트 설치
