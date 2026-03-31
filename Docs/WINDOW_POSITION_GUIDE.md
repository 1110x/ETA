# 창 위치 저장/복원 기능 구현 가이드

## 개요
ETA 애플리케이션의 각 페이지(ActivePage1-4)의 창 위치와 레이아웃을 사용자별로 저장하고 복원하는 기능이 구현되었습니다.

## 구현 파일

### 1. WindowPositionManager.cs (Services/)
- **역할**: 페이지 레이아웃 정보 저장/복원 관리
- **주요 메서드**:
  - `SaveLayouts()`: 레이아웃을 PageHW.log 파일에 저장
  - `GetPageLayout(pageName)`: 저장된 레이아웃 조회
  - `SavePageLayout(pageName, layoutInfo)`: 특정 페이지의 레이아웃 저장

### 2. CurrentUserManager.cs (Services/)
- **역할**: 현재 로그인한 사용자 정보 관리
- **기능**: 싱글톤 패턴으로 사용자 ID 관리

### 3. MainPage.axaml.cs (Views/)
- **변경 사항**:
  - `_positionManager` 필드 추가
  - `MainPage_Opened/Closing` 이벤트 핸들러 추가
  - `SaveCurrentModeLayout()`: 현재 모드의 레이아웃 저장
  - `RestoreModeLayout(modeName)`: 모드별 레이아웃 복원
  - `UpdateCurrentUser(newUserId)`: 사용자 변경 시 호출
  - 각 메뉴 클릭 메서드 끝에 `RestoreModeLayout()` 호출 추가

## 파일 경로

### 저장 위치
```
C:\Users\{WindowsUsername}\AppData\Roaming\ETA\Users\{ApplicationUserId}\PageHW.log
```

예:
```
C:\Users\ironu\AppData\Roaming\ETA\Users\DefaultUser\PageHW.log
```

### 로그 파일 형식 (JSON)
```json
{
  "Mode_Agent": {
    "windowX": 100,
    "windowY": 100,
    "windowWidth": 1400,
    "windowHeight": 800,
    "content2Star": 1,
    "content4Star": 1,
    "upperStar": 1,
    "lowerStar": 0,
    "leftPanelWidth": 260,
    "savedAt": "2026-03-27T10:30:45"
  },
  "Mode_Quotation": {
    "windowX": 100,
    "windowY": 100,
    "windowWidth": 1600,
    "windowHeight": 900,
    "content2Star": 7,
    "content4Star": 3,
    "upperStar": 13,
    "lowerStar": 4,
    "leftPanelWidth": 430,
    "savedAt": "2026-03-27T10:31:20"
  }
}
```

## 작동 방식

### 1. 애플리케이션 시작
```csharp
// MainPage 생성자에서
_positionManager = new WindowPositionManager(CurrentUserManager.Instance.CurrentUserId);
```

### 2. 메뉴 클릭 시
```csharp
private void Agent_Click(object? sender, RoutedEventArgs e)
{
    _currentMode = "Agent";
    // ... UI 설정 ...
    
    // 저장된 레이아웃 복원
    RestoreModeLayout("Agent");
}
```

### 3. 윈도우 종료 시
```csharp
private void MainPage_Closing(object? sender, WindowClosingEventArgs e)
{
    // 현재 모드의 레이아웃 자동 저장
    SaveCurrentModeLayout();
}
```

### 4. 사용자 로그인 후
```csharp
// 로그인 페이지에서
mainPage.UpdateCurrentUser(loggedInUserId);
```

## 저장되는 레이아웃 정보

### PageLayoutInfo 클래스
| 속성 | 설명 | 기본값 |
|------|------|--------|
| WindowX | 윈도우 X 좌표 | 100 |
| WindowY | 윈도우 Y 좌표 | 100 |
| WindowWidth | 윈도우 너비 | 1400 |
| WindowHeight | 윈도우 높이 | 800 |
| Content2Star | Content2/Content4 분할 비율 | 1 |
| Content4Star | Content2/Content4 분할 비율의 우측 | 1 |
| UpperStar | 상단 영역 행 높이 | 4 |
| LowerStar | 하단 영역 행 높이 | 1 |
| LeftPanelWidth | 왼쪽 패널(Content1) 너비 | 260 |

## 설정 방법

### 1. 기본 설정 (자동)
- Windows 사용자명을 기반으로 자동 설정
- 앱 실행 시 로그 파일이 자동 생성

### 2. 커스텀 사용자 설정
```csharp
// 로그인 페이지에서
mainWindow.UpdateCurrentUser("employeeId_12345");
```

### 3. 레이아웃 초기화
```csharp
// 필요시 저장된 레이아웃 모두 삭제
_positionManager?.ClearAllLayouts();
```

## 각 메뉴별 기본 레이아웃

| 메뉴 | Content2Star | Content4Star | UpperStar | LowerStar | LeftPanelWidth |
|------|-------------|-------------|----------|----------|----------------|
| Agent | 1 | 1 | 1 | 0 | 260 |
| WasteCompany | 1 | 0 | 1 | 0 | 260 |
| Analysis | 1 | 1 | 1 | 0 | 380 |
| Contract | 1 | 0 | 8 | 2 | 350 |
| Repair | 1 | 0 | 7 | 3 | 220 |
| Quotation | 7 | 3 | 13 | 4 | 430 |
| Purchase | 1 | 0 | 8 | 2 | 250 |
| TestReport | 8 | 2 | 8.5 | 1.5 | (다양함) |

## 디버깅

### 로그 파일 경로 확인
```csharp
string logPath = mainPage.GetPositionLogFilePath();
Debug.WriteLine($"Log path: {logPath}");
```

### 저장된 레이아웃 확인
```csharp
var layouts = _positionManager?.GetAllLayouts();
foreach (var (key, info) in layouts)
{
    Debug.WriteLine($"{key}: {info}");
}
```

## 남은 작업 (추가 구현 권장)

1. **다른 메뉴 클릭 메서드에 복원 호출 추가**
   - `Quotation_Click()`
   - `Purchase_Click()`
   - `TestReport_Click()`
   - `Admin_Click()`
   - `Permission_Click()`

2. **로그인 페이지 통합**
   - Login.axaml.cs에서 로그인 성공 후 `UpdateCurrentUser()` 호출

3. **UI 개선**
   - 설정 메뉴에서 레이아웃 초기화 옵션 추가
   - 각 메뉴별로 저장된 레이아웃 미리보기

4. **성능 최적화**
   - 비동기 저장 처리
   - 파일 쓰기 간격 최적화

## 예제: 로그인 후 사용자 설정

```csharp
// LoginWindow.axaml.cs
private void LoginSuccessful(string employeeId)
{
    // ... 로그인 처리 ...
    
    // MainPage 생성망 제어
    MainPage mainPage = new MainPage();
    mainPage.UpdateCurrentUser(employeeId);
    mainPage.Show();
    
    this.Close();
}
```

## 기술 스택

- **WinAPI**: `Environment.GetFolderPath()` for AppData path
- **JSON**: `System.Text.Json` for serialization
- **Avalonia**: Control positioning and grid layout management
- **C# 11+**: Records, nullable types, pattern matching
