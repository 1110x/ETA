# Shimmer 효과 적용 가이드

## 현재 구조

`Views/ShimmerEffect.cs` 의 `TextShimmer` 클래스가 3가지 방식으로 shimmer를 제공:

| 메서드 | 용도 | 동작 |
|--------|------|------|
| `AttachHover(ctrl)` | 개별 컨트롤 (Button, TreeViewItem 등) | 호버 시 내부 TextBlock에 그라데이션 sweep |
| `AttachPanelHover(panel)` | Show 패널처럼 내용이 동적 교체되는 Border | 호버 진입 시 현재 자식 TextBlock 최대 20개 스캔 후 shimmer |
| `AttachAll(root)` | 윈도우 전체 비주얼 트리 일괄 부착 | 테마 변경 시 호출 (ApplyTheme 마지막) |
| `AttachIfNew(ctrl)` | 동적 생성된 개별 컨트롤 | `_attached` HashSet으로 중복 방지 |

## Show1/2/3에 shimmer가 안 들어간 이유

### 원인 1: Show2/Show3 패널에 `AttachPanelHover` 미등록
- `MainPage.axaml.cs` L121~128에서 **Show1Border, Content4Border**만 등록
- **Show2, Show3에 대한 `AttachPanelHover` 호출이 없음**

### 원인 2: 동적 UI 컨트롤에 `AttachIfNew` 미호출
- `LoadVerifiedGrid()` 등에서 동적으로 생성되는 Border/TextBlock/Grid 행에 shimmer 부착 안 됨
- `AttachAll()`은 테마 변경 시만 호출되므로, 이후 동적 생성된 UI에는 적용 안 됨

### 원인 3: `LogContentChange`에서 shimmer 코드 제거됨
- L4823: `// (shimmer 제거됨)` — 패널 내용이 교체될 때 새 컨텐츠에 shimmer를 부착하던 코드가 삭제된 상태

## 해결 방법

### 1. Show2/Show3 Border에 `AttachPanelHover` 등록
```csharp
// MainPage.axaml.cs Loaded 이벤트에 추가
var show2Border = this.FindControl<Border>("Show2Border");
var show3Border = this.FindControl<Border>("Show3Border");
if (show2Border != null) TextShimmer.AttachPanelHover(show2Border);
if (show3Border != null) TextShimmer.AttachPanelHover(show3Border);
```

### 2. 동적 UI 생성 후 `AttachIfNew` 호출
```csharp
// LoadVerifiedGrid() 등에서 Border/행을 생성한 후:
TextShimmer.AttachIfNew(border);
```

### 3. LogContentChange에서 패널 교체 시 shimmer 재부착
```csharp
// LogContentChange에서 content가 Control일 때:
if (content is Control ctrl)
    TextShimmer.AttachIfNew(ctrl);
```

## 주의사항
- `AttachPanelHover`는 호버 진입 시 TextBlock을 **최대 20개**만 스캔 (성능 보호)
- 데이터 그리드처럼 행이 많은 경우 개별 행에 `AttachIfNew`를 붙이면 성능 이슈 가능
- 헤더/요약 영역만 shimmer 대상으로 하는 것이 적절
