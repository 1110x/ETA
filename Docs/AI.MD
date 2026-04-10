# AI 작업 로그

---

## 2026-04-10

### AI 학습 메뉴 신규 구현 (AI 문서분류)

**파일**:
- `Views/MainPage.axaml` — "AI 학습 > AI 문서분류" MenuItem 추가
- `Views/MainPage.axaml.cs` — `AiDocClassification_Click`, BT1(학습), BT2(새로고침) 케이스
- `Views/Pages/PAGE5/AiDocClassificationPage.axaml` — 트리뷰 AXAML (신규)
- `Views/Pages/PAGE5/AiDocClassificationPage.axaml.cs` — 전체 로직 (신규)

**화면 구성**:
- Show1: 파서/분석항목 트리뷰 (13종, 파일 수 배지 + AI 학습 완료 표시)
- Show2: `+ 파일 추가` 버튼 + 파일 목록 리스트
- Show3: 선택 파일 미리보기 또는 AI 학습 패턴 결과
- 서브메뉴: `학습` (BT1) → Claude API 호출, `새로고침` (BT2)

**DB 테이블** (앱 시작 시 자동 생성):
- `AI_문서분류_학습데이터`: 파서타입 + 파일명 + 파일형식 + BLOB 데이터
- `AI_문서분류_패턴`: 파서타입 UNIQUE + Claude 분석 결과 텍스트

**AI 학습 흐름**:
1. 사용자가 Show1에서 분석항목 선택 (예: "BOD")
2. Show2의 `+ 파일 추가` 클릭 → 파일 선택 → DB 저장 → 목록 갱신
3. 서브메뉴 `학습` 클릭 → DB 파일 읽기 → Claude API 전송 → 패턴 저장 → Show3 결과 표시

**Claude API 설정**:
- `appsettings.json`에 `"AnthropicApiKey": "sk-ant-..."` 추가 필요
- 모델: `claude-haiku-4-5-20251001` (빠른 분석용)
- CSV/TXT: 텍스트 추출 전송, Excel/PDF: 파일명+메타데이터 전송
- 파일당 최대 3,000자, 한 번에 최대 5개 파일 전송

**수정된 기존 버그**:
- `WasteSampleService.cs` — git stash pop 후 `GetRecentSamples` 중복 메서드 제거

---

## 2026-04-05 (오늘)

### 1. 견적발행내역 🔵🔴 아이콘 배지에 가려지는 문제 수정
**파일**: `Views/Pages/PAGE1/QuotationHistoryPanel.axaml.cs`

**원인**: `MakeIssueLeaf()`에서 아이콘 TextBlock의 FontSize가 `FontSizeXS`(9pt)로 너무 작아 Grid 컬럼 0이 0px로 계산되어 배지(컬럼 1)가 아이콘 위를 덮었음.

**수정**:
- `FontSizeXS` → `FontSizeSM` 으로 상향
- `MinWidth = 16` 추가 (컬럼 최소 폭 확보)
- 오른쪽 마진 `3 → 6 px`으로 확대

---

### 2. 견적발행내역 TreeViewItem expander 공간 잘림 수정
**파일**: `Views/Pages/PAGE1/QuotationHistoryPanel.axaml`

**원인**: Avalonia Fluent 테마가 자식 없는 리프 TreeViewItem에도 `PART_ExpandCollapseChevron` ToggleButton(▶) 공간 19px를 예약하여, 그 뒤에 있는 아이콘이 잘려 보임.

**수정**:
- 리프 노드: `PART_ExpandCollapseChevron` → `Width=12, Opacity=0` (숨김)
- 부모 노드(`:hasitems`): `Width=19, Opacity=1` (원래대로 복원)
- `ClipToBounds="False"`: `treeHostBorder`, `treeHost Grid`, `TreeViewItem` 스타일에 추가

---

### 3. QuotationCheckPanel '전체(N)' 체크박스 위치 개선
**파일**: `Views/Pages/PAGE1/QuotationCheckPanel.axaml.cs`

**변경 전**: `전체 (N)` 체크박스가 항목 스크롤 영역(itemsStack) 맨 위에 독립 행으로 존재.

**변경 후**: 카테고리명 TextBlock 옆에 가로로 나란히 배치.
- `headerText` + `catChk`를 `StackPanel { Orientation=Horizontal }`로 묶어 Row 0에 배치
- `catChk`는 `itemsStack`에서 제거, 헤더 패널로 이동

```
변경 전:  [일반항목]
           ☑ 전체 (12)
           ☑ pH
           ☑ DO ...

변경 후:  [일반항목]  ☑ 전체 (12)
           ☑ pH
           ☑ DO ...
```

---

### 4. 앱 시작 시 기본 화면 설정
**파일**: `Views/MainPage.axaml.cs`

**변경 내용**: `Loaded` 이벤트 핸들러 끝에 `WasteCompany_Click()` 호출 추가.

**효과**: 앱 실행 시 자동으로 **비용부담금 > 폐수배출업소** 화면이 열림.

---

### 이전 세션 (v18.9.0 태그 기준)

> 아래는 `ce1e24d` 커밋(v18.9.0)에 포함된 내용 요약

#### BOD_DATA / *_DATA MariaDB 서버 저장 구현
**파일**: `Services/SERVICE2/WasteSampleService.cs`, `Views/Pages/PAGE2/WasteAnalysisInputPage.axaml.cs`

- BOD_DATA, SS_DATA, NHexan_DATA, TN_DATA, TP_DATA, Phenols_DATA 테이블에 원시 측정값 UPSERT 구현
- `UpsertBodData()`, `UpsertSimpleData()`, `UpsertUvvisData()` 3개 메서드 추가
- `SaveRawData(ExcelRow, WasteSample)` 헬퍼 추가, `ImportWasteSample / SaveCurrentSample` 양쪽에서 호출

#### Shimmer 호버 해제 후 글자 흰색 고정 버그 수정
**파일**: `Views/ShimmerEffect.cs`

**원인**: `StopHover`에서 `ClearValue(ForegroundProperty)` → TextBlock에 로컬 Foreground가 없어지면서 프레임워크 기본값(흰색)으로 떨어짐.

**수정**: `_active` 딕셔너리에 `origBrushes[]` 추가. `StartHover/StartPanelShimmer` 시작 시 원본 색상 저장, `StopHover/StopPanelShimmer` 시 복원.
