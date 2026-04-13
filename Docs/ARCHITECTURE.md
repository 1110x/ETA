# ETA 프로젝트 — 아키텍처 레퍼런스 (Claude Code용)

> 이 문서는 AI 코딩 어시스턴트(Claude Code 등)가 ETA 코드베이스를 빠르게 이해하고
> 올바른 패턴으로 수정할 수 있도록 작성된 기술 레퍼런스입니다.  
> 마지막 갱신: 2026-03-29

---

## 1. 프레임워크 & 의존성

| 항목 | 값 |
|------|-----|
| Framework | .NET 8.0 |
| UI | Avalonia 11.3.12 (FluentTheme, Dark 기본) |
| 폰트 | KBIZ한마음고딕 M (avares://ETA/Assets/Fonts#KBIZ한마음고딕 M) |
| ORM | 없음 — 순수 ADO.NET (DbConnection/DbCommand 직접) |
| DB 드라이버 | MySqlConnector 2.4.0 (MariaDB 전용) |
| Excel | ClosedXML 0.105.0 |
| 애니메이션 | Avalonia.Labs.Lottie |
| MVVM | CommunityToolkit.Mvvm 8.3.0 (부분 사용) |

---

## 2. 디렉터리 구조

```
ETA/
├── Models/            데이터 전달 객체(DTO) — 로직 없음
├── Services/          모든 DB 접근 + 비즈니스 로직 (static class)
├── ViewModels/        Avalonia MVVM 바인딩 (부분 사용)
├── Views/             최상위 Window/Page
│   └── Pages/         메뉴별 페이지 UserControl
├── Styles/            공용 AXAML 스타일 (DataGridCommonStyles)
├── Assets/            폰트/아이콘/Lottie/동영상
├── Data/              템플릿/Export/Photos
├── appsettings.json   MariaDB 연결 정보 (gitignore 권장)
├── ARCHITECTURE.md    ← 이 파일 (AI 코딩 레퍼런스)
└── MASTERBOOK.md      사용자 매뉴얼
```

---

## 3. DB 패턴 — 핵심 규칙 (MariaDB 전용)

### 3-1. 연결 생성

```csharp
using var conn = DbConnectionFactory.CreateConnection();
conn.Open();
```
반드시 `DbConnectionFactory.CreateConnection()`만 사용. `MySqlConnection`을 직접 new하면 안 된다.

### 3-2. SQL 헬퍼 — 반드시 사용

| 용도 | 하드코딩 금지 | 올바른 사용법 |
|------|-------------|-------------|
| PK 자동증분 DDL | `AUTO_INCREMENT` | `{DbConnectionFactory.AutoIncrement}` |
| 마지막 삽입 ID | `LAST_INSERT_ID()` | `{DbConnectionFactory.LastInsertId}` |
| 행 ID 컬럼 (`_id`) | `_id` 직접 기재 | `{DbConnectionFactory.RowId}` |
| 날짜 포맷 | `DATE_FORMAT(...)` | `DbConnectionFactory.DateFmt(col, fmt)` |
| UPSERT 절 | 직접 작성 | `DbConnectionFactory.UpsertSuffix(conflictCols, updateCols)` |

### 3-3. 스키마 헬퍼

```csharp
DbConnectionFactory.GetColumnNames(conn, "테이블명")   // 컬럼 목록
DbConnectionFactory.ColumnExists(conn, "테이블", "컬럼")
DbConnectionFactory.TableExists(conn, "테이블명")
```

### 3-4. 테이블 구조

모든 테이블에 `_id INT AUTO_INCREMENT PRIMARY KEY` 존재. 행 ID는 반드시 `DbConnectionFactory.RowId`로 참조.

---

## 4. 화면 레이아웃 — ContentControl 슬롯

메인 창(`MainPage`)에는 4개의 `ContentControl` 슬롯이 있다.

| 이름 | AXAML x:Name | 위치/역할 |
|------|-------------|---------|
| Show1 | `Show1` | 좌측 — 트리/메인 페이지 |
| Show2 | `Show2` | 우상단 — 리스트/폼 |
| Show3 | `Show3` | 우하단 — 수정 폼/액션 패널 |
| Show4 | `Show4` | 우측 보조 — 출력 보관함/분석항목 |

> **이전 명칭 주의**: 이전 코드(커밋 전)에서는 `ActivePageContent1~4`였으나  
> 2026-03-29 일괄 리네임으로 `Show1~4`로 변경됨.

### 레이아웃 비율 변경

```csharp
SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 8.5, lowerStar: 1.5);
```
모드 전환 시 각 메뉴 핸들러에서 호출.

---

## 5. 메인 메뉴 — 모드 목록

| 메뉴명 | `_currentMode` | 진입 메서드 | 설명 |
|--------|----------------|------------|------|
| 직원정보 | `"Agent"` | `Agent_Click` | 직원 트리 |
| 계약업체 | `"Contract"` | `Contract_Click` | 계약 + 단가 |
| 견적/의뢰서 | `"Quotation"` | `Quotation_Click` | 4-패널 견적 |
| 시험성적서 | `"TestReport"` | `TestReport_Click` | 분석 결과 트리 |
| 보수요청 | `"Repair"` | `Repair_Click` | 수리 요청 |
| 물품구매 | `"Purchase"` | `Purchase_Click` | 구매 관리 |
| 폐수배출업소 | `"WasteCompany"` | `WasteCompany_Click` | 업소 관리 |
| 일반업무관리 | `"Admin"` | `Admin_Click` | 관리자 기능 |

> **삭제됨**: `측정인` MenuItem — 2026-03-29 메인 메뉴에서 제거.  
> 기능은 시험성적서의 서브메뉴 버튼 "자료TO측정인"으로 이동 예정.

---

## 6. 서브메뉴 버튼 구조 (BT1~BT7)

```csharp
SetSubMenu("BT1레이블", "BT2", "BT3", "BT4", "BT5", "BT6", "BT7");
```
빈 문자열(`""`)이면 해당 버튼을 숨긴다.

### 시험성적서 서브메뉴 (최신)

```csharp
SetSubMenu("새로고침", "CSV 저장", "삭제", "엑셀 출력", "PDF 출력", "일괄 엑셀", "자료TO측정인");
```

| 버튼 | 동작 |
|------|------|
| BT1 새로고침 | `_testReportPage.LoadData()` |
| BT2 CSV 저장 | `_testReportPage.SaveCsv()` |
| BT3 삭제 | `_testReportPage.DeleteSampleAsync()` |
| BT4 엑셀 출력 | `_testReportPage.PrintExcel()` |
| BT5 PDF 출력 | `_testReportPage.PrintPdf()` |
| BT6 일괄 엑셀 | `_testReportPage.BatchPrintExcel()` |
| BT7 자료TO측정인 | `new DataToMeasurerWindow().Show(this)` ← 신규 |

---

## 7. Services 레퍼런스

모든 서비스는 `static class`이며 인스턴스화하지 않는다.

| 파일 | 담당 테이블/기능 |
|------|----------------|
| `DbConnectionFactory` | 연결 팩토리 + SQL 방언 헬퍼 |
| `DbPathHelper` | 경로/연결 문자열 설정 |
| `AgentService` | Agent (직원 정보) |
| `AnalysisService` | 분석정보 (분석항목 메타) |
| `AnalysisRequestService` | 분석의뢰및결과 |
| `AnalysisRecordService` | 분석기록부 출력 |
| `ContractService` | 계약 DB, 분석단가 |
| `ContractPriceService` | 계약 단가 관리 |
| `QuotationService` | 견적발행내역 |
| `TestReportService` | 분석의뢰및결과 (시험성적서) |
| `TestReportPrintServices` | Excel/PDF 출력 |
| `MeasurerService` | 측정인_채취지점, 측정인_분석항목 |
| `OrderRequestService` | 시료명칭 |
| `PurchaseService` | 물품구매 |
| `RepairService` | 보수요청 |
| `WasteCompanyService` | 폐수배출업소 |
| `AdminsSerivces` | 관리자 기능 |
| `CurrentUserManager` | 로그인 세션 (Singleton) |
| `TodoService` | To-Do 연동 (외부 API) |
| `WindowPositionManager` | 창 위치/레이아웃 저장 |
| `BadgeColorHelper` | 초성 배지 색상 |
| `AppInstaller` | 초기 패키지 설치 확인 |

---

## 8. 주요 모델

| 클래스 | 용도 |
|--------|------|
| `SampleRequest` | 시험성적서 시료 1건 (Id, 채취일자, 분석결과 Dictionary 포함) |
| `AnalysisResultRow` | 분석 결과 리스트 1행 (항목명, 결과값, 단위 등) |
| `AnalysisItem` | 분석정보 테이블 행 (ES, Method, instrument 등) |
| `Agent` | 직원 정보 |
| `Contract` | 계약 정보 |
| `QuotationIssue` | 견적 발행 내역 |
| `Quotation` | 견적/의뢰서 |
| `RepairItem` | 보수요청 |
| `PurchaseItem` | 물품구매 |
| `WasteCompany` | 폐수배출업소 |

---

## 9. DB 테이블 목록 (MariaDB 기준 — 16개)

| 테이블명 | 레코드 수(마이그레이션 시점) | 비고 |
|----------|--------------------------|------|
| Agent | 14 | 직원 정보 |
| 견적발행내역 | 934 | QuotationIssue |
| 계약 DB | 32 | Contract |
| 물품구매 | 6 | Purchase |
| 방류기준표 | 114 | 법적 방류 기준 |
| 보수요청 | 3 | Repair |
| 분석결과 | 3544 | 시험성적서 결과 |
| 분석단가 | 67 | 계약 단가 |
| 분석의뢰및결과 | 3711 | TestReport 메인 |
| 분석정보 | 67 | 분석 항목 메타 |
| 분장표준처리 | 390 | 업무 분장 |
| 시료명칭 | 252 | OrderRequest |
| 의뢰정보 | 0 | (미사용) |
| 측정인_분석항목 | 67 | MeasurerService |
| 측정인_채취지점 | 101 | MeasurerService |
| 폐수배출업소 | 166 | WasteCompany |

---

## 10. 알려진 DB 방언 이슈 해결 이력

| 증상 | 원인 | 해결 |
|------|------|------|
| `Unknown column 'rowid'` | MariaDB에 rowid 없음 | `RowId` 헬퍼로 교체 |
| `Unknown column '_id'` | 하드코딩 제거 | `RowId` 헬퍼로 교체 |
| `strftime(...)` 오류 | MariaDB에서 미지원 | `DateFmt()` 헬퍼로 교체 |
| `AUTOINCREMENT` 오류 | MariaDB에서 미지원 | `AutoIncrement` 헬퍼로 교체 |
| 포트 3307 연결 불가 | 방화벽 / 포트 오설정 | appsettings.json Port → 3306 |
| `분석완료일자` 컬럼 없음 | 한자 석 vs 한글 석 혼용 | SQL에서 `분석완료일자`(한글)로 통일 |
| `처리내용` 컬럼 없음 | RepairService 오타 `정리내용` | `처리내용`으로 수정 |
| GetSamplesByCompany string interpolation | `` `{c}` `` 에 `$` 누락 | `` $"`{c}`" `` 로 수정 |

---

## 11. 중요 변경 이력 (역순)
### 2026-03-29 (후속)
- **측정인 로그인창 (`MeasurerLoginWindow`)** 개선
  - 계약 DB / 분析DB 업데이트 완료 후 **로그인 창 자동 닫기** (2초 딜레이)
  - 스크래핑 진행 중 **웹페이지 상단 고정 JS 프로그레스바** 주입 (`__eta_pb__`)
    - `WebProgress_InjectAsync(title, total, color)` — 바 삽입
    - `WebProgress_UpdateAsync(current, total, msg)` — 진행률 업데이트
    - `WebProgress_RemoveAsync()` — 완료 후 제거
  - `PollForSyncRequestAsync` 전면 재작성: WebSocket 끊김 감지, 각 작업 별도 try-catch, 오류 시 페이지 알람 + 버튼 복원
  - 세션 시작 시 `측정인.log`에 구분선·타임스탬프 헤더 기록 (`LogSessionStart()`)
  - 로그 포맷: `[HH:mm:ss.fff]` 밀리초 + `Debug.WriteLine` 동시 출력
- **`InstallCheckDialog.CheckAndShowAsync()`**: `Required=false` 패키지(LibreOffice 등) 설치 확인 스킵
- **서브메뉴 `BT8` 추가** (`측정인 LOGIN`)
  - `MainPage.axaml`: BT8 버튼 추가
  - `SetSubMenu()`: `bt8` 선택 파라미터 추가
  - `BT8_Click`: TestReport 모드에서 `MeasurerLoginWindow` 호출
- **`DataToMeasurerWindow` 리팩터링**
  - `SampleRequest? selectedSample` 생성자 파라미터 추가
  - Show1(TestReportPage)에서 현재 선택된 노드의 시료(`SelectedSample`)만 로드
  - 시료 미선택 시 경고 메시지 표시
  - 타이틀에 선택된 약칭/시료명 표시
- **`TestReportPage`**: `public SampleRequest? SelectedSample => _selectedSample;` 공개 프로퍼티 추가
- **`MainPage.axaml.cs` BT7_Click**: `new DataToMeasurerWindow(_testReportPage?.SelectedSample)` 로 변경
### 2026-03-29
- `ActivePageContent1~4` → `Show1~4` 전체 리네임 (sed 일괄 치환)
- 메인 메뉴에서 `측정인` MenuItem 제거
- 시험성적서 BT7: `"측정인 LOGIN"` → `"자료TO측정인"` (DataToMeasurerWindow 호출)
- ContractPage `ActivePageContent3`(구명) 분석단가 테이블 표시 코드 제거

### 2026-03-27~28
- `DbConnectionFactory.RowId` 속성 추가 (`"_id"`)
- 전체 서비스 `_id` 하드코딩 → `{DbConnectionFactory.RowId}` 교체
  - 대상: QuotationService, Testreportservice, AnalysisRequestService, OrderRequestService, AnalysisService
- SQLite→MariaDB 전체 재마이그레이션 — 모든 테이블에 `_id AUTO_INCREMENT` 추가
- `Login.axaml` PasswordChar 제거 (비밀번호 항상 표시)
- `분석완료일자` 한자(석) → 한글(석) 수정
- RepairService `정리내용` → `처리내용` 오타 수정
- appsettings.json Port 3307 → 3306

### 2026-03-26 이전
- MariaDB 이중 DB 모드 도입 (`DbConnectionFactory.UseMariaDb`)
- 시험성적서 Excel 불러오기 기능
- 견적/의뢰서 4-패널 구조
- 측정인.kr CDP 자동 로그인 (`MeasurerLoginWindow`)

---

## 12. 코딩 컨벤션

- **static Services**: `using var conn = DbConnectionFactory.CreateConnection(); conn.Open();` 패턴
- **백틱 식별자**: 한글 컬럼/테이블명은 항상 백틱으로 감쌈 `` `컬럼명` ``
- **파라미터**: 반드시 `@param` 형식 (`cmd.Parameters.AddWithValue`)
- **Lazy 초기화**: `_page ??= new XxxPage()` 패턴으로 첫 진입 시 생성
- **FontFamily**: `private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");`
- **UI 색상**: 다크 테마 기준 `#1e1e2e`(배경), `#2d2d2d`(패널), `#444`(구분선), `#8888bb`(보조 텍스트)
- **로그**: `Debug.WriteLine($"[ServiceName] ...")` 패턴 사용
