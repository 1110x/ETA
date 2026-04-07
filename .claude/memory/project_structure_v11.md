---
name: ETA 프로젝트 폴더 구조 (v11.0.0)
description: PAGE1/PAGE2/Common 페이지 그룹, SERVICE1/SERVICE2/Common 서비스 그룹, Docs/Logs 정리
type: project
---

v11.0.0 기준 폴더 구조로 전면 재편됨.

**Why:** 메뉴별 도메인 분리 + 앞으로 PAGE3/SERVICE3 식으로 확장 관리하기 위해

**How to apply:** 새 페이지/서비스 추가 시 반드시 이 구조에 맞는 폴더에 배치할 것

### Views/Pages/
- `PAGE1/` — 수질분석센터 메뉴 호출 페이지 (네임스페이스: ETA.Views.Pages.PAGE1)
- `PAGE2/` — 비용부담금 관리 메뉴 호출 페이지 (네임스페이스: ETA.Views.Pages.PAGE2)
- `Common/` — 공통 페이지: RepairPage, PurchasePage (네임스페이스: ETA.Views.Pages.Common)

### Services/
- `SERVICE1/` — 수질분석센터 서비스 (ETA.Services.SERVICE1)
- `SERVICE2/` — 비용부담금 관리 서비스 (ETA.Services.SERVICE2)
- `Common/` — 공통 인프라: DbConnectionFactory, DbPathHelper, AppFonts, WindowPositionManager 등 (ETA.Services.Common)

### 기타
- `Docs/` — MD 파일 전체
- `Logs/` — LOG 파일 전체 (코드 내 경로 참조도 Logs/ 접두어로 업데이트됨)

### using 패턴 (MainPage.axaml.cs 등 호출부)
```csharp
using ETA.Views.Pages;
using ETA.Views.Pages.PAGE1;
using ETA.Views.Pages.PAGE2;
using ETA.Views.Pages.Common;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
```
