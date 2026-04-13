# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ETA is a Laboratory Information Management System (LIMS) built with .NET 8.0 and Avalonia UI. It manages water quality analysis, wastewater monitoring, and facility testing for environmental laboratories.

## Build & Development Commands

```bash
# Build the project
dotnet build ETA.sln

# Run the application
dotnet run --project ETA.csproj

# Clean build artifacts
dotnet clean ETA.sln

# Restore dependencies
dotnet restore ETA.sln
```

## Architecture Overview

### Database Architecture (MariaDB)
The system uses MariaDB through a unified connection factory:

```csharp
// Always use this - never instantiate MySqlConnection directly
using var conn = DbConnectionFactory.CreateConnection();
conn.Open();
```

**Critical SQL Dialect Helpers** (must use these for consistency):
- Auto-increment: `DbConnectionFactory.AutoIncrement` 
- Last insert ID: `DbConnectionFactory.LastInsertId`
- Row ID column: `DbConnectionFactory.RowId` ("_id")
- Date formatting: `DbConnectionFactory.DateFmt(column, format)`
- UPSERT operations: `DbConnectionFactory.UpsertSuffix(conflictCols, updateCols)`

### UI Layout System
The main window (`MainPage`) uses 4 content control slots:
- **Show1**: Left panel (tree view, main navigation)
- **Show2**: Top-right panel (lists, forms)
- **Show3**: Bottom-right panel (edit forms, action panels)
- **Show4**: Right auxiliary panel (output, analysis items)

Layout ratios can be adjusted per mode:
```csharp
SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 8.5, lowerStar: 1.5);
```

### Service Architecture
All business logic is in static service classes under `/Services/`:
- **Common/**: Infrastructure (DbConnectionFactory, AppFonts, WindowPositionManager)
- **SERVICE1/**: Water analysis center services (QuotationService, AnalysisRequestService)
- **SERVICE2/**: Wastewater & facility services (WasteSampleService, FacilityResultService)
- **SERVICE4/**: UI helpers (ManualMatchWindow)

### Project Structure (v11.0.0)
```
ETA/
├── Services/
│   ├── Common/          # Infrastructure services
│   ├── SERVICE1/        # Water analysis center (수질분석센터)
│   ├── SERVICE2/        # Wastewater management (비용부담금/처리시설)
│   └── SERVICE4/        # UI utilities
├── Views/
│   ├── Pages/
│   │   ├── PAGE1/       # Water analysis center pages
│   │   ├── PAGE2/       # Wastewater management pages
│   │   └── Common/      # Shared pages (RepairPage, PurchasePage)
│   ├── MainPage.axaml   # Main application window
│   └── Login.axaml      # Login window
├── Models/              # Data transfer objects (DTOs)
├── Data/               # Templates + exported files + Photos
└── Docs/               # Documentation including ARCHITECTURE.md
```

## Key Business Domains

### 1. Water Analysis Center (수질분석센터)
- **Quote Creation**: `QuotationService` manages quotations with dynamic analysis item columns
- **Sample Requests**: `AnalysisRequestService` handles analysis requests linked to quotations
- **Test Reports**: Multi-format output (Excel/PDF) via `TestReportService`

### 2. Wastewater Management (비용부담금)
- **Company Management**: `WasteCompanyService` manages wastewater discharge companies
- **Sample Tracking**: `WasteSampleService` handles sample collection and results
- **Analysis Input**: Unified input system in `WasteAnalysisInputPage` with 3-tab structure

### 3. Treatment Facilities (처리시설)
- **Analysis Planning**: Schedule-based automatic work item generation
- **Facility Results**: `FacilityResultService` manages facility measurement data
- **Work Assignment**: Auto-generates daily work items based on facility schedules

## Critical Coding Patterns

### Database Access
```csharp
// Correct pattern for all DB operations
using var conn = DbConnectionFactory.CreateConnection();
conn.Open();
using var cmd = conn.CreateCommand();
cmd.CommandText = $"SELECT {DbConnectionFactory.RowId}, name FROM `table_name`";
// Always use backticks for Korean table/column names
```

### UI Content Updates
```csharp
// Update content slots
Show1.Content = newPanel;
ListPanelChanged?.Invoke(newPanel);  // For Show2
EditPanelChanged?.Invoke(newPanel);  // For Show3
StatsPanelChanged?.Invoke(newPanel); // For Show4 (Show1 in some modes)
```

### Font and Styling
```csharp
// Font helpers for consistent sizing
FsXS(textBlock);  // Extra small
FsSM(textBlock);  // Small  
FsMD(textBlock);  // Medium
FsLG(textBlock);  // Large
```

## Data Flow Patterns

### Analysis Result Input (분석결과입력)
The system uses a 2-tier save architecture:
1. **Raw measurement data** → `*_DATA` tables (BOD_DATA, TOC_DATA, etc.)
2. **Calculated results** → Request tables (분석의뢰및결과, 폐수의뢰및결과, 처리시설_측정결과)

### Auto-classification System
Sample data is automatically classified by source:
- **수질분석센터**: Matches by 약칭/시료명 → saves to `분석의뢰및결과`
- **비용부담금**: Matches by SN → saves to `폐수의뢰및결과`
- **처리시설**: Matches by 시료명 → saves to `처리시설_측정결과`

## Configuration

### Database Connection
Edit `appsettings.json` for MariaDB connection:
```json
{
  "MariaDb": {
    "Server": "your-server",
    "Port": "3306",
    "Database": "eta_db",
    "User": "eta_user",
    "Password": "your-password"
  }
}
```

## Important Notes

- **한자 금지**: Never use Chinese characters (析, 등) in any text - always use Korean only
- **Static Services**: All service classes are static - never instantiate them
- **Lazy Initialization**: UI pages use `_page ??= new XxxPage()` pattern
- **Error Handling**: All DB operations should have try-catch with Debug.WriteLine logging
- **Korean Column Names**: Always use backticks around Korean table/column names in SQL

## Memory System
The project uses Claude's memory system in `.claude/memory/` for persistent context about user preferences, coding patterns, and project structure.