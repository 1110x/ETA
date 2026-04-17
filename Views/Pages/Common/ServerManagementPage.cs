using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Services.Common;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 서버관리 — DB 테이블 목록 / 엑셀 내려받기 / 엑셀 업로드
/// </summary>
public sealed class ServerManagementPage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    public Control LeftPanel   { get; }   // Show1: 테이블 목록
    public Control CenterPanel { get; }   // Show2: 테이블 미리보기

    private ListBox?   _tableListBox;
    private TextBlock? _statusLabel;
    private DataGrid?  _previewGrid;
    private TextBlock? _previewTitle;
    private List<string> _previewColumns = new();
    private List<object?[]> _previewRows = new();
    private string? _pkColName;
    private int     _pkColIdx = -1;

    public ServerManagementPage()
    {
        LeftPanel   = BuildTableListPanel();
        CenterPanel = BuildCenterPanel();
    }

    public void Reload() => LoadTableList();

    // ═══════════════════════════════════════════════════════════════════
    // Show1: 테이블 목록
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildTableListPanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*,Auto") };

        // 헤더
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "🗄  서버 테이블 목록",
                FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppTheme.FgMuted,
            }
        });

        // 상태 라벨
        _statusLabel = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin = new Thickness(10, 4),
        };
        Grid.SetRow(_statusLabel, 1);
        root.Children.Add(_statusLabel);

        // 테이블 리스트
        _tableListBox = new ListBox
        {
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Margin = new Thickness(5, 0, 5, 5),
            Background = Brush.Parse("#1e2030"),
        };
        _tableListBox.SelectionChanged += (_, _) => LoadPreview();
        Grid.SetRow(_tableListBox, 2);
        root.Children.Add(_tableListBox);

        // 버튼 영역
        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 4,
            Margin = new Thickness(5),
        };
        var btnDownload = MakeButton("⬇ 엑셀로 다운로드", "#1a3a1a", "#7cd87c");
        btnDownload.Click += OnDownload;
        var btnUpload = MakeButton("⬆ 엑셀 업로드", "#1a2a3a", "#7aaae8");
        btnUpload.Click += OnUpload;
        var btnRefresh = MakeButton("🔄 새로고침", "#2a2a2a", "#aaaaaa");
        btnRefresh.Click += (_, _) => Reload();
        btnPanel.Children.Add(btnDownload);
        btnPanel.Children.Add(btnUpload);
        btnPanel.Children.Add(btnRefresh);
        Grid.SetRow(btnPanel, 3);
        root.Children.Add(btnPanel);

        LoadTableList();
        return root;
    }

    private static Button MakeButton(string text, string bg, string fg) => new()
    {
        Content = text,
        FontFamily = Font, FontSize = AppTheme.FontSM,
        Background = new SolidColorBrush(Color.Parse(bg)),
        Foreground = new SolidColorBrush(Color.Parse(fg)),
        BorderThickness = new Thickness(0),
        Padding = new Thickness(10, 4),
        CornerRadius = new CornerRadius(4),
    };

    private void LoadTableList()
    {
        if (_tableListBox == null) return;
        _tableListBox.Items.Clear();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SHOW TABLES";
            using var r = cmd.ExecuteReader();
            var names = new List<string>();
            while (r.Read())
            {
                var n = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                if (!string.IsNullOrEmpty(n)) names.Add(n);
            }
            names.Sort(StringComparer.CurrentCulture);
            foreach (var n in names) _tableListBox.Items.Add(n);
            if (_statusLabel != null) _statusLabel.Text = $"총 {names.Count}개 테이블";
        }
        catch (Exception ex)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgDanger;
                _statusLabel.Text = $"오류: {ex.Message}";
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show2: 선택된 테이블 미리보기 (최대 100행)
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildCenterPanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        _previewTitle = new TextBlock
        {
            Text = "테이블을 선택하면 미리보기가 표시됩니다",
            FontSize = AppTheme.FontBase, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            Margin = new Thickness(10, 6),
        };
        root.Children.Add(_previewTitle);

        _previewGrid = new DataGrid
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Margin = new Thickness(5),
            AutoGenerateColumns = true,
            IsReadOnly = false,  // 편집 가능
            HeadersVisibility = DataGridHeadersVisibility.Column,
            GridLinesVisibility = DataGridGridLinesVisibility.All,
        };
        _previewGrid.CellEditEnded += OnCellEditEnded;
        Grid.SetRow(_previewGrid, 1);
        root.Children.Add(_previewGrid);
        return root;
    }

    private void LoadPreview()
    {
        if (_tableListBox?.SelectedItem is not string tableName) return;
        if (_previewTitle != null) _previewTitle.Text = $"📋 {tableName} (최대 100행 미리보기)";
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `{tableName}` LIMIT 100";
            using var r = cmd.ExecuteReader();

            _previewColumns = new List<string>();
            for (int i = 0; i < r.FieldCount; i++) _previewColumns.Add(r.GetName(i));

            // PK 컬럼 인덱스 찾기
            _pkColName = null; _pkColIdx = -1;
            foreach (var key in new[] { "_id", "id", "Id", "ID" })
            {
                var idx = _previewColumns.FindIndex(c =>
                    string.Equals(c, key, StringComparison.OrdinalIgnoreCase));
                if (idx >= 0) { _pkColName = _previewColumns[idx]; _pkColIdx = idx; break; }
            }

            _previewRows = new List<object?[]>();
            while (r.Read())
            {
                var row = new object?[r.FieldCount];
                for (int i = 0; i < r.FieldCount; i++)
                    row[i] = r.IsDBNull(i) ? null : r.GetValue(i);
                _previewRows.Add(row);
            }

            if (_previewGrid != null)
            {
                _previewGrid.Columns.Clear();
                for (int i = 0; i < _previewColumns.Count; i++)
                {
                    var colName = _previewColumns[i];
                    _previewGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = colName,
                        Binding = new Avalonia.Data.Binding($"[{i}]"),
                        IsReadOnly = (i == _pkColIdx),  // PK는 편집 불가
                    });
                }
                _previewGrid.ItemsSource = _previewRows;
            }
        }
        catch (Exception ex)
        {
            if (_previewTitle != null)
            {
                _previewTitle.Foreground = AppTheme.FgDanger;
                _previewTitle.Text = $"미리보기 오류: {ex.Message}";
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // 다운로드 - 선택된 테이블 전체 → Excel
    // ═══════════════════════════════════════════════════════════════════
    private async void OnDownload(object? sender, RoutedEventArgs e)
    {
        if (_tableListBox?.SelectedItem is not string tableName)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgWarn;
                _statusLabel.Text = "테이블을 먼저 선택하세요";
            }
            return;
        }

        try
        {
            var top = TopLevel.GetTopLevel(_tableListBox);
            if (top == null) return;
            var saveFile = await top.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                SuggestedFileName = $"{tableName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                DefaultExtension = "xlsx",
                FileTypeChoices = new[] { new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx" } } },
            });
            if (saveFile == null) return;

            var outPath = saveFile.Path.LocalPath;
            await Task.Run(() => ExportTableToExcel(tableName, outPath));

            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgSuccess;
                _statusLabel.Text = $"✅ {Path.GetFileName(outPath)} 저장";
            }
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true }); }
            catch { }
        }
        catch (Exception ex)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgDanger;
                _statusLabel.Text = $"❌ {ex.Message}";
            }
        }
    }

    private static void ExportTableToExcel(string tableName, string outPath)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM `{tableName}`";
        using var r = cmd.ExecuteReader();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add(tableName.Length > 31 ? tableName.Substring(0, 31) : tableName);

        // 헤더
        for (int i = 0; i < r.FieldCount; i++)
            ws.Cell(1, i + 1).Value = r.GetName(i);

        int rowIdx = 2;
        while (r.Read())
        {
            for (int i = 0; i < r.FieldCount; i++)
            {
                if (r.IsDBNull(i)) continue;
                var val = r.GetValue(i);
                ws.Cell(rowIdx, i + 1).Value = XLCellValue.FromObject(val);
            }
            rowIdx++;
        }
        ws.Columns().AdjustToContents();
        wb.SaveAs(outPath);
    }

    // ═══════════════════════════════════════════════════════════════════
    // 업로드 - Excel → 선택된 테이블로 INSERT
    // ═══════════════════════════════════════════════════════════════════
    private async void OnUpload(object? sender, RoutedEventArgs e)
    {
        if (_tableListBox?.SelectedItem is not string tableName)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgWarn;
                _statusLabel.Text = "업로드할 대상 테이블을 선택하세요";
            }
            return;
        }

        try
        {
            var top = TopLevel.GetTopLevel(_tableListBox);
            if (top == null) return;
            var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                AllowMultiple = false,
                FileTypeFilter = new[] { new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx" } } },
            });
            if (files == null || files.Count == 0) return;

            var filePath = files[0].Path.LocalPath;
            var result = await Task.Run(() => ImportExcelToTable(tableName, filePath));
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = result.err == 0 ? AppTheme.FgSuccess : AppTheme.FgWarn;
                _statusLabel.Text = $"✅ {result.ok}건 업로드 (오류 {result.err}건)";
            }
            LoadPreview();
        }
        catch (Exception ex)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgDanger;
                _statusLabel.Text = $"❌ 업로드 오류: {ex.Message}";
            }
        }
    }

    private static (int ok, int err) ImportExcelToTable(string tableName, string filePath)
    {
        int ok = 0, err = 0;
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet(1);
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        if (lastCol == 0 || lastRow < 2) return (0, 0);

        // 1행 헤더 → 컬럼명
        var cols = new List<string>();
        for (int c = 1; c <= lastCol; c++)
        {
            var h = ws.Cell(1, c).GetString()?.Trim();
            if (!string.IsNullOrEmpty(h)) cols.Add(h);
            else cols.Add("");
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        // 대상 테이블 컬럼 확인 (누락된 엑셀 컬럼은 스킵)
        var tableCols = new HashSet<string>(DbConnectionFactory.GetColumnNames(conn, tableName),
            StringComparer.OrdinalIgnoreCase);
        var colIndices = new List<(int excelCol, string dbCol)>();
        for (int i = 0; i < cols.Count; i++)
        {
            if (!string.IsNullOrEmpty(cols[i]) && tableCols.Contains(cols[i]))
                colIndices.Add((i + 1, cols[i]));
        }
        if (colIndices.Count == 0) return (0, 0);

        var colList = string.Join(", ", colIndices.Select(c => $"`{c.dbCol}`"));
        var placeholders = string.Join(",", colIndices.Select(_ => "%"));
        var paramNames = colIndices.Select((_, idx) => $"@p{idx}").ToList();
        var paramPart = string.Join(",", paramNames);

        for (int r = 2; r <= lastRow; r++)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"INSERT INTO `{tableName}` ({colList}) VALUES ({paramPart})";
                for (int i = 0; i < colIndices.Count; i++)
                {
                    var cell = ws.Cell(r, colIndices[i].excelCol);
                    object? val = cell.IsEmpty() ? (object)DBNull.Value : cell.Value.ToString();
                    var p = cmd.CreateParameter();
                    p.ParameterName = paramNames[i];
                    p.Value = val ?? DBNull.Value;
                    cmd.Parameters.Add(p);
                }
                cmd.ExecuteNonQuery();
                ok++;
            }
            catch { err++; }
        }
        return (ok, err);
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show2 셀 편집 → DB UPDATE
    // ═══════════════════════════════════════════════════════════════════
    private void OnCellEditEnded(object? sender, DataGridCellEditEndedEventArgs e)
    {
        if (_tableListBox?.SelectedItem is not string tableName) return;
        if (e.Row?.DataContext is not object?[] row) return;

        // 컬럼 인덱스 (DataGrid DisplayIndex 사용)
        var colIdx = e.Column?.DisplayIndex ?? -1;
        if (colIdx < 0 || colIdx >= _previewColumns.Count) return;
        var columnName = _previewColumns[colIdx];
        var newVal = row[colIdx];

        if (_pkColName == null || _pkColIdx < 0)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgWarn;
                _statusLabel.Text = "PK 컬럼(_id/id)이 없어 편집 저장 불가";
            }
            return;
        }
        var pkVal = row[_pkColIdx];

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `{tableName}` SET `{columnName}`=@val WHERE `{_pkColName}`=@pk";
            cmd.Parameters.AddWithValue("@val", newVal ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@pk", pkVal ?? (object)DBNull.Value);
            var n = cmd.ExecuteNonQuery();
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = n > 0 ? AppTheme.FgSuccess : AppTheme.FgWarn;
                _statusLabel.Text = n > 0 ? $"✅ {columnName} 저장" : "변경 없음";
            }
        }
        catch (Exception ex)
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgDanger;
                _statusLabel.Text = $"❌ 저장 오류: {ex.Message}";
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // DB 백업 - 모든 테이블 → 단일 SQL 덤프 파일
    // ═══════════════════════════════════════════════════════════════════
    public static void BackupDatabase(string outPath)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8);
        sw.WriteLine($"-- ETA DB Backup {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sw.WriteLine("SET NAMES utf8mb4;");
        sw.WriteLine();

        // 모든 테이블 조회
        var tables = new List<string>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SHOW TABLES";
            using var r = cmd.ExecuteReader();
            while (r.Read()) tables.Add(r.GetString(0));
        }

        foreach (var tbl in tables)
        {
            // CREATE TABLE
            sw.WriteLine($"DROP TABLE IF EXISTS `{tbl}`;");
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SHOW CREATE TABLE `{tbl}`";
                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    var createSql = r.GetString(1);
                    sw.WriteLine(createSql + ";");
                }
            }
            sw.WriteLine();

            // INSERT 데이터
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM `{tbl}`";
                using var r = cmd.ExecuteReader();
                var cols = new List<string>();
                for (int i = 0; i < r.FieldCount; i++) cols.Add(r.GetName(i));
                var colList = string.Join(",", cols.Select(c => $"`{c}`"));

                while (r.Read())
                {
                    var vals = new List<string>();
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        if (r.IsDBNull(i)) vals.Add("NULL");
                        else
                        {
                            var v = r.GetValue(i);
                            if (v is string s) vals.Add($"'{s.Replace("\\", "\\\\").Replace("'", "\\'")}'");
                            else if (v is DateTime dt) vals.Add($"'{dt:yyyy-MM-dd HH:mm:ss}'");
                            else if (v is bool b) vals.Add(b ? "1" : "0");
                            else vals.Add(v.ToString() ?? "NULL");
                        }
                    }
                    sw.WriteLine($"INSERT INTO `{tbl}` ({colList}) VALUES ({string.Join(",", vals)});");
                }
            }
            sw.WriteLine();
        }
    }

    private static readonly Avalonia.Media.FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");
}
