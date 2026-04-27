using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE1;
using ETA.Services.Common;
using ETA.Views;
using SkiaSharp;
using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// *_시험기록부 테이블 뷰어.
/// Show1=테이블 목록, Show4=선택 테이블의 분석일 목록, Show2=선택 날짜의 행 데이터.
/// </summary>
public class TestRecordBookViewerPage : UserControl
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    public Control TableListControl { get; }
    public Control DateListControl  { get; }
    public Control RowsGridControl  { get; }

    public event Action<Control>? DateListChanged;
    public event Action<Control>? RowsGridChanged;
    public event Action<Control>? CompanyListChanged;  // Show3 — 선택 날짜의 업체 목록

    private readonly TreeView _tableTree;
    private readonly TextBox _tableSearch;
    private readonly TextBlock _tableStatus;
    private List<string> _allTables = new();
    private Dictionary<string, (string Category, string ES, string 약칭)> _tableMeta = new();

    private readonly TreeView _dateTree;
    private readonly TextBlock _dateStatus;

    private readonly Border _rowsHost;
    private readonly TextBlock _rowsStatus;

    private string? _selectedTable;
    private string? _selectedDate;
    private string? _selectedCompany;   // 선택된 업체명 필터 ("" 또는 null = 전체)

    public TestRecordBookViewerPage()
    {
        // ── Show1: 테이블 목록 ────────────────────────────────────────
        _tableSearch = new TextBox
        {
            Watermark = "🔍 테이블 검색...",
            Height = 24, FontSize = AppTheme.FontSM, FontFamily = Font,
            Margin = new Thickness(6, 6, 6, 2),
        };
        _tableSearch.TextChanged += (_, _) => ApplyTableFilter();

        _tableStatus = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin = new Thickness(8, 2, 8, 2),
            Text = "",
        };

        _tableTree = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Margin = new Thickness(4),
            SelectionMode = SelectionMode.Single,
        };
        _tableTree.SelectionChanged += TableTree_SelectionChanged;

        var leftGrid = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(_tableSearch, 0); leftGrid.Children.Add(_tableSearch);
        Grid.SetRow(_tableStatus, 1); leftGrid.Children.Add(_tableStatus);
        Grid.SetRow(_tableTree, 2);   leftGrid.Children.Add(_tableTree);
        TableListControl = leftGrid;

        // ── Show4: 분석일 목록 ───────────────────────────────────────
        _dateStatus = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin = new Thickness(8, 6, 8, 2),
            Text = "테이블을 선택하세요",
        };
        _dateTree = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Margin = new Thickness(4),
            SelectionMode = SelectionMode.Single,
        };
        _dateTree.SelectionChanged += DateTree_SelectionChanged;

        var dateGrid = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        Grid.SetRow(_dateStatus, 0); dateGrid.Children.Add(_dateStatus);
        Grid.SetRow(_dateTree, 1);   dateGrid.Children.Add(_dateTree);
        DateListControl = dateGrid;

        // ── Show2: 행 그리드 + 인쇄 버튼 ─────────────────────────────
        _rowsStatus = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 8, 12, 4),
            Text = "날짜를 선택하세요",
        };
        // Excel 출력은 계획에서 제외 — 코드는 보존하되 버튼만 숨김
        var printBtn = new Button
        {
            Content = "🖨 Excel",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Padding = new Thickness(10, 4),
            Margin = new Thickness(8, 4, 4, 4),
            Background = new SolidColorBrush(Color.Parse("#1a3a5c")),
            Foreground = new SolidColorBrush(Color.Parse("#90caf9")),
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
            IsVisible = false,
        };
        printBtn.Click += async (_, _) => await PrintShow2Async();

        var wordBtn = new Button
        {
            Content = "📝 Word",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Padding = new Thickness(10, 4),
            Margin = new Thickness(8, 4, 12, 4),
            Background = new SolidColorBrush(Color.Parse("#2a3e5c")),
            Foreground = new SolidColorBrush(Color.Parse("#aac7e8")),
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
        };
        wordBtn.Click += async (_, _) => await ExportWordAsync();

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal };
        btnRow.Children.Add(printBtn);
        btnRow.Children.Add(wordBtn);

        var statusRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        Grid.SetColumn(_rowsStatus, 0); statusRow.Children.Add(_rowsStatus);
        Grid.SetColumn(btnRow, 1);     statusRow.Children.Add(btnRow);

        _rowsHost = new Border
        {
            Padding = new Thickness(8),
            Child = new TextBlock
            {
                Text = "내용 없음",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgDimmed,
            }
        };
        var rowsGrid = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        Grid.SetRow(statusRow, 0); rowsGrid.Children.Add(statusRow);
        Grid.SetRow(_rowsHost, 1); rowsGrid.Children.Add(_rowsHost);
        RowsGridControl = rowsGrid;
    }

    public void LoadTables()
    {
        _allTables = TestRecordBookViewerService.GetAllTables();
        _tableMeta = TestRecordBookViewerService.GetAnalyteMetaForTables(_allTables);
        _tableStatus.Text = $"총 {_allTables.Count}개";
        ApplyTableFilter();
    }

    private void ApplyTableFilter()
    {
        string kw = (_tableSearch.Text ?? "").Trim();
        IEnumerable<string> filtered = _allTables;
        if (kw.Length > 0)
            filtered = _allTables.Where(t =>
                t.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                TestRecordBookViewerService.PrettyName(t).Contains(kw, StringComparison.OrdinalIgnoreCase));

        // Category 그룹핑 + ES 순 정렬 (카테고리는 각 그룹의 최소 ES 로 정렬)
        var grouped = filtered
            .Select(t => new { Table = t, Meta = _tableMeta.TryGetValue(t, out var m) ? m : ("기타", "", "") })
            .GroupBy(x => x.Meta.Item1)
            .OrderBy(g => CategorySortKey(g.Key))
            .ThenBy(g => g.Min(x => EsSortKey(x.Meta.Item2)), StringComparer.Ordinal)
            .ThenBy(g => g.Key, StringComparer.CurrentCulture)
            .ToList();

        _tableTree.Items.Clear();
        foreach (var g in grouped)
        {
            var catNode = new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text = $"📂 {g.Key}  ({g.Count()})",
                    FontFamily = Font, FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgMuted,
                    Padding = new Thickness(4, 2),
                },
                IsExpanded = true,
            };
            foreach (var x in g.OrderBy(x => EsSortKey(x.Meta.Item2)).ThenBy(x => x.Table))
            {
                var leaf = new TreeViewItem
                {
                    Tag = x.Table,
                    Header = MakeLeafHeader(x.Table, x.Meta.Item2, x.Meta.Item3),
                };
                catNode.Items.Add(leaf);
            }
            _tableTree.Items.Add(catNode);
        }
    }

    private static int CategorySortKey(string cat)
    {
        // "기타" 는 맨 뒤
        if (string.IsNullOrWhiteSpace(cat) || cat == "기타") return 999;
        return 1;
    }

    // ES "ES 04302.1b" / "04302.1b" 등에서 숫자 부분만 추출해 정렬용 zero-padded 키 반환
    private static string EsSortKey(string es)
    {
        if (string.IsNullOrWhiteSpace(es)) return "zzzz_" + (es ?? "");
        var m = System.Text.RegularExpressions.Regex.Match(es, @"(\d{4,5})(?:\.(\d+))?");
        if (!m.Success) return "zzz_" + es;
        var main   = m.Groups[1].Value.PadLeft(5, '0');
        var subRaw = m.Groups[2].Success ? m.Groups[2].Value : "";
        var sub    = subRaw.PadLeft(3, '0');
        return $"{main}.{sub}_{es}";
    }

    private Control MakeLeafHeader(string table, string es, string 약칭)
    {
        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };

        // 약칭 뱃지 (좌)
        if (!string.IsNullOrWhiteSpace(약칭))
        {
            var (bg, fg) = ETA.Services.Common.BadgeColorHelper.GetBadgeColor(약칭);
            var abbrBadge = new Border
            {
                Background = Brush.Parse(bg),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1),
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = 약칭, FontFamily = Font, FontSize = AppTheme.FontXS,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = Brush.Parse(fg),
                }
            };
            Grid.SetColumn(abbrBadge, 0);
            row.Children.Add(abbrBadge);
        }

        // 이름
        var nameTxt = new TextBlock
        {
            Text = TestRecordBookViewerService.PrettyName(table),
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            Padding = new Thickness(4, 2),
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(nameTxt, 1);
        row.Children.Add(nameTxt);

        // ES 뱃지 (우)
        if (!string.IsNullOrWhiteSpace(es))
        {
            var esBadge = new Border
            {
                Background = Brush.Parse("#2a3a5a"),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1),
                Margin = new Thickness(4, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = es, FontFamily = Font, FontSize = AppTheme.FontXS,
                    Foreground = Brush.Parse("#a8c5e8"),
                }
            };
            Grid.SetColumn(esBadge, 2);
            row.Children.Add(esBadge);
        }
        return row;
    }

    private void TableTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_tableTree.SelectedItem is not TreeViewItem item || item.Tag is not string table) return;
        _selectedTable = table;
        _selectedDate = null;
        LoadDates();
        DateListChanged?.Invoke(DateListControl);
    }

    private void LoadDates()
    {
        _dateTree.Items.Clear();
        if (string.IsNullOrEmpty(_selectedTable))
        {
            _dateStatus.Text = "테이블을 선택하세요";
            return;
        }

        var dates = TestRecordBookViewerService.GetDates(_selectedTable);
        _dateStatus.Text = $"{TestRecordBookViewerService.PrettyName(_selectedTable)} — {dates.Count}일";

        // 월별 그룹 (최근월만 Expanded)
        var grouped = dates
            .Select(d => new { Raw = d, Month = MonthKey(d) })
            .GroupBy(x => x.Month)
            .OrderByDescending(g => g.Key, StringComparer.Ordinal)
            .ToList();

        bool isFirst = true;
        foreach (var g in grouped)
        {
            var monthNode = new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text = $"📅 {PrettyMonth(g.Key)}  ({g.Count()}일)",
                    FontFamily = Font, FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgMuted,
                    Padding = new Thickness(4, 2),
                },
                IsExpanded = isFirst,
            };
            isFirst = false;
            foreach (var x in g.OrderByDescending(x => x.Raw, StringComparer.Ordinal))
            {
                monthNode.Items.Add(new TreeViewItem
                {
                    Tag = x.Raw,
                    Header = new TextBlock
                    {
                        Text = x.Raw, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgPrimary, Padding = new Thickness(6, 2),
                    }
                });
            }
            _dateTree.Items.Add(monthNode);
        }
    }

    private static string MonthKey(string date)
    {
        if (date.Length >= 7 && (date[4] == '-' || date[4] == '/')) return date[..7].Replace('/', '-');
        if (date.Length >= 6 && int.TryParse(date[..4], out _))      return date[..4] + "-" + date.Substring(4, 2);
        return date;
    }

    private static string PrettyMonth(string key)
    {
        // "2026-04" → "2026년 04월"
        if (key.Length >= 7 && key[4] == '-') return $"{key[..4]}년 {key.Substring(5, 2)}월";
        return key;
    }

    private void DateTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_dateTree.SelectedItem is not TreeViewItem item || item.Tag is not string date) return;
        if (string.IsNullOrEmpty(_selectedTable)) return;
        _selectedDate = date;
        LoadRows();
        RowsGridChanged?.Invoke(RowsGridControl);
    }

    private void LoadRows()
    {
        if (string.IsNullOrEmpty(_selectedTable) || string.IsNullOrEmpty(_selectedDate)) return;

        var (_, rows) = TestRecordBookViewerService.GetRowsByDate(_selectedTable, _selectedDate);
        _rowsStatus.Text = $"{TestRecordBookViewerService.PrettyName(_selectedTable)} · {_selectedDate} — {rows.Count}건";

        // Show3: 업체 목록 갱신 (날짜 변경 시 새로 빌드)
        BuildCompanyListPanel(rows);

        if (rows.Count == 0)
        {
            _rowsHost.Child = new TextBlock
            {
                Text = "해당 날짜의 행이 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(16),
            };
            return;
        }

        var abbr = _selectedTable != null && _tableMeta.TryGetValue(_selectedTable, out var mm) ? mm.약칭 : "";
        var model = TestRecordBookViewerService.BuildParsedModel(_selectedTable, _selectedDate, abbr, _selectedCompany);
        _rowsHost.Child = TestRecordBookParsedView.Build(model);
    }

    /// <summary>Show3 — 선택 날짜의 distinct 업체명 리스트. 클릭 시 Show2 를 그 업체만 필터.</summary>
    private void BuildCompanyListPanel(List<Dictionary<string, string>> rows)
    {
        var root = new StackPanel { Spacing = 4, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock
        {
            Text = "🏢 업체 선택",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 6),
        });

        var companies = rows
            .Select(r => r.TryGetValue("업체명", out var v) ? (v ?? "").Trim() : "")
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.Ordinal)
            .ToList();

        // "전체" 항목 추가
        Border MakeItem(string label, string? value, int count)
        {
            var stack = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            stack.Children.Add(new TextBlock
            {
                Text = label, FontSize = AppTheme.FontSM,
                FontWeight = (value == _selectedCompany) ? FontWeight.SemiBold : FontWeight.Regular,
                Foreground = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
            });
            stack.Children.Add(new TextBlock
            {
                Text = $"({count})", FontSize = AppTheme.FontXS,
                Foreground = AppTheme.FgDimmed,
                VerticalAlignment = VerticalAlignment.Center,
            });
            var bgActive  = new SolidColorBrush(Color.Parse("#2a4a6a"));
            var bgIdle    = new SolidColorBrush(Color.Parse("#1f1f28"));
            var bgHover   = new SolidColorBrush(Color.Parse("#2a2a35"));
            var border = new Border
            {
                Padding = new Thickness(8, 6),
                Margin = new Thickness(0, 0, 0, 3),
                Background = (value == _selectedCompany) ? bgActive : bgIdle,
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(StandardCursorType.Hand),
                Child = stack,
            };
            border.PointerEntered += (_, _) => { if (value != _selectedCompany) border.Background = bgHover; };
            border.PointerExited  += (_, _) => { if (value != _selectedCompany) border.Background = bgIdle; };
            border.PointerPressed += (_, _) =>
            {
                _selectedCompany = value;
                LoadRows();
                RowsGridChanged?.Invoke(RowsGridControl);
            };
            return border;
        }

        root.Children.Add(MakeItem("전체", null, rows.Count));
        foreach (var co in companies)
        {
            int cnt = rows.Count(r => r.TryGetValue("업체명", out var v)
                && string.Equals((v ?? "").Trim(), co, StringComparison.OrdinalIgnoreCase));
            root.Children.Add(MakeItem(co, co, cnt));
        }

        CompanyListChanged?.Invoke(new ScrollViewer { Content = root });
    }

    /// <summary>현재 선택된 (테이블, 날짜) 를 MainPage 에서 접근.</summary>
    public string? SelectedTable => _selectedTable;
    public string? SelectedDate  => _selectedDate;

    /// <summary>Show2 (시험기록부 미리보기) 를 비트맵 → PDF → 시스템 인쇄.
    ///   1) 컨트롤 측정·렌더 (실제 컨텐트 크기로 RenderTargetBitmap)
    ///   2) PNG 인코딩 → SkiaSharp SKDocument 로 단일 페이지 PDF 합성
    ///   3) 시스템 print 명령 호출 (macOS: lp, Windows: print)
    ///   벳지·테두리·색상·글자 모두 비트맵으로 보존됨.</summary>
    public async System.Threading.Tasks.Task PrintShow2Async()
    {
        try
        {
            if (string.IsNullOrEmpty(_selectedTable) || string.IsNullOrEmpty(_selectedDate))
            {
                _rowsStatus.Text = "테이블/날짜 선택 후 인쇄하세요";
                return;
            }

            // Show2 의 섹션 구조를 그대로 Excel 로 출력 — 시료별 alt-row 에 계산식 풀어쓰기
            var abbr = _tableMeta.TryGetValue(_selectedTable, out var mm) ? mm.약칭 : "";
            var printModel = TestRecordBookViewerService.BuildParsedModel(_selectedTable, _selectedDate, abbr, _selectedCompany);
            var xlsxPath = await System.Threading.Tasks.Task.Run(() =>
                TestRecordBookExcelExporter.Export(printModel));

            // 시스템 기본 앱(Excel/Numbers)으로 열기 — 사용자가 거기서 인쇄
            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    if (System.OperatingSystem.IsMacOS())
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo { FileName = "open", UseShellExecute = false };
                        psi.ArgumentList.Add(xlsxPath);
                        System.Diagnostics.Process.Start(psi);
                    }
                    else if (System.OperatingSystem.IsWindows())
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = xlsxPath, UseShellExecute = true,
                        });
                    }
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Excel Open] {ex.Message}");
                }
            });

            _rowsStatus.Text = $"📄 Excel 출력 완료: {System.IO.Path.GetFileName(xlsxPath)}";
        }
        catch (System.Exception ex)
        {
            _rowsStatus.Text = $"인쇄 오류: {ex.Message}";
        }
    }

    /// <summary>Show2 모델을 Word(.docx) 로 출력 — 프로토타입.</summary>
    public async System.Threading.Tasks.Task ExportWordAsync()
    {
        try
        {
            if (string.IsNullOrEmpty(_selectedTable) || string.IsNullOrEmpty(_selectedDate))
            {
                _rowsStatus.Text = "테이블/날짜 선택 후 출력하세요";
                return;
            }
            var abbr = _tableMeta.TryGetValue(_selectedTable, out var mm) ? mm.약칭 : "";
            var model = TestRecordBookViewerService.BuildParsedModel(_selectedTable, _selectedDate, abbr, _selectedCompany);
            var docxPath = await System.Threading.Tasks.Task.Run(() =>
                TestRecordBookWordExporter.Export(model));

            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    if (System.OperatingSystem.IsMacOS())
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo { FileName = "open", UseShellExecute = false };
                        psi.ArgumentList.Add(docxPath);
                        System.Diagnostics.Process.Start(psi);
                    }
                    else if (System.OperatingSystem.IsWindows())
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = docxPath, UseShellExecute = true,
                        });
                    }
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Word Open] {ex.Message}");
                }
            });

            _rowsStatus.Text = $"📝 Word 출력 완료: {System.IO.Path.GetFileName(docxPath)}";
        }
        catch (System.Exception ex)
        {
            _rowsStatus.Text = $"Word 출력 오류: {ex.Message}";
        }
    }
}
