using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE4;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 분석조건 입력 페이지 (동시분석 지원).
/// 좌측: 분석항목 체크리스트(다중선택) — 체크된 항목 전부가 저장 대상
/// 중앙: 프라이머리 항목의 Key/Value 편집 — 저장 시 체크된 모든 항목에 덮어쓰기
/// </summary>
public class AnalysisConditionPage
{
    public Control LeftPanel   { get; }
    public Control CenterPanel { get; }

    private readonly StackPanel _analyteRows = new() { Spacing = 0 };
    private readonly StackPanel _kvRows      = new() { Spacing = 4 };
    private readonly TextBlock  _title       = new()
    {
        FontSize   = AppTheme.FontLG,
        FontWeight = FontWeight.Bold,
        Foreground = AppTheme.FgPrimary,
        Margin     = new Thickness(0, 0, 0, 8),
        Text       = "분석조건 입력",
    };
    private readonly TextBlock _selectionStatus = new()
    {
        FontSize   = AppTheme.FontSM,
        Foreground = AppTheme.FgMuted,
        Margin     = new Thickness(8, 2, 8, 6),
        Text       = "선택 0개",
    };

    private readonly HashSet<string> _checked = new(StringComparer.Ordinal);
    private string? _primaryAnalyte;
    private readonly Dictionary<string, (CheckBox chk, Border row)> _rowMap = new(StringComparer.Ordinal);

    public AnalysisConditionPage()
    {
        LeftPanel   = BuildLeft();
        CenterPanel = BuildCenter();
        RefreshAnalyteList();
    }

    // ── Left: Analyte 체크리스트 ───────────────────────────────────────
    private Control BuildLeft()
    {
        var header = new TextBlock
        {
            Text       = "분석항목",
            FontSize   = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgSecondary,
            Margin     = new Thickness(8, 8, 8, 4),
        };

        var btnAll = new Button
        {
            Content  = "모두 선택",
            FontSize = AppTheme.FontSM,
            Padding  = new Thickness(8, 2),
        };
        btnAll.Click += (_, _) => SetAllChecked(true);

        var btnNone = new Button
        {
            Content  = "모두 해제",
            FontSize = AppTheme.FontSM,
            Padding  = new Thickness(8, 2),
        };
        btnNone.Click += (_, _) => SetAllChecked(false);

        var toolbar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 4,
            Margin      = new Thickness(8, 0, 8, 4),
            Children    = { btnAll, btnNone },
        };

        return new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
            Background     = AppTheme.BgSecondary,
            Children =
            {
                header,
                new Grid { Children = { toolbar }, [Grid.RowProperty] = 1 },
                new Grid { Children = { _selectionStatus }, [Grid.RowProperty] = 2 },
                new ScrollViewer
                {
                    Content = _analyteRows,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    [Grid.RowProperty] = 3,
                },
            },
        };
    }

    private void RefreshAnalyteList()
    {
        try
        {
            var items = AnalysisConditionService.GetAnalyteList();
            _analyteRows.Children.Clear();
            _rowMap.Clear();
            _checked.Clear();

            foreach (var name in items)
            {
                var captured = name;
                var chk = new CheckBox
                {
                    IsChecked         = false,
                    Margin            = new Thickness(8, 0, 4, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                chk.IsCheckedChanged += (_, _) =>
                {
                    if (chk.IsChecked == true) _checked.Add(captured);
                    else                       _checked.Remove(captured);
                    UpdateSelectionStatus();
                };

                var label = new TextBlock
                {
                    Text              = name,
                    FontSize          = AppTheme.FontBase,
                    Foreground        = AppTheme.FgPrimary,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin            = new Thickness(0, 4, 8, 4),
                };

                var grid = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("Auto,*"),
                    Background        = Brushes.Transparent,
                    Cursor            = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                };
                Grid.SetColumn(chk,   0);
                Grid.SetColumn(label, 1);
                grid.Children.Add(chk);
                grid.Children.Add(label);

                var row = new Border
                {
                    Padding       = new Thickness(0),
                    Background    = Brushes.Transparent,
                    BorderBrush   = Brushes.Transparent,
                    BorderThickness = new Thickness(0, 0, 0, 0),
                    Child         = grid,
                };
                // 라벨 영역 클릭 → 프라이머리로 지정(조건 로드). 체크박스 클릭은 체크 토글만.
                label.PointerPressed += (_, _) => SetPrimary(captured);
                // 빈 영역 클릭도 프라이머리 지정
                row.PointerPressed += (_, e) =>
                {
                    // 체크박스를 직접 눌렀을 땐 건드리지 않음 (CheckBox 자체 이벤트가 처리)
                    if (e.Source is CheckBox) return;
                    SetPrimary(captured);
                };

                _analyteRows.Children.Add(row);
                _rowMap[name] = (chk, row);
            }

            if (items.Count > 0) SetPrimary(items[0]);
            UpdateSelectionStatus();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisConditionPage] 목록 로드 실패: {ex.Message}");
        }
    }

    private void SetAllChecked(bool on)
    {
        foreach (var (_, v) in _rowMap) v.chk.IsChecked = on;
    }

    private void SetPrimary(string analyte)
    {
        _primaryAnalyte = analyte;
        foreach (var kv in _rowMap)
        {
            kv.Value.row.Background = (kv.Key == analyte)
                ? AppTheme.BgActiveBlue
                : Brushes.Transparent;
        }
        LoadConditions();
    }

    private void UpdateSelectionStatus()
    {
        var primary = string.IsNullOrEmpty(_primaryAnalyte) ? "(없음)" : _primaryAnalyte;
        _selectionStatus.Text = $"선택 {_checked.Count}개  ·  프라이머리: {primary}";
    }

    // ── Center: Key/Value 편집 ────────────────────────────────────────
    private Control BuildCenter()
    {
        var btnAdd = new Button
        {
            Content  = "➕ 행 추가",
            FontSize = AppTheme.FontBase,
            Padding  = new Thickness(10, 4),
            Margin   = new Thickness(0, 0, 6, 0),
        };
        btnAdd.Click += (_, _) => AddKvRow("", "");

        var btnLoad = new Button
        {
            Content  = "↻ 프라이머리로 불러오기",
            FontSize = AppTheme.FontBase,
            Padding  = new Thickness(10, 4),
            Margin   = new Thickness(0, 0, 6, 0),
        };
        btnLoad.Click += (_, _) => LoadConditions();

        var btnSave = new Button
        {
            Content    = "💾 선택 항목 모두에 저장",
            FontSize   = AppTheme.FontBase,
            Padding    = new Thickness(10, 4),
            Background = AppTheme.StatusOkBg,
            Foreground = AppTheme.StatusOkFg,
        };
        btnSave.Click += OnSave;

        var toolbar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 4,
            Margin      = new Thickness(0, 0, 0, 8),
            Children    = { btnAdd, btnLoad, btnSave },
        };

        var kvHeader = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("180,*,40"),
            Margin            = new Thickness(0, 0, 0, 2),
            Children =
            {
                HeaderCell("항목(Key)", 0),
                HeaderCell("값(Value)", 1),
                HeaderCell("",          2),
            },
        };

        var scroller = new ScrollViewer
        {
            Content                       = _kvRows,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };

        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
            Margin         = new Thickness(16),
        };
        Grid.SetRow(_title,   0);
        Grid.SetRow(toolbar,  1);
        Grid.SetRow(kvHeader, 2);
        Grid.SetRow(scroller, 3);
        root.Children.Add(_title);
        root.Children.Add(toolbar);
        root.Children.Add(kvHeader);
        root.Children.Add(scroller);

        return new Border
        {
            Background = AppTheme.BgPrimary,
            Child      = root,
        };
    }

    private TextBlock HeaderCell(string text, int col)
    {
        var tb = new TextBlock
        {
            Text       = text,
            FontSize   = AppTheme.FontSM,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            Margin     = new Thickness(4, 0, 0, 0),
        };
        Grid.SetColumn(tb, col);
        return tb;
    }

    private void LoadConditions()
    {
        _kvRows.Children.Clear();
        _title.Text = string.IsNullOrEmpty(_primaryAnalyte)
            ? "분석조건 입력"
            : $"분석조건 입력 — {_primaryAnalyte} (프라이머리)";
        UpdateSelectionStatus();

        if (string.IsNullOrEmpty(_primaryAnalyte)) return;
        try
        {
            var list = AnalysisConditionService.Load(_primaryAnalyte);
            foreach (var c in list) AddKvRow(c.Key, c.Value);
            if (list.Count == 0) AddKvRow("", "");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisConditionPage] 조건 로드 실패: {ex.Message}");
        }
    }

    private void AddKvRow(string key, string value)
    {
        var txKey = new TextBox
        {
            Text      = key,
            FontSize  = AppTheme.FontBase,
            Watermark = "예: 오븐온도",
            Margin    = new Thickness(0, 0, 4, 0),
        };
        var txVal = new TextBox
        {
            Text      = value,
            FontSize  = AppTheme.FontBase,
            Watermark = "예: 40 °C",
            Margin    = new Thickness(0, 0, 4, 0),
        };
        var btnDel = new Button
        {
            Content    = "✕",
            FontSize   = AppTheme.FontSM,
            Padding    = new Thickness(6, 2),
            Background = Brushes.Transparent,
            Foreground = AppTheme.FgDanger,
        };
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("180,*,40"),
            Margin            = new Thickness(0, 1, 0, 1),
        };
        Grid.SetColumn(txKey,  0);
        Grid.SetColumn(txVal,  1);
        Grid.SetColumn(btnDel, 2);
        row.Children.Add(txKey);
        row.Children.Add(txVal);
        row.Children.Add(btnDel);
        btnDel.Click += (_, _) => _kvRows.Children.Remove(row);
        _kvRows.Children.Add(row);
    }

    private void OnSave(object? sender, RoutedEventArgs e)
    {
        // 저장 대상 = 체크된 항목(프라이머리가 체크 안 돼 있어도 체크된 것만 저장)
        if (_checked.Count == 0)
        {
            _title.Text = $"분석조건 입력 — ✗ 저장 대상이 없습니다 (좌측 체크박스를 선택하세요)";
            return;
        }

        var list = new List<AnalysisCondition>();
        foreach (var child in _kvRows.Children)
        {
            if (child is not Grid row) continue;
            var k = (row.Children.OfType<TextBox>().ElementAtOrDefault(0)?.Text ?? "").Trim();
            var v = (row.Children.OfType<TextBox>().ElementAtOrDefault(1)?.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(k)) continue;
            list.Add(new AnalysisCondition { Key = k, Value = v });
        }

        int ok = 0, fail = 0;
        var firstError = "";
        foreach (var a in _checked)
        {
            try   { AnalysisConditionService.Save(a, list); ok++; }
            catch (Exception ex)
            {
                fail++;
                if (string.IsNullOrEmpty(firstError)) firstError = ex.Message;
            }
        }

        _title.Text = fail == 0
            ? $"분석조건 입력 — ✓ {ok}개 항목 저장 완료 (프라이머리: {_primaryAnalyte})"
            : $"분석조건 입력 — ⚠ {ok}개 성공 / {fail}개 실패: {firstError}";
    }
}
