using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE2;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 처리시설 항목별 결과값 테이블 패널 (Show3).
/// 시설 선택 시 기간 내 시료별 날짜/항목 매트릭스를 표로 표시.
/// 데이터 소스는 추이 패널과 동일한 FacilityResultService.GetRowsInRange.
/// </summary>
public class ProcessingResultTablePanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly int[] Periods = { 30, 60, 90, 120, 150, 180 };

    private readonly TextBlock _tbTitle;
    private readonly StackPanel _tableHost;
    private string? _facility;
    private int _periodDays = 30;

    public ProcessingResultTablePanel()
    {
        _tbTitle = new TextBlock
        {
            Text       = "처리시설을 선택하세요",
            FontFamily = Font,
            FontSize   = AppTheme.FontLG,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            Margin     = new Thickness(10, 8, 10, 4),
        };

        var topBar = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            Margin      = new Thickness(10, 2, 10, 4),
            VerticalAlignment = VerticalAlignment.Center,
        };

        var periodCombo = new ComboBox
        {
            FontSize   = AppTheme.FontSM,
            FontFamily = Font,
            Padding    = new Thickness(6, 2),
            MinWidth   = 70,
            ItemsSource  = Periods,
            SelectedItem = _periodDays,
            VerticalAlignment = VerticalAlignment.Center,
        };
        periodCombo.SelectionChanged += (_, _) =>
        {
            if (periodCombo.SelectedItem is int v)
            {
                _periodDays = v;
                if (_facility != null) LoadAndRender();
            }
        };
        topBar.Children.Add(new TextBlock
        {
            Text       = "기간",
            FontFamily = Font,
            FontSize   = AppTheme.FontSM,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin     = new Thickness(0, 0, 4, 0),
        });
        topBar.Children.Add(periodCombo);

        _tableHost = new StackPanel { Spacing = 10, Margin = new Thickness(10, 0, 10, 10) };
        var scroll = new ScrollViewer
        {
            Content = _tableHost,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        var root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,Auto,*") };
        Grid.SetRow(_tbTitle, 0); root.Children.Add(_tbTitle);
        Grid.SetRow(topBar,   1); root.Children.Add(topBar);
        Grid.SetRow(scroll,   2); root.Children.Add(scroll);
        Content = root;
    }

    public void SetFacility(string facility)
    {
        _facility = facility;
        LoadAndRender();
    }

    private void LoadAndRender()
    {
        _tableHost.Children.Clear();
        if (string.IsNullOrEmpty(_facility)) return;

        var to   = DateTime.Today;
        var from = to.AddDays(-_periodDays + 1);
        _tbTitle.Text = $"{_facility} · {from:MM-dd} ~ {to:MM-dd} 항목별 결과";
        _tbTitle.Foreground = AppTheme.FgPrimary;

        Dictionary<string, List<(string Date, Dictionary<string, string> Values)>> data;
        try
        {
            data = FacilityResultService.GetRowsInRange(
                _facility!, from.ToString("yyyy-MM-dd"), to.ToString("yyyy-MM-dd"));
        }
        catch (Exception ex)
        {
            _tableHost.Children.Add(new TextBlock
            {
                Text       = $"조회 오류: {ex.Message}",
                Foreground = AppTheme.FgDanger,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
            });
            return;
        }

        if (data.Count == 0)
        {
            _tableHost.Children.Add(new TextBlock
            {
                Text       = "기간 내 데이터가 없습니다.",
                Foreground = AppTheme.FgMuted,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
                Margin     = new Thickness(0, 20, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            return;
        }

        var itemOrder = FacilityResultService.GetAnalysisItems(activeOnly: false)
            .Select(i => i.컬럼명.Trim('`'))
            .ToList();
        var dataKeys = data.Values
            .SelectMany(rows => rows.SelectMany(r => r.Values.Keys))
            .Distinct()
            .ToHashSet();
        var keys = itemOrder.Where(k => dataKeys.Contains(k)).ToList();
        foreach (var k in dataKeys) if (!keys.Contains(k)) keys.Add(k);

        foreach (var (sample, rows) in data)
            _tableHost.Children.Add(BuildSampleTable(sample, rows, keys));
    }

    private static Border BuildSampleTable(
        string sample,
        List<(string Date, Dictionary<string, string> Values)> rows,
        List<string> keys)
    {
        var header = new TextBlock
        {
            Text       = $"{sample}  ·  {rows.Count}건",
            FontFamily = Font,
            FontSize   = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
            Margin     = new Thickness(0, 0, 0, 4),
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition(100, GridUnitType.Pixel));
        foreach (var _ in keys)
            grid.ColumnDefinitions.Add(new ColumnDefinition(80, GridUnitType.Pixel));

        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        foreach (var _ in rows)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        AddHeader(grid, 0, 0, "날짜");
        for (int i = 0; i < keys.Count; i++) AddHeader(grid, 0, i + 1, keys[i]);

        // 최신일이 위로 오도록 내림차순
        var sortedRows = rows.OrderByDescending(r => r.Date).ToList();
        for (int r = 0; r < sortedRows.Count; r++)
        {
            bool even = r % 2 == 0;
            var (date, vmap) = sortedRows[r];
            string shortDate = date.Length >= 10 ? date.Substring(0, 10) : date;
            AddCell(grid, r + 1, 0, shortDate, even, isDate: true);
            for (int c = 0; c < keys.Count; c++)
            {
                string v = vmap.TryGetValue(keys[c], out var s) ? s : "";
                AddCell(grid, r + 1, c + 1, v, even, isDate: false);
            }
        }

        var stack = new StackPanel { Spacing = 2 };
        stack.Children.Add(header);
        stack.Children.Add(grid);

        return new Border
        {
            Background      = AppTheme.BgCard,
            BorderBrush     = AppTheme.BorderMuted,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(10, 8),
            Child           = stack,
        };
    }

    private static void AddHeader(Grid grid, int row, int col, string text)
    {
        var cell = new Border
        {
            Background      = AppTheme.BgCard,
            BorderBrush     = AppTheme.BorderMuted,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Padding         = new Thickness(4, 4),
            Child = new TextBlock
            {
                Text       = text,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgSecondary,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
        Grid.SetRow(cell, row);
        Grid.SetColumn(cell, col);
        grid.Children.Add(cell);
    }

    private static void AddCell(Grid grid, int row, int col, string text, bool even, bool isDate)
    {
        string bg = even ? "#35343c" : "#2d2d35";
        var cell = new Border
        {
            Background      = Brush.Parse(bg),
            BorderBrush     = AppTheme.BorderMuted,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Padding         = new Thickness(4, 3),
            MinHeight       = 22,
            Child = new TextBlock
            {
                Text       = text,
                FontFamily = Font,
                FontSize   = AppTheme.FontSM,
                Foreground = isDate ? AppTheme.FgSecondary : AppTheme.FgPrimary,
                HorizontalAlignment = isDate ? HorizontalAlignment.Left : HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
        Grid.SetRow(cell, row);
        Grid.SetColumn(cell, col);
        grid.Children.Add(cell);
    }
}
