using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE2;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views;

/// <summary>처리시설 엑셀 임포트 확인/수정 창</summary>
public class FacilityImportWindow : Window
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static IBrush Res(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is IBrush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 결과 이벤트 ──────────────────────────────────────────────────────────
    public event Action<List<Dictionary<string, string>>>? ImportConfirmed;

    private readonly List<Dictionary<string, string>>                    _rows;
    private readonly List<(string 시설명, string 시료명, int 마스터Id)>  _masters;

    // 매칭 결과: rowIndex → masterEntry
    private readonly Dictionary<int, (string 시설명, string 시료명, int 마스터Id)> _matched = new();

    private StackPanel _gridPanel = new() { Spacing = 0 };

    // ── 컬럼 정의 ────────────────────────────────────────────────────────────
    private static readonly string[] KnownKeys =
        { "BOD", "TOC", "SS", "TN", "TP", "COD", "총대장균군", "염소이온", "영양염류", "함수율", "중금속", "비고" };

    public FacilityImportWindow(
        string filePath,
        List<Dictionary<string, string>> rows,
        List<(string 시설명, string 시료명, int 마스터Id)> masters)
    {
        _rows    = rows;
        _masters = masters;

        Title                 = $"처리시설 임포트 — {System.IO.Path.GetFileName(filePath)}";
        Width                 = 900;
        Height                = 560;
        MinWidth              = 600;
        MinHeight             = 400;
        CanResize             = true;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        SystemDecorations     = SystemDecorations.Full;
        Background            = AppTheme.BgPrimary;
        FontFamily            = Font;

        // 자동 매칭
        for (int i = 0; i < _rows.Count; i++)
        {
            var name = _rows[i].TryGetValue("시료명", out var n) ? n : "";
            var m = FacilityResultService.FindBySampleName(_masters, name);
            if (m != null)
                _matched[i] = _masters.First(x => x.마스터Id == m.Value.마스터Id);
        }

        BuildUI();
    }

    private void BuildUI()
    {
        var root = new DockPanel { Margin = new Thickness(10) };

        // ── 헤더 ─────────────────────────────────────────────────────────────
        var header = new Border
        {
            Background    = Res("GridHeaderBg", "#252535"),
            CornerRadius  = new CornerRadius(6),
            Padding       = new Thickness(10, 8),
            Margin        = new Thickness(0, 0, 0, 8),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 16,
                Children =
                {
                    new TextBlock
                    {
                        Text = $"📋 {_rows.Count}행 파싱됨  |  매칭: {_matched.Count}/{_rows.Count}",
                        FontFamily = Font, FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
                        Foreground = AppTheme.FgPrimary, VerticalAlignment = VerticalAlignment.Center,
                    },
                }
            }
        };
        DockPanel.SetDock(header, Dock.Top);
        root.Children.Add(header);

        // ── 하단 버튼 ─────────────────────────────────────────────────────────
        var btnBar = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 8,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 0, 0),
        };
        var btnConfirm = MakeBtn("✅ 저장", "#1a3a1a", "#88ff88");
        var btnCancel  = MakeBtn("닫기",   "#2a2a3a", "#aaaacc");
        btnConfirm.Click += (_, _) =>
        {
            ImportConfirmed?.Invoke(_rows);
            Close();
        };
        btnCancel.Click += (_, _) => Close();
        btnBar.Children.Add(btnConfirm);
        btnBar.Children.Add(btnCancel);
        DockPanel.SetDock(btnBar, Dock.Bottom);
        root.Children.Add(btnBar);

        // ── 그리드 ────────────────────────────────────────────────────────────
        RefreshGrid();
        var scroll = new ScrollViewer
        {
            Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        root.Children.Add(scroll);

        Content = root;
    }

    private void RefreshGrid()
    {
        _gridPanel.Children.Clear();

        // 표시할 컬럼 결정 (파일에 있는 KnownKeys만)
        var activeCols = KnownKeys.Where(k => _rows.Any(r => r.ContainsKey(k))).ToArray();
        // colDefs: 매칭(30) + 시료명(140) + 시설명(120) + 항목들(60씩)
        string colDefs = "30,140,120," + string.Join(",", activeCols.Select(_ => "65"));

        // 헤더
        var hGrid = new Grid { ColumnDefinitions = new ColumnDefinitions(colDefs), MinHeight = 28,
            Background = Res("GridHeaderBg", "#252535") };
        int col = 0;
        AddHdrCell(hGrid, col++, "");
        AddHdrCell(hGrid, col++, "시료명");
        AddHdrCell(hGrid, col++, "매칭 시설/시료");
        foreach (var k in activeCols) AddHdrCell(hGrid, col++, k);
        _gridPanel.Children.Add(new Border
        {
            Child = hGrid, BorderBrush = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        });

        // 데이터 행
        for (int i = 0; i < _rows.Count; i++)
        {
            var row      = _rows[i];
            int captured = i;
            bool matched = _matched.TryGetValue(i, out var m);

            var rGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDefs),
                MinHeight = 30,
                Background = i % 2 == 0 ? Res("GridRowBg", "#1e1e28") : Res("GridRowAltBg", "#23233a"),
            };

            // 매칭 아이콘
            AddCell(rGrid, 0, matched ? "🟢" : "🔴", matched ? AppTheme.FgSuccess : AppTheme.FgDanger);

            // 시료명
            var sampleName = row.TryGetValue("시료명", out var sn) ? sn : "";
            AddCell(rGrid, 1, sampleName, AppTheme.FgPrimary);

            // 매칭 시설명 + 클릭으로 선택
            var matchLabel = matched ? $"{m.시설명} / {m.시료명}" : "— 클릭해서 매칭 —";
            var matchTb = new TextBlock
            {
                Text = matchLabel, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = matched ? AppTheme.FgInfo : AppTheme.FgMuted,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(4, 0),
                Cursor = new Cursor(StandardCursorType.Hand),
            };
            Grid.SetColumn(matchTb, 2); rGrid.Children.Add(matchTb);

            // 항목 값
            int ac = 3;
            foreach (var k in activeCols)
            {
                var val = row.TryGetValue(k, out var v) ? v : "";
                AddCell(rGrid, ac++, val, string.IsNullOrEmpty(val) ? AppTheme.FgMuted : AppTheme.FgSecondary);
            }

            var border = new Border
            {
                Child = rGrid, Cursor = new Cursor(StandardCursorType.Hand),
                BorderBrush = Res("ThemeBorderSubtle", "#333344"), BorderThickness = new Thickness(0, 0, 0, 1),
            };
            border.PointerPressed += (_, _) => OpenMasterPicker(captured);
            TextShimmer.AttachHover(border);
            _gridPanel.Children.Add(border);
        }
    }

    // ── 처리시설 마스터 선택 팝업 (ManualMatchWindow 재활용) ────────────────
    private void OpenMasterPicker(int rowIndex)
    {
        var row      = _rows[rowIndex];
        var sampleNm = row.TryGetValue("시료명", out var sn) ? sn : "";
        var result   = row.TryGetValue("BOD", out var v) ? v : "";

        var win = new ManualMatchWindow(sampleNm, result,
            analysisRecords: new(),
            facilityMasters: _masters,
            wasteSamples: new());

        win.MatchConfirmed += w =>
        {
            if (w.SelectedFacility != null)
            {
                _matched[rowIndex] = w.SelectedFacility.Value;

                // 기존 시료명 → 비고, 픽한 시설 시료명으로 교체
                var r = _rows[rowIndex];
                var originalName = r.TryGetValue("시료명", out var on) ? on : "";
                if (!string.IsNullOrEmpty(originalName))
                    r["비고"] = originalName;
                r["시료명"] = w.SelectedFacility.Value.시료명;

                Avalonia.Threading.Dispatcher.UIThread.Post(RefreshGrid);
            }
        };
        win.Show();
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────
    private static Button MakeBtn(string text, string bg, string fg) => new()
    {
        Content = text, FontFamily = Font, FontSize = AppTheme.FontBase,
        Background = new SolidColorBrush(Color.Parse(bg)),
        Foreground = new SolidColorBrush(Color.Parse(fg)),
        BorderThickness = new Thickness(0), Padding = new Thickness(14, 6),
        CornerRadius = new CornerRadius(4),
    };

    private static void AddHdrCell(Grid g, int col, string text)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontSM,
            FontWeight = FontWeight.SemiBold, Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = col >= 3 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
            Margin = new Thickness(4, 0),
        };
        Grid.SetColumn(tb, col); g.Children.Add(tb);
    }

    private static void AddCell(Grid g, int col, string text, IBrush fg)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = fg, VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = col >= 3 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(4, 0),
        };
        Grid.SetColumn(tb, col); g.Children.Add(tb);
    }
}
