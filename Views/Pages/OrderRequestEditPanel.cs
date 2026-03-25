using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

/// <summary>
/// 의뢰서 편집 패널 — ActivePageContent2 에 표시
/// 1열: 시료명 / 2열~: 분析항목 / 셀: 체크박스 (견적항목 기본 체크)
/// </summary>
public class OrderRequestEditPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    private QuotationIssue    _issue       = null!;
    private List<string>      _samples     = new();
    private List<string>      _analysisCols = new();
    private HashSet<string>   _quotedItems = new();

    // [시료명][항목명] → CheckBox
    private readonly Dictionary<string, Dictionary<string, CheckBox>> _cbGrid = new();

    public event Action? SubmitCompleted;
    public event Action? Cancelled;

    public OrderRequestEditPanel() { }

    // ── 외부 호출 ─────────────────────────────────────────────────────────
    public void Load(QuotationIssue issue, List<string> samples, HashSet<string> quotedItems)
    {
        _issue        = issue;
        _samples      = samples;
        _quotedItems  = quotedItems;
        _analysisCols = OrderRequestService.GetAnalysisColumns();
        _cbGrid.Clear();
        Content = BuildUI();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  UI 빌드
    // ══════════════════════════════════════════════════════════════════════
    private Control BuildUI()
    {
        // 헤더
        var titleBar = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
        titleBar.Children.Add(new StackPanel
        {
            Spacing = 2,
            Children =
            {
                new TextBlock
                {
                    Text = "📋  의뢰서 편집",
                    FontSize = 13, FontWeight = FontWeight.Bold,
                    FontFamily = Font, Foreground = Brush.Parse("#e0e0e0"),
                },
                new TextBlock
                {
                    Text = $"{_issue.업체명}  |  {_issue.견적번호}  |  {_issue.견적구분}",
                    FontSize = 9, FontFamily = Font, Foreground = Brush.Parse("#666688"),
                },
            }
        });

        var btnSubmit = new Button
        {
            Content = "✅ 제출", Height = 26, FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#1a4a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 0), Margin = new Thickness(0, 0, 6, 0),
        };
        var btnCancel = new Button
        {
            Content = "← 뒤로", Height = 26, FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#888"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 0),
        };
        btnSubmit.Click += async (_, _) => await OnSubmit();
        btnCancel.Click += (_, _) => Cancelled?.Invoke();
        Grid.SetColumn(btnSubmit, 1);
        Grid.SetColumn(btnCancel, 2);
        titleBar.Children.Add(btnSubmit);
        titleBar.Children.Add(btnCancel);

        // 안내 텍스트
        var info = new TextBlock
        {
            Text = $"시료 {_samples.Count}건  ·  견적항목 {_quotedItems.Count}개 기본 체크  ·  체크 해제로 제외 가능",
            FontSize = 9, FontFamily = Font, Foreground = Brush.Parse("#555577"),
            Margin = new Thickness(0, 4, 0, 6),
        };

        // 그리드 (스크롤)
        var gridPanel = BuildCheckGrid();
        var scroll = new ScrollViewer
        {
            Content = gridPanel,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
        };

        return new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
            Margin = new Thickness(8),
            Children =
            {
                Rc(titleBar, 0),
                Rc(new Border { Height = 1, Background = Brush.Parse("#333"), Margin = new Thickness(0,5,0,5) }, 1),
                Rc(info, 2),
                Rc(scroll, 3),
            }
        };
    }

    private Panel BuildCheckGrid()
    {
        var grid = new Grid();

        // 컬럼 정의
        var colDefs = new ColumnDefinitions();
        colDefs.Add(new ColumnDefinition { Width = new GridLength(170) }); // 시료명
        foreach (var _ in _analysisCols)
            colDefs.Add(new ColumnDefinition { Width = new GridLength(54) });
        grid.ColumnDefinitions = colDefs;

        // 행 정의
        var rowDefs = new RowDefinitions();
        rowDefs.Add(new RowDefinition { Height = new GridLength(40) }); // 헤더
        foreach (var _ in _samples)
            rowDefs.Add(new RowDefinition { Height = new GridLength(28) });
        grid.RowDefinitions = rowDefs;

        // 헤더: 시료명
        grid.Children.Add(HCell("시료명", 0, 0));

        // 헤더: 분析항목
        for (int ci = 0; ci < _analysisCols.Count; ci++)
        {
            var col  = _analysisCols[ci];
            bool isQ = _quotedItems.Contains(col);
            var tb = new TextBlock
            {
                Text              = col,
                FontSize          = 8,
                FontFamily        = Font,
                Foreground        = isQ ? Brush.Parse("#88cc88") : Brush.Parse("#555577"),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
                TextAlignment       = TextAlignment.Center,
                TextWrapping        = Avalonia.Media.TextWrapping.Wrap,
                Margin              = new Thickness(1),
            };
            ToolTip.SetTip(tb, col);
            var hdr = new Border
            {
                Background      = Brush.Parse("#1a1a30"),
                BorderBrush     = Brush.Parse("#2a2a3a"),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child           = tb,
            };
            Grid.SetColumn(hdr, ci + 1);
            Grid.SetRow(hdr, 0);
            grid.Children.Add(hdr);
        }

        // 데이터 행
        for (int ri = 0; ri < _samples.Count; ri++)
        {
            var sName  = _samples[ri];
            var rowBg  = ri % 2 == 0 ? "#14141e" : "#18182a";
            var cbRow  = new Dictionary<string, CheckBox>(StringComparer.OrdinalIgnoreCase);
            _cbGrid[sName] = cbRow;

            // 시료명 셀
            var nameCell = new Border
            {
                Background = Brush.Parse(rowBg),
                BorderBrush = Brush.Parse("#2a2a3a"), BorderThickness = new Thickness(0,0,1,1),
                Padding = new Thickness(6, 0),
                Child = new TextBlock
                {
                    Text = sName, FontSize = 10, FontFamily = Font,
                    Foreground = Brushes.WhiteSmoke,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                }
            };
            ToolTip.SetTip(nameCell, sName);
            Grid.SetColumn(nameCell, 0);
            Grid.SetRow(nameCell, ri + 1);
            grid.Children.Add(nameCell);

            // 체크박스 셀
            for (int ci = 0; ci < _analysisCols.Count; ci++)
            {
                var colName = _analysisCols[ci];
                var cb = new CheckBox
                {
                    IsChecked           = _quotedItems.Contains(colName),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                };
                cbRow[colName] = cb;

                var cell = new Border
                {
                    Background = Brush.Parse(rowBg),
                    BorderBrush = Brush.Parse("#222230"), BorderThickness = new Thickness(0,0,1,1),
                    Child = cb,
                };
                Grid.SetColumn(cell, ci + 1);
                Grid.SetRow(cell, ri + 1);
                grid.Children.Add(cell);
            }
        }

        return grid;
    }

    // ── 제출 ──────────────────────────────────────────────────────────────
    private async Task OnSubmit()
    {
        var owner = TopLevel.GetTopLevel(this) as Window;

        // 중복 확인
        var dups = _samples
            .Where(s => OrderRequestService.CheckDuplicate(_issue.견적번호, s))
            .ToList();

        if (dups.Count > 0)
        {
            bool overwrite = await ShowConfirm(owner,
                $"이미 등록된 시료명이 있습니다:\n\n" +
                string.Join("\n", dups.Select(d => $"  • {d}")) +
                "\n\n덮어쓰시겠습니까?");
            if (!overwrite) return;
            foreach (var d in dups)
                OrderRequestService.DeleteByKey(_issue.견적번호, d);
        }

        int ok = 0;
        foreach (var sName in _samples)
        {
            if (!_cbGrid.TryGetValue(sName, out var row)) continue;
            var checked_ = row
                .Where(kv => kv.Value.IsChecked == true)
                .Select(kv => kv.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (OrderRequestService.InsertOrderRequest(sName, _issue, checked_)) ok++;
        }

        await ShowMsg(owner, $"✅  {ok}건 의뢰서 제출 완료");
        SubmitCompleted?.Invoke();
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static Control Rc(Control c, int row) { Grid.SetRow(c, row); return c; }

    private static Border HCell(string text, int col, int row)
    {
        var b = new Border
        {
            Background = Brush.Parse("#1e1e30"),
            BorderBrush = Brush.Parse("#333"), BorderThickness = new Thickness(0,0,1,1),
            Padding = new Thickness(6, 0),
            Child = new TextBlock
            {
                Text = text, FontSize = 10, FontFamily = Font,
                FontWeight = FontWeight.SemiBold, Foreground = Brush.Parse("#aaaacc"),
                VerticalAlignment = VerticalAlignment.Center,
            }
        };
        Grid.SetColumn(b, col); Grid.SetRow(b, row);
        return b;
    }

    private static async Task<bool> ShowConfirm(Window? owner, string msg)
    {
        bool r = false;
        var dlg = new Window { Title = "확인", Width = 360, Height = 200,
            CanResize = false, WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Brush.Parse("#1e1e2a") };
        var yes = new Button { Content = "덮어쓰기", Width = 90, Height = 26,
            Background = Brush.Parse("#4a2a2a"), Foreground = Brush.Parse("#f0aeae"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11 };
        var no  = new Button { Content = "취소", Width = 80, Height = 26,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11 };
        yes.Click += (_, _) => { r = true;  dlg.Close(); };
        no.Click  += (_, _) => { r = false; dlg.Close(); };
        dlg.Content = new StackPanel { Margin = new Thickness(20), Spacing = 16, Children = {
            new TextBlock { Text = msg, FontFamily = Font, FontSize = 11,
                Foreground = Brush.Parse("#ddd"), TextWrapping = Avalonia.Media.TextWrapping.Wrap },
            new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8,
                HorizontalAlignment = HorizontalAlignment.Right, Children = { yes, no } }
        }};
        if (owner != null) await dlg.ShowDialog(owner); else dlg.Show();
        return r;
    }

    private static async Task ShowMsg(Window? owner, string msg)
    {
        var dlg = new Window { Title = "완료", Width = 300, Height = 130,
            CanResize = false, WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Brush.Parse("#1e1e2a") };
        var ok = new Button { Content = "확인", Width = 70, Height = 26,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11, HorizontalAlignment = HorizontalAlignment.Right };
        ok.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel { Margin = new Thickness(20), Spacing = 14, Children = {
            new TextBlock { Text = msg, FontFamily = Font, FontSize = 11,
                Foreground = Brush.Parse("#88ee88"), TextWrapping = Avalonia.Media.TextWrapping.Wrap },
            ok }};
        if (owner != null) await dlg.ShowDialog(owner); else dlg.Show();
    }
}
