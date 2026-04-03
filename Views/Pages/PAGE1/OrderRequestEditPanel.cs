using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks;
using ETA.Views;

namespace ETA.Views.Pages.PAGE1;

public class OrderRequestEditPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private QuotationIssue    _issue        = null!;
    private List<string>      _samples      = new();
    private List<string>      _analysisCols = new();
    private HashSet<string>   _quotedItems  = new();

    private readonly Dictionary<string, Dictionary<string, CheckBox>> _cbGrid = new();

    public event Action? SubmitCompleted;
    public event Action? Cancelled;

    public OrderRequestEditPanel() { }

    public void Load(QuotationIssue issue, List<string> samples, HashSet<string> quotedItems)
    {
        _issue        = issue;
        _samples      = samples;
        _quotedItems  = quotedItems;
        _analysisCols = OrderRequestService.GetAnalysisColumns();
        _cbGrid.Clear();
        Content = BuildUI();
    }

    private Control BuildUI()
    {
        var titleBar = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
        titleBar.Children.Add(new StackPanel
        {
            Spacing = 2,
            Children =
            {
                new TextBlock
                {
                    Text = "📋  의뢰서 편집",
                    FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
                    FontFamily = Font, Foreground = AppTheme.FgPrimary,
                },
                new TextBlock
                {
                    Text = $"{_issue.업체명}  |  {_issue.견적번호}  |  {_issue.견적구분}",
                    FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brush.Parse("#666688"),
                },
            }
        });

        var btnSubmit = new Button
        {
            Content = "🔄 업데이트", Height = 26, FontSize = AppTheme.FontBase, FontFamily = Font,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 0), Margin = new Thickness(0, 0, 6, 0),
        };
        var btnCancel = new Button
        {
            Content = "← 뒤로", Height = 26, FontSize = AppTheme.FontBase, FontFamily = Font,
            Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 0),
        };
        btnSubmit.Click += async (_, _) => await OnSubmit();
        btnCancel.Click += (_, _) => Cancelled?.Invoke();
        Grid.SetColumn(btnSubmit, 1);
        Grid.SetColumn(btnCancel, 2);
        titleBar.Children.Add(btnSubmit);
        titleBar.Children.Add(btnCancel);

        var info = new TextBlock
        {
            Text = $"시료 {_samples.Count}건  ·  견적항목 {_quotedItems.Count}개 기본 체크  ·  체크 해제로 제외 가능",
            FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppTheme.FgDimmed,
            Margin = new Thickness(0, 4, 0, 6),
        };

        var scroll = new ScrollViewer
        {
            Content = BuildCheckGrid(),
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
                Rc(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0,5,0,5) }, 1),
                Rc(info, 2),
                Rc(scroll, 3),
            }
        };
    }

    // 항목명=행(수직), 시료명=열(수평)
    private Panel BuildCheckGrid()
    {
        var grid = new Grid();

        var colDefs = new ColumnDefinitions();
        colDefs.Add(new ColumnDefinition { Width = new GridLength(140) });
        foreach (var _ in _samples)
            colDefs.Add(new ColumnDefinition { Width = new GridLength(90) });
        grid.ColumnDefinitions = colDefs;

        var rowDefs = new RowDefinitions();
        rowDefs.Add(new RowDefinition { Height = new GridLength(36) });
        foreach (var _ in _analysisCols)
            rowDefs.Add(new RowDefinition { Height = new GridLength(26) });
        grid.RowDefinitions = rowDefs;

        grid.Children.Add(HCell("항목명", 0, 0));

        for (int ci = 0; ci < _samples.Count; ci++)
        {
            var sName = _samples[ci];
            var tb = new TextBlock
            {
                Text                = sName,
                FontSize            = AppTheme.FontXS,
                FontFamily          = Font,
                Foreground          = Brush.Parse("#ccccee"),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
                TextAlignment       = TextAlignment.Center,
                TextWrapping        = Avalonia.Media.TextWrapping.Wrap,
                Margin              = new Thickness(2),
            };
            ToolTip.SetTip(tb, sName);
            var hdr = new Border
            {
                Background      = Brush.Parse("#1a1a30"),
                BorderBrush     = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child           = tb,
            };
            Grid.SetColumn(hdr, ci + 1);
            Grid.SetRow(hdr, 0);
            grid.Children.Add(hdr);
        }

        foreach (var sName in _samples)
            _cbGrid[sName] = new Dictionary<string, CheckBox>(StringComparer.OrdinalIgnoreCase);

        for (int ri = 0; ri < _analysisCols.Count; ri++)
        {
            var colName = _analysisCols[ri];
            bool isQ    = _quotedItems.Contains(colName);
            var rowBg   = ri % 2 == 0 ? "#14141e" : "#18182a";

            var nameCell = new Border
            {
                Background      = Brush.Parse(rowBg),
                BorderBrush     = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding         = new Thickness(6, 0),
                Child           = new TextBlock
                {
                    Text              = colName,
                    FontSize          = AppTheme.FontXS,
                    FontFamily        = Font,
                    Foreground        = isQ ? AppTheme.FgSuccess : AppTheme.FgMuted,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                }
            };
            ToolTip.SetTip(nameCell, colName);
            Grid.SetColumn(nameCell, 0);
            Grid.SetRow(nameCell, ri + 1);
            grid.Children.Add(nameCell);

            for (int ci = 0; ci < _samples.Count; ci++)
            {
                var sName = _samples[ci];
                var cb = new CheckBox
                {
                    IsChecked           = _quotedItems.Contains(colName),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                };
                _cbGrid[sName][colName] = cb;

                var cell = new Border
                {
                    Background      = Brush.Parse(rowBg),
                    BorderBrush     = AppTheme.BorderSubtle,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Child           = cb,
                };
                Grid.SetColumn(cell, ci + 1);
                Grid.SetRow(cell, ri + 1);
                grid.Children.Add(cell);
            }
        }

        return grid;
    }

    private async Task OnSubmit()
    {
        var owner = TopLevel.GetTopLevel(this) as Window;

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

    private static Control Rc(Control c, int row) { Grid.SetRow(c, row); return c; }

    private static Border HCell(string text, int col, int row)
    {
        var b = new Border
        {
            Background      = AppTheme.BgPrimary,
            BorderBrush     = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Padding         = new Thickness(6, 0),
            Child           = new TextBlock
            {
                Text              = text,
                FontSize          = AppTheme.FontSM,
                FontFamily        = Font,
                FontWeight        = FontWeight.SemiBold,
                Foreground        = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
            }
        };
        Grid.SetColumn(b, col);
        Grid.SetRow(b, row);
        return b;
    }

    private static async Task<bool> ShowConfirm(Window? owner, string msg)
    {
        bool r = false;
        var dlg = new Window
        {
            Title = "확인", Width = 360, Height = 200,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgPrimary,
        };
        var yes = new Button
        {
            Content = "덮어쓰기", Width = 90, Height = 26,
            Background = AppTheme.BgDanger, Foreground = AppTheme.FgDanger,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontBase,
        };
        var no = new Button
        {
            Content = "취소", Width = 80, Height = 26,
            Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontBase,
        };
        yes.Click += (_, _) => { r = true;  dlg.Close(); };
        no.Click  += (_, _) => { r = false; dlg.Close(); };
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text = msg, FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgPrimary,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 8,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children = { yes, no },
                },
            }
        };
        if (owner != null) await dlg.ShowDialog(owner); else dlg.Show();
        return r;
    }

    private static async Task ShowMsg(Window? owner, string msg)
    {
        var dlg = new Window
        {
            Title = "완료", Width = 300, Height = 130,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgPrimary,
        };
        var ok = new Button
        {
            Content = "확인", Width = 70, Height = 26,
            Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontBase,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        ok.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock
                {
                    Text = msg, FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgSuccess,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                ok,
            }
        };
        if (owner != null) await dlg.ShowDialog(owner); else dlg.Show();
    }
}
