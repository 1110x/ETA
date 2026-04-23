using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views;

/// <summary>
/// 페놀류(직접법/추출법) · TOC(NPOC/TCIC) 분석방법 충돌 선택 팝업.
/// 동일 시료에 기존 결과와 다른 방법으로 재업로드될 때, 사용자가 최종값을
/// 직접 고르도록 한다. 원자료 시험기록부는 방법별로 분리 저장되어 손실 없음.
/// </summary>
public class MethodChoiceDialog : Window
{
    public class ConflictItem
    {
        public int SampleId { get; set; }
        public string 시료명 { get; set; } = "";
        public string OldValue { get; set; } = "";
        public string OldMethod { get; set; } = "";
        public string NewValue { get; set; } = "";
        public string NewMethod { get; set; } = "";
        /// <summary>true=새값 적용, false=기존값 유지</summary>
        public bool UseNew { get; set; } = true;
    }

    public bool Confirmed { get; private set; }
    public IReadOnlyList<ConflictItem> Items => _items;

    private readonly List<ConflictItem> _items;
    private readonly List<(RadioButton old, RadioButton @new)> _radios = new();

    public MethodChoiceDialog(string itemLabel, IList<ConflictItem> items)
    {
        _items = items.ToList();

        Title = $"{itemLabel} 분석방법 충돌 — 최종값 선택";
        Width = 720;
        Height = System.Math.Min(560, 220 + _items.Count * 42);
        CanResize = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = AppTheme.BgPrimary;

        BuildUI(itemLabel);
    }

    private void BuildUI(string itemLabel)
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        var title = new TextBlock
        {
            Text = $"⚠  {itemLabel} — 동일 시료에 다른 분석방법 결과 존재",
            FontFamily = font,
            FontSize = AppTheme.FontLG,
            FontWeight = FontWeight.Bold,
            Foreground = AppTheme.FgWarn,
            Margin = new Thickness(0, 0, 0, 6),
        };

        var sub = new TextBlock
        {
            Text = "최종값으로 저장할 방법을 시료별로 선택하세요. 선택하지 않은 쪽의 원자료는 시험기록부에 그대로 남습니다.",
            FontFamily = font,
            FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10),
        };

        // 일괄 버튼
        var allOld = new Button { Content = "모두 기존값 유지", Margin = new Thickness(0, 0, 6, 0) };
        var allNew = new Button { Content = "모두 새값 적용" };
        allOld.Click += (_, _) => { foreach (var r in _radios) r.old.IsChecked = true; };
        allNew.Click += (_, _) => { foreach (var r in _radios) r.@new.IsChecked = true; };
        var bulkRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 0, 0, 8),
            Children = { allOld, allNew }
        };

        // 헤더
        var header = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("2*,1*,1*,1*,1*"),
            Margin = new Thickness(0, 0, 0, 2),
        };
        AddHeader(header, 0, "시료명");
        AddHeader(header, 1, "기존값 (방법)");
        AddHeader(header, 2, "신규값 (방법)");
        AddHeader(header, 3, "기존 유지");
        AddHeader(header, 4, "새값 적용");

        // 행
        var listPanel = new StackPanel { Spacing = 2 };
        for (int i = 0; i < _items.Count; i++)
        {
            var it = _items[i];
            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("2*,1*,1*,1*,1*"),
                Background = i % 2 == 0 ? AppTheme.BgSecondary : AppTheme.BgCard,
            };

            AddCell(rowGrid, 0, it.시료명);
            AddCell(rowGrid, 1, $"{it.OldValue}  ({it.OldMethod})", AppTheme.FgMuted);
            AddCell(rowGrid, 2, $"{it.NewValue}  ({it.NewMethod})", AppTheme.FgInfo);

            string grp = $"g_{i}";
            var rbOld = new RadioButton
            {
                GroupName = grp,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                IsChecked = !it.UseNew,
            };
            var rbNew = new RadioButton
            {
                GroupName = grp,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                IsChecked = it.UseNew,
            };
            var localIt = it;
            rbOld.IsCheckedChanged += (_, _) => { if (rbOld.IsChecked == true) localIt.UseNew = false; };
            rbNew.IsCheckedChanged += (_, _) => { if (rbNew.IsChecked == true) localIt.UseNew = true; };

            Grid.SetColumn(rbOld, 3); rowGrid.Children.Add(rbOld);
            Grid.SetColumn(rbNew, 4); rowGrid.Children.Add(rbNew);
            _radios.Add((rbOld, rbNew));

            listPanel.Children.Add(rowGrid);
        }

        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = listPanel,
        };

        // 확인/취소
        var okBtn = new Button
        {
            Content = "확인",
            Width = 90,
            Height = 32,
            Background = new SolidColorBrush(Color.Parse("#2a4a2a")),
            Foreground = Brushes.White,
        };
        var cancelBtn = new Button
        {
            Content = "취소",
            Width = 90,
            Height = 32,
            Margin = new Thickness(0, 0, 6, 0),
        };
        okBtn.Click += (_, _) => { Confirmed = true; Close(); };
        cancelBtn.Click += (_, _) => { Confirmed = false; Close(); };

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 10, 0, 0),
            Children = { cancelBtn, okBtn }
        };

        var root = new DockPanel();
        var topBlock = new StackPanel { Children = { title, sub, bulkRow, header } };
        DockPanel.SetDock(topBlock, Dock.Top);
        DockPanel.SetDock(btnRow, Dock.Bottom);
        root.Children.Add(topBlock);
        root.Children.Add(btnRow);
        root.Children.Add(scroll);

        Content = new Border
        {
            Padding = new Thickness(18),
            Child = root,
        };
    }

    private static void AddHeader(Grid g, int col, string text)
    {
        var tb = new TextBlock
        {
            Text = text,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgSecondary,
            FontSize = AppTheme.FontBase,
            HorizontalAlignment = col == 0 ? HorizontalAlignment.Left : HorizontalAlignment.Center,
            Margin = new Thickness(6, 4),
        };
        Grid.SetColumn(tb, col);
        g.Children.Add(tb);
    }

    private static void AddCell(Grid g, int col, string text, IBrush? fg = null)
    {
        var tb = new TextBlock
        {
            Text = text,
            Foreground = fg ?? AppTheme.FgPrimary,
            FontSize = AppTheme.FontBase,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = col == 0 ? HorizontalAlignment.Left : HorizontalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(6, 6),
        };
        Grid.SetColumn(tb, col);
        g.Children.Add(tb);
    }
}
