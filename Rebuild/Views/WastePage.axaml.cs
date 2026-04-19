using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class WastePage : UserControl
{
    public WastePage()
    {
        InitializeComponent();
        LoadTree();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadTree()
    {
        var tree = this.FindControl<StackPanel>("WasteTree");
        if (tree is null) return;

        var groups = WasteCompanyListService.GetGrouped();
        if (groups.Count == 0) return;

        tree.Children.Clear();

        this.TryFindResource("PaperHeader", out var headerBgObj);
        var headerBg = headerBgObj as IBrush;
        bool first = true;
        foreach (var (name, rows) in groups)
        {
            var group = new StackPanel { Spacing = 3 };

            var header = new Border { Padding = new Thickness(8, 5) };
            header.Classes.Add("card-flat");
            if (headerBg is not null) header.Background = headerBg;

            var hg = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
            var caret = new TextBlock { Text = first ? "▾" : "▸", Margin = new Thickness(0, 0, 6, 0) };
            caret.Classes.Add("muted");
            var title = new TextBlock { Text = LabelFor(name) };
            title.Classes.Add("h3");
            Grid.SetColumn(title, 1);
            var chip = new Border();
            chip.Classes.Add("chip"); chip.Classes.Add("ok");
            chip.Child = new TextBlock { Text = rows.Count.ToString() };
            Grid.SetColumn(chip, 2);
            hg.Children.Add(caret); hg.Children.Add(title); hg.Children.Add(chip);
            header.Child = hg;
            group.Children.Add(header);

            if (first)
            {
                int shown = 0;
                foreach (var r in rows)
                {
                    if (shown >= 8) break;
                    group.Children.Add(BuildRow(r.업체명, shown == 0));
                    shown++;
                }
                if (rows.Count > 8)
                {
                    var more = new TextBlock { Text = $"… +{rows.Count - 8}개 더", Margin = new Thickness(20, 4, 8, 4) };
                    more.Classes.Add("sub");
                    group.Children.Add(more);
                }
            }

            tree.Children.Add(group);
            first = false;
        }
    }

    private Border BuildRow(string name, bool active)
    {
        var row = new Border { Padding = new Thickness(20, 6, 8, 6) };
        row.Classes.Add(active ? "row-active" : "row");
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        var t = new TextBlock { Text = name };
        t.Classes.Add("body");
        g.Children.Add(t);
        row.Child = g;
        return row;
    }

    private static string LabelFor(string group) => group switch
    {
        "여수" => "여수 · TOC 1법",
        "율촌" => "율촌 · TOC 2법",
        "세풍" => "세풍",
        _      => group,
    };
}
