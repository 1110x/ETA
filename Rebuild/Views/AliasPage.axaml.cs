using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class AliasPage : UserControl
{
    public AliasPage()
    {
        InitializeComponent();
        LoadAlias();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadAlias()
    {
        var chip     = this.FindControl<TextBlock>("AliasCountChip");
        var stdList  = this.FindControl<StackPanel>("AliasStdList");
        var mapList  = this.FindControl<StackPanel>("AliasMapList");
        if (stdList is null || mapList is null) return;

        try
        {
            var snapshot = CompoundAliasService.GetSnapshot();

            this.TryFindResource("PaperInkHair", out var hairObj);
            this.TryFindResource("PaperInfoBg",  out var infoObj);
            var hair = hairObj as IBrush;
            var info = infoObj as IBrush;

            foreach (var std in snapshot.Standards)
            {
                var row = new Border
                {
                    Padding = new Avalonia.Thickness(8, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                row.Child = new TextBlock { Text = std, FontSize = 12 };
                stdList.Children.Add(row);
            }

            foreach (var m in snapshot.Mappings)
            {
                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 4),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("180,*,120,Auto") };
                var raw = new TextBlock { Text = m.원문 }; raw.Classes.Add("mono"); raw.Classes.Add("body");
                var alias = new TextBlock { Text = m.약칭 }; alias.Classes.Add("body"); Grid.SetColumn(alias, 1);
                var src = new Border(); src.Classes.Add("chip"); src.Classes.Add(string.IsNullOrEmpty(m.소스) ? "mute" : "info");
                src.Child = new TextBlock { Text = string.IsNullOrEmpty(m.소스) ? "-" : m.소스 };
                Grid.SetColumn(src, 2);
                var btn = new Button { Content = "✎", Classes = { "ghost" }, Padding = new Avalonia.Thickness(8, 2) };
                Grid.SetColumn(btn, 3);
                g.Children.Add(raw); g.Children.Add(alias); g.Children.Add(src); g.Children.Add(btn);
                row.Child = g;
                mapList.Children.Add(row);
            }

            if (chip is not null)
                chip.Text = $"약칭 {snapshot.Standards.Count} · 매핑 {snapshot.Mappings.Count}";
        }
        catch (System.Exception ex)
        {
            if (chip is not null) chip.Text = "DB 연결 실패: " + ex.Message;
        }
    }
}
