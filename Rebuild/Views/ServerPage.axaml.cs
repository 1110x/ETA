using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class ServerPage : UserControl
{
    public ServerPage()
    {
        InitializeComponent();
        LoadInfo();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadInfo()
    {
        var chip = this.FindControl<TextBlock>("ServerChip");
        var list = this.FindControl<StackPanel>("TableStatList");

        try
        {
            var info = ServerInfoService.GetInfo();
            this.FindControl<TextBlock>("SrvHost")!.Text = string.IsNullOrEmpty(info.Host) ? "-" : info.Host;
            this.FindControl<TextBlock>("SrvPort")!.Text = string.IsNullOrEmpty(info.Port) ? "-" : info.Port;
            this.FindControl<TextBlock>("SrvDb")!.Text   = string.IsNullOrEmpty(info.Database) ? "-" : info.Database;
            this.FindControl<TextBlock>("SrvUser")!.Text = string.IsNullOrEmpty(info.User) ? "-" : info.User;
            this.FindControl<TextBlock>("SrvVer")!.Text  = string.IsNullOrEmpty(info.Version) ? "-" : info.Version;

            if (chip is not null)
                chip.Text = info.IsConnected ? $"● 접속됨 · {info.Host}" : "접속 실패";

            if (list is null) return;

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            foreach (var t in info.Tables)
            {
                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 4),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("*,120,120") };
                var nameTb = new TextBlock { Text = t.Table }; nameTb.Classes.Add("mono"); nameTb.Classes.Add("body");
                var totalTb = new TextBlock { Text = t.Total.ToString("N0") };
                totalTb.Classes.Add("mono"); totalTb.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
                Grid.SetColumn(totalTb, 1);
                var todayTb = new TextBlock { Text = t.Today.ToString("N0") };
                todayTb.Classes.Add("mono"); todayTb.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
                Grid.SetColumn(todayTb, 2);
                g.Children.Add(nameTb); g.Children.Add(totalTb); g.Children.Add(todayTb);
                row.Child = g;
                list.Children.Add(row);
            }
        }
        catch (System.Exception ex)
        {
            if (chip is not null) chip.Text = "DB 연결 실패: " + ex.Message;
        }
    }
}
