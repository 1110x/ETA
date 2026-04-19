using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class HomePage : UserControl
{
    public HomePage()
    {
        InitializeComponent();
        LoadStats();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadStats()
    {
        var chip = this.FindControl<TextBlock>("LiveCountChip");
        if (chip is null) return;

        try
        {
            var s = HomeStatsService.GetToday();
            chip.Text = $"오늘 · 수질 {s.수질} / 폐수 {s.폐수} / 처리시설 {s.처리시설} · 총 {s.전체}";
        }
        catch (System.Exception ex)
        {
            chip.Text = "DB 연결 실패: " + ex.Message;
        }
    }
}
