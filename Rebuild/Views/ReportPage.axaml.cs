using System;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class ReportPage : UserControl
{
    public ReportPage()
    {
        InitializeComponent();
        Load();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void Load()
    {
        var chip     = this.FindControl<TextBlock>("RepChip");
        var list     = this.FindControl<StackPanel>("RepList");
        var totalTb  = this.FindControl<TextBlock>("RepTotalTb");
        var sampleTb = this.FindControl<TextBlock>("RepSamplesTb");
        if (list is null) return;

        try
        {
            var groups = ReportListService.GetGroups(120);

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            int totalSamples = 0;
            foreach (var r in groups)
            {
                totalSamples += r.시료수;

                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("180,120,*,80,120") };
                AddCell(g, 0, r.견적번호,  true);
                AddCell(g, 1, Short(r.채취일자), true);
                AddCell(g, 2, string.IsNullOrEmpty(r.약칭) ? "-" : r.약칭, false);

                var cntTb = new TextBlock { Text = r.시료수.ToString("N0") };
                cntTb.Classes.Add("mono"); cntTb.Classes.Add("body");
                cntTb.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
                Grid.SetColumn(cntTb, 3);
                g.Children.Add(cntTb);

                var status = new Border();
                status.Classes.Add("chip");
                status.Classes.Add("info");
                status.Child = new TextBlock { Text = "발행 대기" };
                status.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
                Grid.SetColumn(status, 4);
                g.Children.Add(status);

                row.Child = g;
                list.Children.Add(row);
            }

            if (groups.Count == 0)
            {
                var empty = new TextBlock { Text = "발행 대상 견적번호가 없습니다." };
                empty.Classes.Add("muted");
                empty.Margin = new Avalonia.Thickness(10, 12);
                list.Children.Add(empty);
            }

            if (chip     is not null) chip.Text     = $"그룹 {groups.Count} · 시료 {totalSamples}";
            if (totalTb  is not null) totalTb.Text  = $"전체 {groups.Count}건";
            if (sampleTb is not null) sampleTb.Text = $"시료 {totalSamples}건";
        }
        catch (Exception ex)
        {
            if (chip is not null) chip.Text = "DB 연결 실패: " + ex.Message;
            var err = new TextBlock { Text = ex.GetType().Name + ": " + ex.Message };
            err.Classes.Add("muted");
            err.TextWrapping = TextWrapping.Wrap;
            err.Margin = new Avalonia.Thickness(10, 12);
            list.Children.Add(err);
        }
    }

    private static void AddCell(Grid g, int col, string text, bool mono)
    {
        var tb = new TextBlock { Text = text };
        tb.Classes.Add("body");
        if (mono) tb.Classes.Add("mono");
        Grid.SetColumn(tb, col);
        g.Children.Add(tb);
    }

    private static string Short(string date)
    {
        if (string.IsNullOrEmpty(date)) return "-";
        if (DateTime.TryParse(date, out var d)) return d.ToString("yy-MM-dd");
        return date.Length > 10 ? date[..10] : date;
    }
}
