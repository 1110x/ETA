using System;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class RequestPage : UserControl
{
    public RequestPage()
    {
        InitializeComponent();
        Load();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void Load()
    {
        var chip    = this.FindControl<TextBlock>("ReqChip");
        var list    = this.FindControl<StackPanel>("ReqList");
        var total   = this.FindControl<TextBlock>("SumTotalTb");
        var today   = this.FindControl<TextBlock>("SumTodayTb");
        var week    = this.FindControl<TextBlock>("SumWeekTb");
        var month   = this.FindControl<TextBlock>("SumMonthTb");
        var qaRatio = this.FindControl<TextBlock>("QaRatioTb");
        if (list is null) return;

        try
        {
            var rows = RequestListService.GetRecent(200);

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            var t = DateTime.Today;
            var weekStart  = t.AddDays(-(int)t.DayOfWeek);
            var monthStart = new DateTime(t.Year, t.Month, 1);
            int nToday = 0, nWeek = 0, nMonth = 0, nQa = 0;

            foreach (var r in rows)
            {
                if (DateTime.TryParse(r.채취일자, out var d))
                {
                    if (d.Date == t)          nToday++;
                    if (d.Date >= weekStart)  nWeek++;
                    if (d.Date >= monthStart) nMonth++;
                }
                if (!string.IsNullOrEmpty(r.정도보증)) nQa++;

                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("160,100,*,80,120") };
                AddCell(g, 0, string.IsNullOrEmpty(r.견적번호) ? "-" : r.견적번호, true);
                AddCell(g, 1, Short(r.채취일자), true);
                AddCell(g, 2, string.IsNullOrEmpty(r.시료명) ? "-" : r.시료명, false);
                AddCell(g, 3, string.IsNullOrEmpty(r.약칭) ? "-" : r.약칭, false);

                var qaCell = new Border();
                qaCell.Classes.Add("chip");
                qaCell.Classes.Add(string.IsNullOrEmpty(r.정도보증) ? "mute" : "info");
                qaCell.Child = new TextBlock { Text = string.IsNullOrEmpty(r.정도보증) ? "-" : r.정도보증 };
                Grid.SetColumn(qaCell, 4);
                g.Children.Add(qaCell);
                row.Child = g;
                list.Children.Add(row);
            }

            if (rows.Count == 0)
            {
                var empty = new TextBlock { Text = "`수질분석센터_결과` 테이블이 비어 있습니다." };
                empty.Classes.Add("muted");
                empty.Margin = new Avalonia.Thickness(10, 12);
                list.Children.Add(empty);
            }

            if (chip    is not null) chip.Text    = $"최근 {rows.Count}건";
            if (total   is not null) total.Text   = $"표시: 최근 {rows.Count}건";
            if (today   is not null) today.Text   = $"오늘 {nToday}";
            if (week    is not null) week.Text    = $"이번 주 {nWeek}";
            if (month   is not null) month.Text   = $"이번 달 {nMonth}";
            if (qaRatio is not null)
            {
                var pct = rows.Count == 0 ? 0 : nQa * 100.0 / rows.Count;
                qaRatio.Text = $"정도관리 {nQa} / 전체 {rows.Count}   ({pct:F1}%)";
            }
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
