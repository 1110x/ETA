using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class QuotePage : UserControl
{
    public QuotePage()
    {
        InitializeComponent();
        Load();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void Load()
    {
        var chip       = this.FindControl<TextBlock>("QuoteChip");
        var list       = this.FindControl<StackPanel>("QuoteList");
        var recentCnt  = this.FindControl<TextBlock>("RecentCountTb");
        var recentAmt  = this.FindControl<TextBlock>("RecentAmountTb");
        var topCompany = this.FindControl<TextBlock>("TopCompanyTb");
        if (list is null) return;

        try
        {
            var rows = QuoteListService.GetAll(200);

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            var cutoff = DateTime.Today.AddDays(-30);
            int recentCount = 0;
            decimal recentAmount = 0;
            var companyTally = new Dictionary<string, int>();

            foreach (var r in rows)
            {
                if (DateTime.TryParse(r.견적발행일자, out var d) && d >= cutoff)
                {
                    recentCount++;
                    recentAmount += r.합계금액;
                }
                if (!string.IsNullOrEmpty(r.업체명))
                    companyTally[r.업체명] = companyTally.GetValueOrDefault(r.업체명) + 1;

                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("160,100,*,200,140") };
                AddCell(g, 0, string.IsNullOrEmpty(r.견적번호) ? "-" : r.견적번호, true);
                AddCell(g, 1, Short(r.견적발행일자), true);
                AddCell(g, 2, string.IsNullOrEmpty(r.업체명) ? "-" : r.업체명, false);
                AddCell(g, 3, string.IsNullOrEmpty(r.시료명) ? "-" : r.시료명, false);

                var amtTb = new TextBlock { Text = r.합계금액.ToString("N0") + "원" };
                amtTb.Classes.Add("mono"); amtTb.Classes.Add("body");
                amtTb.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
                Grid.SetColumn(amtTb, 4);
                g.Children.Add(amtTb);
                row.Child = g;
                list.Children.Add(row);
            }

            if (rows.Count == 0)
            {
                var empty = new TextBlock { Text = "`견적발행내역` 테이블이 비어 있습니다." };
                empty.Classes.Add("muted");
                empty.Margin = new Avalonia.Thickness(10, 12);
                list.Children.Add(empty);
            }

            if (chip       is not null) chip.Text       = $"전체 {rows.Count}건 · 최근 30일 {recentCount}건";
            if (recentCnt  is not null) recentCnt.Text  = $"건수 {recentCount}";
            if (recentAmt  is not null) recentAmt.Text  = $"금액 {recentAmount:N0}원";
            if (topCompany is not null)
            {
                var top = companyTally.OrderByDescending(p => p.Value).FirstOrDefault();
                topCompany.Text = string.IsNullOrEmpty(top.Key)
                    ? "상위 업체: -"
                    : $"상위 업체: {top.Key} · {top.Value}건";
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
