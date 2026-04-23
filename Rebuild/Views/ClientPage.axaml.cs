using System;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class ClientPage : UserControl
{
    public ClientPage()
    {
        InitializeComponent();
        Load();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void Load()
    {
        var chip     = this.FindControl<TextBlock>("ClientChip");
        var chipA    = this.FindControl<TextBlock>("ChipActive");
        var chipE    = this.FindControl<TextBlock>("ChipExpired");
        var summary  = this.FindControl<TextBlock>("SummaryLine");
        var list     = this.FindControl<StackPanel>("ClientList");
        if (list is null) return;

        try
        {
            var rows = ContractListService.GetAll();

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;
            var today = DateTime.Today;

            int active = 0, expired = 0, soon = 0;
            foreach (var r in rows)
            {
                var isActive = ContractListService.IsActive(r.계약시작, r.계약종료);
                if (isActive) active++; else expired++;
                if (DateTime.TryParse(r.계약종료, out var endDate))
                {
                    var days = (endDate.Date - today).Days;
                    if (isActive && days <= 30) soon++;
                }

                var row = new Border
                {
                    Padding = new Avalonia.Thickness(0, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("*,80,100,140,100,100") };
                AddCell(g, 0, r.업체명,   false);
                AddCell(g, 1, r.약칭,      true);
                AddCell(g, 2, string.IsNullOrEmpty(r.계약종류) ? "-" : r.계약종류, false);
                AddCell(g, 3, FormatRange(r.계약시작, r.계약종료), true);
                AddCell(g, 4, string.IsNullOrEmpty(r.시설구분) ? "-" : r.시설구분, false);

                var badge = new Border();
                badge.Classes.Add("chip");
                badge.Classes.Add(isActive ? "ok" : "bad");
                badge.Child = new TextBlock { Text = isActive ? "유효" : "만료" };
                Grid.SetColumn(badge, 5);
                g.Children.Add(badge);
                row.Child = g;
                list.Children.Add(row);
            }

            if (rows.Count == 0)
            {
                var empty = new TextBlock { Text = "`계약 DB` 테이블이 비어 있습니다." };
                empty.Classes.Add("muted");
                empty.Margin = new Avalonia.Thickness(10, 12);
                list.Children.Add(empty);
            }

            if (chip    is not null) chip.Text    = $"전체 {rows.Count} · 유효 {active} · 만료 {expired}";
            if (chipA   is not null) chipA.Text   = $"유효 {active}";
            if (chipE   is not null) chipE.Text   = $"만료 {expired}";
            if (summary is not null) summary.Text = soon > 0 ? $"30일 이내 만료 {soon}건 · 재계약 검토" : "임박 만료 없음";
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

    private static string FormatRange(string start, string end)
    {
        string s = Shorten(start), e = Shorten(end);
        if (string.IsNullOrEmpty(s) && string.IsNullOrEmpty(e)) return "-";
        return $"{s} ~ {e}";
    }

    private static string Shorten(string date)
    {
        if (string.IsNullOrEmpty(date)) return "";
        if (DateTime.TryParse(date, out var d)) return d.ToString("yy-MM-dd");
        return date.Length > 10 ? date[..10] : date;
    }
}
