using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class StaffPage : UserControl
{
    public StaffPage()
    {
        InitializeComponent();
        LoadStaff();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadStaff()
    {
        var countChip = this.FindControl<TextBlock>("StaffCountChip");
        var pendingLine = this.FindControl<TextBlock>("PendingLine");
        var list = this.FindControl<StackPanel>("StaffList");
        if (list is null) return;

        try
        {
            var rows = StaffListService.GetAll();
            var approved = 0;
            var pending = 0;
            var rejected = 0;

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            foreach (var r in rows)
            {
                switch (r.상태)
                {
                    case "approved": approved++; break;
                    case "pending":  pending++;  break;
                    case "rejected": rejected++; break;
                }

                var row = new Border
                {
                    Padding = new Avalonia.Thickness(10, 6),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("80,*,Auto") };
                var idTb = new TextBlock { Text = r.사번 };
                idTb.Classes.Add("mono"); idTb.Classes.Add("body");
                var nameTb = new TextBlock { Text = r.성명 + "  · " + (string.IsNullOrEmpty(r.부서) ? "-" : r.부서) };
                nameTb.Classes.Add("body");
                Grid.SetColumn(nameTb, 1);

                var chip = new Border();
                chip.Classes.Add("chip");
                chip.Classes.Add(r.상태 switch
                {
                    "approved" => "ok",
                    "pending"  => "warn",
                    "rejected" => "bad",
                    _           => "mute",
                });
                chip.Child = new TextBlock { Text = LabelFor(r.상태) };
                Grid.SetColumn(chip, 2);

                g.Children.Add(idTb);
                g.Children.Add(nameTb);
                g.Children.Add(chip);
                row.Child = g;
                list.Children.Add(row);
            }

            if (rows.Count == 0)
            {
                var empty = new TextBlock { Text = "Agent 테이블에 직원 데이터가 없습니다." };
                empty.Classes.Add("muted");
                empty.Margin = new Avalonia.Thickness(10, 12);
                list.Children.Add(empty);
            }

            if (countChip is not null)
                countChip.Text = $"전체 {rows.Count} · 승인 {approved} · 대기 {pending} · 거부 {rejected}";
            if (pendingLine is not null)
                pendingLine.Text = pending == 0 ? "대기 건 없음" : $"{pending}건 승인 대기 중";
        }
        catch (System.Exception ex)
        {
            if (countChip is not null) countChip.Text = "DB 연결 실패: " + ex.Message;
            if (pendingLine is not null) pendingLine.Text = "조회 실패";
            var err = new TextBlock { Text = ex.GetType().Name + ": " + ex.Message };
            err.Classes.Add("muted");
            err.TextWrapping = Avalonia.Media.TextWrapping.Wrap;
            err.Margin = new Avalonia.Thickness(10, 12);
            list.Children.Add(err);
        }
    }

    private static string LabelFor(string status) => status switch
    {
        "approved" => "승인",
        "pending"  => "대기",
        "rejected" => "거부",
        _           => status ?? "-",
    };
}
