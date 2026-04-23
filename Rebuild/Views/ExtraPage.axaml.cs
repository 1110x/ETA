using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;

namespace ETA.Rebuild.Views;

public partial class ExtraPage : UserControl
{
    private string _currentTab = "risk";

    public ExtraPage()
    {
        InitializeComponent();
        ShowTab("risk");
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void OnTab(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        var tag = btn.Tag as string ?? "risk";
        ShowTab(tag);
    }

    private void ShowTab(string tab)
    {
        _currentTab = tab;

        foreach (var name in new[] { "TabRisk", "TabRepair", "TabPurchase" })
        {
            var b = this.FindControl<Button>(name);
            if (b is not null) b.Classes.Remove("active");
        }
        var target = tab switch
        {
            "repair"   => "TabRepair",
            "purchase" => "TabPurchase",
            _          => "TabRisk",
        };
        this.FindControl<Button>(target)?.Classes.Add("active");

        var title = this.FindControl<TextBlock>("ExtraListTitle");
        var sub   = this.FindControl<TextBlock>("ExtraListSub");
        var list  = this.FindControl<StackPanel>("ExtraList");
        var chipOk   = this.FindControl<TextBlock>("ExtraChipOk");
        var chipWarn = this.FindControl<TextBlock>("ExtraChipWarn");
        var chipBad  = this.FindControl<TextBlock>("ExtraChipBad");

        (string t, string s, (string 날짜, string 내용, string 담당, string 상태, string 비고)[] rows) data = tab switch
        {
            "repair" => (
                "🔧 수리요청 대장",
                "장비/시설 고장 수리 이력",
                new (string, string, string, string, string)[]
                {
                    ("2026-04-18", "GCMS #1 진공 누설", "김영수", "접수", "업체 방문 예정"),
                    ("2026-04-15", "ICP 토치 교체",     "박지훈", "진행", "부품 입고 대기"),
                    ("2026-04-10", "냉각수 순환 펌프",  "이민호", "완료", "-"),
                }
            ),
            "purchase" => (
                "🛒 구매요청 대장",
                "시약/소모품/장비 구매 요청",
                new (string, string, string, string, string)[]
                {
                    ("2026-04-19", "표준용액 DCM 1L",     "azrael",  "대기", "MSDS 첨부"),
                    ("2026-04-17", "VOA Vial 1,000ea",    "김영수",  "승인", "발주 완료"),
                    ("2026-04-14", "GC 컬럼 DB-624",      "박지훈",  "완료", "입고 완료"),
                }
            ),
            _ => (
                "⚠ 위험물 관리대장",
                "보관·사용·폐기 기록",
                new (string, string, string, string, string)[]
                {
                    ("2026-04-19", "메탄올 4L 입고",      "azrael",  "완료", "MSDS OK"),
                    ("2026-04-18", "농염산 500mL 사용",   "김영수",  "완료", "COD 분석"),
                    ("2026-04-17", "폐액 드럼 18L",        "박지훈",  "진행", "위탁업체 예약"),
                }
            ),
        };

        if (title is not null) title.Text = data.t;
        if (sub   is not null) sub.Text   = data.s;

        if (list is null) return;
        list.Children.Clear();

        this.TryFindResource("PaperInkHair", out var hairObj);
        var hair = hairObj as IBrush;

        int ok = 0, warn = 0, bad = 0;
        foreach (var r in data.rows)
        {
            switch (r.상태)
            {
                case "완료": ok++;   break;
                case "진행": case "접수": case "승인": warn++; break;
                case "대기": bad++;  break;
            }

            var row = new Border
            {
                Padding = new Avalonia.Thickness(0, 4),
                BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                BorderBrush = hair,
            };
            var g = new Grid { ColumnDefinitions = new ColumnDefinitions("140,*,100,100,100") };
            var d = new TextBlock { Text = r.날짜 }; d.Classes.Add("mono"); d.Classes.Add("sub");
            var c = new TextBlock { Text = r.내용 }; c.Classes.Add("body"); Grid.SetColumn(c, 1);
            var w = new TextBlock { Text = r.담당 }; w.Classes.Add("body"); Grid.SetColumn(w, 2);
            var s2 = new Border(); s2.Classes.Add("chip");
            s2.Classes.Add(r.상태 switch
            {
                "완료"           => "ok",
                "대기"           => "bad",
                _                 => "warn",
            });
            s2.Child = new TextBlock { Text = r.상태 };
            Grid.SetColumn(s2, 3);
            var n = new TextBlock { Text = r.비고 }; n.Classes.Add("sub"); Grid.SetColumn(n, 4);
            g.Children.Add(d); g.Children.Add(c); g.Children.Add(w); g.Children.Add(s2); g.Children.Add(n);
            row.Child = g;
            list.Children.Add(row);
        }

        if (chipOk   is not null) chipOk.Text   = $"완료 {ok}";
        if (chipWarn is not null) chipWarn.Text = $"진행 {warn}";
        if (chipBad  is not null) chipBad.Text  = $"대기 {bad}";
    }
}
