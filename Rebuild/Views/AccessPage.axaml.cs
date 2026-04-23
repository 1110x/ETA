using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class AccessPage : UserControl
{
    private static readonly (string Key, string Label)[] _menus =
    {
        ("home",     "🏠 홈"),
        ("water",    "📇 수질분석센터"),
        ("input",    "✍ 분석결과입력"),
        ("file",     "🗂 파일플로우"),
        ("note",     "📓 시험기록부"),
        ("facility", "🏭 처리시설"),
        ("waste",    "💧 폐수배출업소"),
        ("ecotox",   "🧬 생태독성"),
        ("cheuk",    "🤖 측정인.kr"),
        ("parser",   "⚙ 파서"),
        ("admin",    "🛠 관리자"),
    };

    public AccessPage()
    {
        InitializeComponent();
        LoadAll();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void LoadAll()
    {
        var chip    = this.FindControl<TextBlock>("AccessCountChip");
        var staffLv = this.FindControl<StackPanel>("AccessStaffList");
        var matrix  = this.FindControl<StackPanel>("AccessMatrix");
        if (staffLv is null || matrix is null) return;

        try
        {
            var staff = StaffListService.GetAll();

            this.TryFindResource("PaperInkHair", out var hairObj);
            var hair = hairObj as IBrush;

            foreach (var s in staff)
            {
                var row = new Border
                {
                    Padding = new Avalonia.Thickness(8, 5),
                    BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
                    BorderBrush = hair,
                };
                var g = new Grid { ColumnDefinitions = new ColumnDefinitions("60,*,Auto") };
                var idTb = new TextBlock { Text = s.사번 }; idTb.Classes.Add("mono");
                var name = new TextBlock { Text = s.성명 }; name.Classes.Add("body"); Grid.SetColumn(name, 1);
                var dot  = new TextBlock { Text = s.상태 == "approved" ? "●" : "○" }; dot.Classes.Add("muted"); Grid.SetColumn(dot, 2);
                g.Children.Add(idTb); g.Children.Add(name); g.Children.Add(dot);
                row.Child = g;
                staffLv.Children.Add(row);
            }

            foreach (var (key, label) in _menus)
            {
                var row = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto,Auto"),
                    Margin = new Avalonia.Thickness(0, 2),
                };
                var lbl = new TextBlock { Text = label }; lbl.Classes.Add("body");
                var cb1 = new CheckBox { Margin = new Avalonia.Thickness(16, 0) };
                var cb2 = new CheckBox { Margin = new Avalonia.Thickness(16, 0) };
                var cb3 = new CheckBox { Margin = new Avalonia.Thickness(16, 0) };
                Grid.SetColumn(cb1, 1); Grid.SetColumn(cb2, 2); Grid.SetColumn(cb3, 3);
                row.Children.Add(lbl); row.Children.Add(cb1); row.Children.Add(cb2); row.Children.Add(cb3);
                matrix.Children.Add(row);
            }

            if (chip is not null)
                chip.Text = $"직원 {staff.Count} · 메뉴 {_menus.Length}";
        }
        catch (System.Exception ex)
        {
            if (chip is not null) chip.Text = "DB 연결 실패: " + ex.Message;
        }
    }
}
