using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Markup.Xaml;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        Title = $"ETA · 수질분석센터 · {AppVersion.Display}";
        var brand = this.FindControl<TextBlock>("BrandTag");
        if (brand is not null) brand.Text = $"· 수질분석센터 · {AppVersion.Display}";
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void OnNav(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;

        foreach (var b in this.GetLogicalDescendants())
        {
            if (b is Button nav && nav.Classes.Contains("navtab"))
            {
                nav.Classes.Remove("active");
            }
        }
        btn.Classes.Add("active");

        var host = this.FindControl<ContentControl>("PageHost");
        if (host is null) return;

        host.Content = (btn.Tag as string) switch
        {
            "home"     => new HomePage(),
            "input"    => new InputPage(),
            "file"     => new FileflowPage(),
            "note"     => new NotebookPage(),
            "facility" => new FacilityPage(),
            "waste"    => new WastePage(),
            "ecotox"   => new EcotoxPage(),
            "cheuk"    => new CheukPage(),
            "parser"   => new ParserPage(),
            _          => new ComingSoonPage(btn.Content?.ToString() ?? "?"),
        };
    }

    private void OnNavItem(object? sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi) return;
        var tag = mi.Tag as string ?? "";
        var label = mi.Header?.ToString() ?? "?";

        foreach (var b in this.GetLogicalDescendants())
        {
            if (b is Button nav && nav.Classes.Contains("navtab"))
                nav.Classes.Remove("active");
        }

        var parentName = tag switch
        {
            "client" or "quote" or "request" or "report" => "NavWater",
            _ => "NavAdmin",
        };
        this.FindControl<Button>(parentName)?.Classes.Add("active");

        var host = this.FindControl<ContentControl>("PageHost");
        if (host is null) return;

        host.Content = tag switch
        {
            "client"  => new ClientPage(),
            "quote"   => new QuotePage(),
            "request" => new RequestPage(),
            "report"  => new ReportPage(),
            "staff"   => new StaffPage(),
            "alias"   => new AliasPage(),
            "access"  => new AccessPage(),
            "server"  => new ServerPage(),
            "extra"   => new ExtraPage(),
            _         => new ComingSoonPage(label),
        };
    }
}
