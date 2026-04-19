using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
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
            _          => new ComingSoonPage(btn.Content?.ToString() ?? "?"),
        };
    }
}
