using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class ComingSoonPage : UserControl
{
    public ComingSoonPage() : this("?")
    {
    }

    public ComingSoonPage(string title)
    {
        InitializeComponent();
        var tb = this.FindControl<TextBlock>("tbTitle");
        if (tb is not null) tb.Text = title;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}
