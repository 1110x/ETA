using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class HomePage : UserControl
{
    public HomePage()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}
