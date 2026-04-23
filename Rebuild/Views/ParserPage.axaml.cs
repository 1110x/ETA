using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class ParserPage : UserControl
{
    public ParserPage()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}
