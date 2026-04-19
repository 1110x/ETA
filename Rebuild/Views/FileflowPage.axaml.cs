using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class FileflowPage : UserControl
{
    public FileflowPage()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}
