using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class NotebookPage : UserControl
{
    public NotebookPage()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}
