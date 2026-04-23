using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;

namespace ETA.Rebuild.Views;

public partial class EcotoxPage : UserControl
{
    public EcotoxPage()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void OnModeA(object? sender, RoutedEventArgs e)
    {
        var a = this.FindControl<Grid>("AContent");
        var b = this.FindControl<ScrollViewer>("BContent");
        var btnA = this.FindControl<Button>("ModeABtn");
        var btnB = this.FindControl<Button>("ModeBBtn");
        if (a is not null) a.IsVisible = true;
        if (b is not null) b.IsVisible = false;
        btnA?.Classes.Add("active");
        btnB?.Classes.Remove("active");
    }

    private void OnModeB(object? sender, RoutedEventArgs e)
    {
        var a = this.FindControl<Grid>("AContent");
        var b = this.FindControl<ScrollViewer>("BContent");
        var btnA = this.FindControl<Button>("ModeABtn");
        var btnB = this.FindControl<Button>("ModeBBtn");
        if (a is not null) a.IsVisible = false;
        if (b is not null) b.IsVisible = true;
        btnA?.Classes.Remove("active");
        btnB?.Classes.Add("active");
    }
}
