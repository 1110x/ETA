using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Services;

namespace ETA.Views;

public partial class Page1 : Window
{
    public Page1()
    {
        InitializeComponent();
    }

    private async void Admin_Click(object? sender, RoutedEventArgs e)
    {
        MainText.IsVisible = false;
        MainList.IsVisible = true;

        MainList.ItemsSource = AdminServices.GetAdminMenu();
    }

    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
    }
}