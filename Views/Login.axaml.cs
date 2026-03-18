using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using System;
using System.Media;
using Avalonia.Platform;
using LibVLCSharp.Shared;

namespace ETA.Views;

public partial class Login : Window
{
    public Login()
    {
        InitializeComponent();
    }

    private void Login_Click(object? sender, RoutedEventArgs e)
    {
        DoLogin();
    }

    private void DoLogin()
    {
        string email = txtEmail?.Text?.Trim() ?? "";
        string password = txtPassword?.Text ?? "";
        var main = new Page1();
        main.Show();
        Close();
        if (email == "admin" && password == "1234")
        {
            //var main = new Page1();
            //main.Show();
            //Close();
        }
        // else { // 실패 시 메시지 추가 가능 }
    }

    private void OnShowPasswordChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (txtPassword != null && tglShowPassword != null)
        {
            txtPassword.PasswordChar = tglShowPassword.IsChecked == true ? '\0' : '*';
        }
    }
    static void PlayClickSound()
    {
        var uri = new Uri("avares://ETA/Assets/POP.wav");
        using var stream = AssetLoader.Open(uri);

        // var player = new SoundPlayer(stream);
        // player.Play();
    }
}