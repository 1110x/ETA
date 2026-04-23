using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using ETA.Services.SERVICE2;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ETA.Views;

public partial class WaybleLoginWindow : Window
{
    public bool LoginSucceeded { get; private set; }

    private readonly CancellationTokenSource _cts = new();

    public WaybleLoginWindow()
    {
        InitializeComponent();
        Opened  += OnOpened;
        Closing += (_, _) => _cts.Cancel();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void OnOpened(object? sender, EventArgs e)
    {
        var saved = WaybleCredentialService.Load();
        if (saved is null) return;

        var id = this.FindControl<TextBox>("IdBox");
        var pw = this.FindControl<TextBox>("PwBox");
        var ck = this.FindControl<CheckBox>("AutoCheck");
        if (id is not null) id.Text      = saved.UserId;
        if (pw is not null) pw.Text      = saved.Password;
        if (ck is not null) ck.IsChecked = saved.AutoLogin;

        if (saved.AutoLogin)
            Dispatcher.UIThread.Post(() => OnSubmit(this, new RoutedEventArgs()),
                                     DispatcherPriority.Background);
    }

    private async void OnSubmit(object? sender, RoutedEventArgs e)
    {
        var idBox = this.FindControl<TextBox>("IdBox");
        var pwBox = this.FindControl<TextBox>("PwBox");
        var ck    = this.FindControl<CheckBox>("AutoCheck");
        var st    = this.FindControl<TextBlock>("StatusTb");
        var btn   = this.FindControl<Button>("SubmitBtn");

        var id = idBox?.Text?.Trim() ?? "";
        var pw = pwBox?.Text ?? "";

        void Status(string msg, bool error = false) =>
            Dispatcher.UIThread.Post(() =>
            {
                if (st is null) return;
                st.Text = msg;
                var key      = error ? "PaperBadFg" : "PaperFgMuted";
                var fallback = error ? "#c0392b"    : "#888888";
                if (this.TryFindResource(key, null, out var r) && r is IBrush br)
                    st.Foreground = br;
                else
                    st.Foreground = Brush.Parse(fallback);
            });

        if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(pw))
        {
            Status("아이디/비밀번호를 입력하세요.", error: true);
            return;
        }

        if (btn is not null) btn.IsEnabled = false;

        WaybleCdpLogin.Result result;
        try
        {
            result = await WaybleCdpLogin.LoginAsync(
                id, pw, s => Status(s, error: false), _cts.Token);
        }
        catch (OperationCanceledException)
        {
            Status("취소됨", error: true);
            if (btn is not null) btn.IsEnabled = true;
            return;
        }
        catch (Exception ex)
        {
            Status("오류: " + ex.Message, error: true);
            if (btn is not null) btn.IsEnabled = true;
            return;
        }

        if (result.Ok && !string.IsNullOrEmpty(result.Cookie))
        {
            WaybleSession.SeedStpSession(result.Cookie!);
            if (ck?.IsChecked == true) WaybleCredentialService.Save(id, pw, true);
            else                       WaybleCredentialService.Clear();
            LoginSucceeded = true;
            Close();
        }
        else
        {
            Status(result.Message, error: true);
            if (btn is not null) btn.IsEnabled = true;
        }
    }

    private void OnClose(object? sender, RoutedEventArgs e) => Close();
}
