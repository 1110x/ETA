using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class LoginWindow : Window
{
    private readonly bool _suppressAutoLogin;

    public LoginWindow() : this(false) { }

    public LoginWindow(bool suppressAutoLogin)
    {
        _suppressAutoLogin = suppressAutoLogin;
        InitializeComponent();
        Title = $"ETA · 로그인 · {AppVersion.Display}";
        var verTag = this.FindControl<TextBlock>("VersionTag");
        if (verTag is not null) verTag.Text = "· " + AppVersion.Display;
        var footerVer = this.FindControl<TextBlock>("FooterVersionTag");
        if (footerVer is not null) footerVer.Text = AppVersion.Display;
        Opened += OnWindowOpened;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private void OnWindowOpened(object? sender, System.EventArgs e)
    {
        var saved = AutoLoginService.Load();
        if (saved is null) return;

        var idBox = this.FindControl<TextBox>("IdBox");
        var pwBox = this.FindControl<TextBox>("PwBox");
        var check = this.FindControl<CheckBox>("AutoLoginCheck");

        if (idBox is not null) idBox.Text = saved.Id;
        if (pwBox is not null) pwBox.Text = saved.Pw;
        if (check is not null) check.IsChecked = true;
    }

    private void OnLoginSubmit(object? sender, RoutedEventArgs e)
    {
        var idBox    = this.FindControl<TextBox>("IdBox");
        var pwBox    = this.FindControl<TextBox>("PwBox");
        var errBlock = this.FindControl<TextBlock>("ErrBlock");
        var check    = this.FindControl<CheckBox>("AutoLoginCheck");

        var id = (idBox?.Text ?? "").Trim();
        var pw = pwBox?.Text ?? "";

        void ShowError(string msg)
        {
            if (errBlock is null) return;
            if (this.TryFindResource("PaperBadFg", out var badObj) && badObj is IBrush bad)
                errBlock.Foreground = bad;
            errBlock.Text = msg;
            errBlock.IsVisible = true;
        }

        if (string.IsNullOrEmpty(id)) { ShowError("사번을 입력하세요."); return; }
        if (string.IsNullOrEmpty(pw)) { ShowError("비밀번호를 입력하세요."); return; }

        try
        {
            var (success, message) = AuthService.ValidateLogin(id, pw);
            if (!success)
            {
                AutoLoginService.Clear();
                ShowError(message);
                return;
            }
        }
        catch (System.Exception ex)
        {
            ShowError("DB 연결 실패: " + ex.Message);
            return;
        }

        if (check?.IsChecked == true) AutoLoginService.Save(id, pw, 1);
        else                          AutoLoginService.Clear();

        ETA.Views.MainPage.CurrentEmployeeId = id;
        var main = new ETA.Views.MainPage();
        main.Show();
        this.Close();
    }

    private async void OnSignUp(object? sender, PointerPressedEventArgs e)
    {
        await new ETA.Views.SignUpWindow().ShowDialog(this);
    }

    private async void OnPasswordRecovery(object? sender, PointerPressedEventArgs e)
    {
        var idBox    = this.FindControl<TextBox>("IdBox");
        var pwBox    = this.FindControl<TextBox>("PwBox");
        var errBlock = this.FindControl<TextBlock>("ErrBlock");
        var empId    = (idBox?.Text ?? "").Trim();

        if (string.IsNullOrEmpty(empId))
        {
            if (errBlock is not null)
            {
                errBlock.Text = "사번을 먼저 입력하세요.";
                errBlock.IsVisible = true;
            }
            return;
        }

        var dlg = new ETA.Views.ChangePasswordWindow(empId, isForced: false);
        await dlg.ShowDialog(this);

        if (dlg.IsChanged)
        {
            if (pwBox is not null) { pwBox.Text = ""; pwBox.Focus(); }
            if (errBlock is not null)
            {
                errBlock.Foreground = Avalonia.Media.Brush.Parse("#10B981");
                errBlock.Text       = "✓ 비밀번호가 변경되었습니다. 새 비밀번호로 로그인해주세요.";
                errBlock.IsVisible  = true;
            }
        }
    }
}
