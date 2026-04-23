using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ETA.Rebuild.Services;

namespace ETA.Rebuild.Views;

public partial class LoginWindow : Window
{
    private int _pickedVersion = 2;
    private bool _autoSubmitScheduled = false;
    private readonly bool _suppressAutoLogin;

    public LoginWindow() : this(false) { }

    public LoginWindow(bool suppressAutoLogin)
    {
        _suppressAutoLogin = suppressAutoLogin;
        InitializeComponent();
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
        Pick(saved.Version == 1 ? 1 : 2);

        // 자동 Submit 비활성화: ID/PW 만 자동으로 채우고 로그인 버튼은 사용자가 수동으로 누름.
        // (기존 자동 실행 블록 제거)
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

        if (check?.IsChecked == true) AutoLoginService.Save(id, pw, _pickedVersion);
        else                          AutoLoginService.Clear();

        if (_pickedVersion == 1)
        {
            ETA.Views.MainPage.CurrentEmployeeId = id;
            var v1 = new ETA.Views.MainPage();
            v1.Show();
            this.Close();
            return;
        }

        var main = new MainWindow();
        main.Show();
        this.Close();
    }

    private void OnPickV1(object? sender, PointerPressedEventArgs e) => Pick(1);
    private void OnPickV2(object? sender, PointerPressedEventArgs e) => Pick(2);

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

    private void Pick(int version)
    {
        _pickedVersion = version;

        var v1 = this.FindControl<Border>("V1Card");
        var v2 = this.FindControl<Border>("V2Card");
        var d1 = this.FindControl<TextBlock>("V1Dot");
        var d2 = this.FindControl<TextBlock>("V2Dot");
        var label = this.FindControl<TextBlock>("PickLabel");

        this.TryFindResource("PaperAccent",   out var accentObj);
        this.TryFindResource("PaperInfoBg",   out var infoBgObj);
        this.TryFindResource("PaperFgMuted",  out var mutedFgObj);
        this.TryFindResource("PaperInkHair",  out var inkHairObj);
        this.TryFindResource("PaperCard",     out var cardBgObj);
        var accent  = accentObj  as IBrush;
        var infoBg  = infoBgObj  as IBrush;
        var mutedFg = mutedFgObj as IBrush;
        var inkHair = inkHairObj as IBrush;
        var cardBg  = cardBgObj  as IBrush;

        if (version == 1)
        {
            if (v1 is not null) { v1.BorderBrush = accent ?? v1.BorderBrush; v1.BorderThickness = new Avalonia.Thickness(2); v1.Background = infoBg ?? v1.Background; }
            if (v2 is not null) { v2.BorderBrush = inkHair ?? v2.BorderBrush; v2.BorderThickness = new Avalonia.Thickness(1); v2.Background = cardBg ?? v2.Background; }
            if (d1 is not null) { d1.Text = "●"; d1.Foreground = accent ?? d1.Foreground; }
            if (d2 is not null) { d2.Text = "○"; d2.Foreground = mutedFg ?? d2.Foreground; }
            if (label is not null) label.Text = "선택: 버전 1 · 기존 ETA (v21)";
        }
        else
        {
            if (v2 is not null) { v2.BorderBrush = accent ?? v2.BorderBrush; v2.BorderThickness = new Avalonia.Thickness(2); v2.Background = infoBg ?? v2.Background; }
            if (v1 is not null) { v1.BorderBrush = inkHair ?? v1.BorderBrush; v1.BorderThickness = new Avalonia.Thickness(1); v1.Background = cardBg ?? v1.Background; }
            if (d2 is not null) { d2.Text = "●"; d2.Foreground = accent ?? d2.Foreground; }
            if (d1 is not null) { d1.Text = "○"; d1.Foreground = mutedFg ?? d1.Foreground; }
            if (label is not null) label.Text = "선택: 버전 2 · Rebuild";
        }
    }
}
