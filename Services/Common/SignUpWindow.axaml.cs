using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Threading.Tasks;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views;

public class SignUpWindow : Window
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private readonly TextBox  _tbName;
    private readonly TextBox  _tbEmployeeId;
    private readonly TextBox  _tbPassword;
    private readonly TextBox  _tbPasswordConfirm;
    private readonly ComboBox _cbDepartment;
    private readonly TextBlock _tbError;
    private readonly Button   _btnSignUp;

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    public SignUpWindow()
    {
        Title             = "ETA — 회원가입";
        Width             = 420;
        Height            = 560;
        CanResize         = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        SystemDecorations = SystemDecorations.BorderOnly;
        Background        = new SolidColorBrush(Color.Parse("#1e1f2e"));

        _tbName            = MakeTb("이름");
        _tbEmployeeId      = MakeTb("사번 (예: 2024001)");
        _tbPassword        = MakeTb("비밀번호", isPassword: true);
        _tbPasswordConfirm = MakeTb("비밀번호 확인", isPassword: true);
        _cbDepartment      = new ComboBox
        {
            ItemsSource   = new[] { "수질분석센터", "처리시설", "일반업무", "기타" },
            SelectedIndex = 0,
            FontFamily    = Font,
            FontSize      = 13,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Background    = new SolidColorBrush(Color.Parse("#2a2b3d")),
            Foreground    = new SolidColorBrush(Color.Parse("#e8eaf6")),
            BorderBrush   = new SolidColorBrush(Color.Parse("#3d3f5c")),
        };

        _tbError = new TextBlock
        {
            FontFamily  = Font,
            FontSize    = 11,
            Foreground  = new SolidColorBrush(Color.Parse("#FF6B6B")),
            TextWrapping = TextWrapping.Wrap,
            IsVisible   = false,
        };

        _btnSignUp = new Button
        {
            Content             = "가입 신청",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Height              = 46,
            FontFamily          = Font,
            FontSize            = 14,
            Foreground          = Brushes.White,
            BorderThickness     = new Thickness(0),
            CornerRadius        = new CornerRadius(8),
            Background          = new SolidColorBrush(Color.Parse("#1a5c3a")),
        };
        _btnSignUp.Click += BtnSignUp_Click;

        var cancelBtn = new Button
        {
            Content             = "취소",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Height              = 36,
            FontFamily          = Font,
            FontSize            = 12,
            Foreground          = new SolidColorBrush(Color.Parse("#8890b0")),
            BorderThickness     = new Thickness(0),
            Background          = Brushes.Transparent,
        };
        cancelBtn.Click += (_, _) => Close();

        Content = new Border
        {
            Padding    = new Thickness(36, 32),
            Background = new SolidColorBrush(Color.Parse("#1e1f2e")),
            Child      = new StackPanel
            {
                Spacing  = 14,
                Children =
                {
                    new TextBlock
                    {
                        Text       = "신규 계정 신청",
                        FontFamily = Font,
                        FontSize   = 22,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = new SolidColorBrush(Color.Parse("#e8eaf6")),
                        Margin     = new Thickness(0, 0, 0, 4),
                    },
                    new TextBlock
                    {
                        Text       = "관리자 승인 후 로그인이 가능합니다",
                        FontFamily = Font,
                        FontSize   = 11,
                        Foreground = new SolidColorBrush(Color.Parse("#8890b0")),
                        Margin     = new Thickness(0, 0, 0, 8),
                    },
                    LabeledField("이  름", _tbName),
                    LabeledField("사  번", _tbEmployeeId),
                    LabeledField("부  서", _cbDepartment),
                    LabeledField("비밀번호", _tbPassword),
                    LabeledField("비밀번호 확인", _tbPasswordConfirm),
                    _tbError,
                    _btnSignUp,
                    cancelBtn,
                }
            }
        };
    }

    private async void BtnSignUp_Click(object? sender, RoutedEventArgs e)
    {
        _tbError.IsVisible = false;

        var name     = _tbName.Text?.Trim()            ?? "";
        var empId    = _tbEmployeeId.Text?.Trim()      ?? "";
        var pw       = _tbPassword.Text               ?? "";
        var pwConfirm = _tbPasswordConfirm.Text       ?? "";
        var dept     = _cbDepartment.SelectedItem?.ToString() ?? "";

        if (string.IsNullOrEmpty(name))   { ShowError("이름을 입력해주세요."); return; }
        if (string.IsNullOrEmpty(empId))  { ShowError("사번을 입력해주세요."); return; }
        if (pw.Length < 4)                { ShowError("비밀번호는 4자 이상이어야 합니다."); return; }
        if (pw != pwConfirm)              { ShowError("비밀번호가 일치하지 않습니다."); return; }

        _btnSignUp.IsEnabled = false;
        _btnSignUp.Content   = "처리 중...";

        // 1. DB 저장 (pending)
        bool ok = AgentService.SignUp(name, empId, pw, dept);
        if (!ok)
        {
            ShowError("이미 등록된 사번입니다.");
            _btnSignUp.IsEnabled = true;
            _btnSignUp.Content   = "가입 신청";
            return;
        }

        // 2. Microsoft To Do 에 승인 요청 태스크 추가
        try
        {
            var taskId = await TodoService.AddApprovalTaskAsync(name, empId);
            if (taskId != null)
                AgentService.SaveTodoTaskId(empId, taskId);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SignUp] Todo 전송 실패: {ex.Message}");
            // Todo 실패해도 회원가입은 완료된 것으로 처리
        }

        // 3. 완료 안내
        var dlg = new Window
        {
            Title   = "신청 완료",
            Width   = 320,
            Height  = 180,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.Parse("#1e1f2e")),
        };
        var okBtn = new Button
        {
            Content = "확인", Width = 80,
            HorizontalAlignment = HorizontalAlignment.Center,
            Background = new SolidColorBrush(Color.Parse("#1a5c3a")),
            Foreground = Brushes.White, BorderThickness = new Thickness(0),
        };
        okBtn.Click += (_, _) => { dlg.Close(); Close(); };
        dlg.Content = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Spacing = 16, Margin = new Thickness(20),
            Children =
            {
                new TextBlock
                {
                    Text = "✅  가입 신청이 완료되었습니다.\n관리자 승인 후 로그인 가능합니다.",
                    FontFamily = Font, FontSize = 12,
                    Foreground = new SolidColorBrush(Color.Parse("#e8eaf6")),
                    TextWrapping = TextWrapping.Wrap,
                    TextAlignment = TextAlignment.Center,
                },
                okBtn
            }
        };
        await dlg.ShowDialog(this);
    }

    private void ShowError(string msg)
    {
        _tbError.Text      = msg;
        _tbError.IsVisible = true;
        _btnSignUp.IsEnabled = true;
        _btnSignUp.Content   = "가입 신청";
    }

    private static StackPanel LabeledField(string label, Control input)
    {
        input.HorizontalAlignment = HorizontalAlignment.Stretch;
        return new StackPanel
        {
            Spacing  = 4,
            Children =
            {
                new TextBlock
                {
                    Text       = label,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    FontSize   = 11,
                    Foreground = new SolidColorBrush(Color.Parse("#8890b0")),
                },
                input
            }
        };
    }

    private static TextBox MakeTb(string hint, bool isPassword = false) => new TextBox
    {
        Watermark       = hint,
        FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        FontSize        = 13,
        Background      = new SolidColorBrush(Color.Parse("#2a2b3d")),
        Foreground      = new SolidColorBrush(Color.Parse("#e8eaf6")),
        BorderBrush     = new SolidColorBrush(Color.Parse("#3d3f5c")),
        BorderThickness = new Thickness(1),
        CornerRadius    = new CornerRadius(6),
        Padding         = new Thickness(10, 8),
        PasswordChar    = isPassword ? '●' : '\0',
    };
}
