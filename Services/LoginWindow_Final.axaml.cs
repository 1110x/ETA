using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using LibVLCSharp.Avalonia;
using LibVLCSharp.Shared;
using System;
using System.Threading.Tasks;
using ETA.Services;

namespace ETA.Views;

public partial class LoginWindow : Window
{
    private LibVLC?      _libVLC;
    private MediaPlayer? _player;

    private readonly TextBox   _tbEmployeeId;
    private readonly TextBox   _tbPassword;
    private readonly TextBlock _tbError;
    private readonly Button    _btnLogin;

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    public LoginWindow()
    {
        Title             = "ETA Water Analysis Center";
        Width             = 1100;
        Height            = 700;
        CanResize         = false;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        SystemDecorations = SystemDecorations.None;

        // ── DB 마이그레이션 (앱 최초 실행 시 컬럼 추가) ───────────────────
        AgentService.MigrateAccountColumns();
        // ★ 기존 직원 초기 비밀번호 일괄 설정 (비밀번호 없는 approved → 123456)
        AgentService.MigrateInitialPasswords();

        _tbEmployeeId = MakeTb("사번 입력");
        _tbPassword   = MakeTb("비밀번호", isPassword: true);
        _tbError      = new TextBlock
        {
            FontFamily   = Font,
            FontSize     = 11,
            Foreground   = new SolidColorBrush(Color.Parse("#FF6B6B")),
            IsVisible    = false,
            TextWrapping = TextWrapping.Wrap,
            Margin       = new Thickness(0, 0, 0, 4),
        };
        _btnLogin = new Button
        {
            Content             = "시스템 접속",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Height              = 48,
            FontFamily          = Font,
            FontSize            = 14,
            Foreground          = Brushes.White,
            BorderThickness     = new Thickness(0),
            CornerRadius        = new CornerRadius(8),
            Background          = new SolidColorBrush(Color.Parse("#1a5c3a")),
        };
        _btnLogin.Click += BtnLogin_Click;

        // Enter 키 로그인
        _tbPassword.KeyDown += (_, e) => { if (e.Key == Key.Enter) DoLogin(); };
        _tbEmployeeId.KeyDown += (_, e) => { if (e.Key == Key.Enter) _tbPassword.Focus(); };

        // ── 회원가입 버튼 ─────────────────────────────────────────────────
        var btnSignUp = new Button
        {
            Content         = "신규 계정 신청",
            HorizontalAlignment = HorizontalAlignment.Right,
            Height          = 30,
            FontFamily      = Font,
            FontSize        = 11,
            Foreground      = new SolidColorBrush(Color.Parse("#66aaff")),
            Background      = Brushes.Transparent,
            BorderThickness = new Thickness(0),
        };
        btnSignUp.Click += (_, _) =>
        {
            var sw = new SignUpWindow();
            sw.ShowDialog(this);
        };

        // ── 종료 버튼 ─────────────────────────────────────────────────────
        var btnExit = new Button
        {
            Content         = "종료",
            HorizontalAlignment = HorizontalAlignment.Right,
            Height          = 28,
            FontFamily      = Font,
            FontSize        = 11,
            Foreground      = new SolidColorBrush(Color.Parse("#555555")),
            Background      = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Margin          = new Thickness(0, 8, 0, 0),
        };
        btnExit.Click += (_, _) => { StopVideo(); Close(); };

        // ── 로그인 카드 ───────────────────────────────────────────────────
        var card = new Border
        {
            Width               = 360,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment   = VerticalAlignment.Stretch,
            Background          = new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)),
            BorderBrush         = new SolidColorBrush(Color.FromArgb(35, 255, 255, 255)),
            BorderThickness     = new Thickness(1, 0, 0, 0),
            Child               = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin            = new Thickness(44, 0),
                Spacing           = 0,
                Children          =
                {
                    new Border
                    {
                        Width               = 52,
                        Height              = 52,
                        Background          = new SolidColorBrush(Color.FromArgb(35,255,255,255)),
                        CornerRadius        = new CornerRadius(14),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        Margin              = new Thickness(0, 0, 0, 28),
                        Child               = new TextBlock
                        {
                            Text = "💧", FontSize = 26,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment   = VerticalAlignment.Center,
                        }
                    },
                    new TextBlock
                    {
                        Text       = "로그인",
                        FontFamily = Font,
                        FontSize   = 26,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = Brushes.White,
                        Margin     = new Thickness(0, 0, 0, 4),
                    },
                    new TextBlock
                    {
                        Text       = "수질분석센터 시스템",
                        FontFamily = Font,
                        FontSize   = 11,
                        Foreground = new SolidColorBrush(Color.FromArgb(136, 255, 255, 255)),
                        Margin     = new Thickness(0, 0, 0, 32),
                    },
                    FieldLabel("사  번"),
                    WithMargin(_tbEmployeeId, 0, 0, 0, 16),
                    FieldLabel("비밀번호"),
                    WithMargin(_tbPassword, 0, 0, 0, 8),
                    _tbError,
                    WithMargin(_btnLogin, 0, 12, 0, 0),
                    btnSignUp,
                    btnExit,
                }
            }
        };

        // ── 왼쪽 타이틀 ──────────────────────────────────────────────────
        var title = new StackPanel
        {
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment   = VerticalAlignment.Bottom,
            Margin              = new Thickness(60, 0, 0, 70),
            Spacing             = 8,
            Children            =
            {
                new TextBlock
                {
                    Text       = "ETA",
                    FontSize   = 72,
                    FontWeight = FontWeight.Bold,
                    Foreground = new SolidColorBrush(Color.Parse("#EAFFF4")),
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                },
                new TextBlock
                {
                    Text       = "Water Analysis Center",
                    FontSize   = 16,
                    Foreground = new SolidColorBrush(Color.FromArgb(136, 234, 255, 244)),
                    FontFamily = Font,
                },
            }
        };

        // ── 드래그 영역 ───────────────────────────────────────────────────
        var dragArea = new Border
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Top,
            Height              = 40,
            Background          = Brushes.Transparent,
        };
        dragArea.PointerPressed += (_, e) =>
        {
            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                BeginMoveDrag(e);
        };

        // ── 비디오 + 오버레이 ─────────────────────────────────────────────
        var videoView = new VideoView
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Stretch,
            Name                = "VideoBackground",
        };

        var overlay = new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(160, 0, 0, 0)),
        };

        var grid = new Grid();
        grid.Children.Add(videoView);
        grid.Children.Add(overlay);
        grid.Children.Add(title);
        grid.Children.Add(card);
        grid.Children.Add(dragArea);

        Content = grid;

        // VideoView 참조 저장
        _videoView = videoView;
    }

    private VideoView? _videoView;

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        StartVideo();
    }

    private void StartVideo()
    {
        try
        {
            Core.Initialize();
            _libVLC = new LibVLC();
            _player = new MediaPlayer(_libVLC);

            if (_videoView != null)
                _videoView.MediaPlayer = _player;

            using var media = new Media(
                _libVLC,
                "avares://ETA/Assets/Videos/login_bg.mp4",
                FromType.FromLocation);
            media.AddOption(":input-repeat=65535");

            _player.Volume = 0;
            _player.Media  = media;
            _player.Play();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[LoginWindow] 영상 오류: {ex.Message}");
        }
    }

    // ── 로그인 처리 ───────────────────────────────────────────────────────
    private void BtnLogin_Click(object? sender, RoutedEventArgs e) => DoLogin();

    private async void DoLogin()
    {
        _tbError.IsVisible   = false;
        _btnLogin.IsEnabled  = false;
        _btnLogin.Content    = "확인 중...";

        var empId = _tbEmployeeId.Text?.Trim() ?? "";
        var pw    = _tbPassword.Text           ?? "";

        if (string.IsNullOrEmpty(empId))
        {
            ShowError("사번을 입력해주세요.");
            return;
        }

        // To Do 승인 상태 동기화 (백그라운드)
        try { await TodoService.SyncApprovalStatusAsync(); }
        catch { /* 네트워크 없어도 로그인 계속 */ }

        // DB 로그인 검증 (3-tuple: success, message, mustChangePw)
        var (success, message, mustChangePw) = AgentService.ValidateLogin(empId, pw);

        if (success)
        {
            // ★ 최초 비밀번호 강제 변경 (초기값 123456 사용자)
            if (mustChangePw)
            {
                var dlg = new ChangePasswordWindow(empId, isForced: true);
                await dlg.ShowDialog(this);

                if (!dlg.IsChanged)
                {
                    ShowError("비밀번호를 변경해야 로그인할 수 있습니다.");
                    return;
                }
            }

            StopVideo();
            var main = new MainPage();
            main.Show();
            Close();
        }
        else
        {
            ShowError(message);
        }
    }

    private void ShowError(string msg)
    {
        _tbError.Text       = msg;
        _tbError.IsVisible  = true;
        _btnLogin.IsEnabled = true;
        _btnLogin.Content   = "시스템 접속";
        _tbPassword.Text    = "";
    }

    private void StopVideo()
    {
        try { _player?.Stop(); _player?.Dispose(); _libVLC?.Dispose(); }
        catch { }
    }

    protected override void OnClosed(EventArgs e)
    {
        StopVideo();
        base.OnClosed(e);
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static TextBlock FieldLabel(string text) => new TextBlock
    {
        Text       = text,
        FontFamily = Font,
        FontSize   = 11,
        Foreground = new SolidColorBrush(Color.FromArgb(153, 255, 255, 255)),
        Margin     = new Thickness(0, 0, 0, 8),
    };

    private static Control WithMargin(Control c, double l, double t, double r, double b)
    {
        c.Margin = new Thickness(l, t, r, b);
        return c;
    }

    private static TextBox MakeTb(string hint, bool isPassword = false) => new TextBox
    {
        Watermark           = hint,
        FontFamily          = Font,
        FontSize            = 13,
        Background          = new SolidColorBrush(Color.FromArgb(35, 255, 255, 255)),
        Foreground          = Brushes.White,
        BorderBrush         = new SolidColorBrush(Color.FromArgb(51, 255, 255, 255)),
        BorderThickness     = new Thickness(1),
        CornerRadius        = new CornerRadius(8),
        Padding             = new Thickness(14, 10),
        PasswordChar        = isPassword ? '●' : '\0',
        HorizontalAlignment = HorizontalAlignment.Stretch,
        Height              = 44,
    };
}
