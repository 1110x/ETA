using Avalonia;
using Avalonia.Animation;
using Avalonia.Animation.Easings;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using System;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views;

public partial class Login : Window
{
    // ── 로그 경로 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Logs", "Logs/eta_sync.log");

    private static void Log(string msg)
    {
        string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
        Debug.WriteLine(line);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath, line + "\n");
        }
        catch { }
    }

    public Login()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Log("========== 로그인 창 시작 ==========");

        try
        {
            Log("[Init] MigrateAccountColumns 시작");
            AgentService.MigrateAccountColumns();
            Log("[Init] MigrateAccountColumns 완료");
        }
        catch (Exception ex) { Log($"[Init] MigrateAccountColumns 실패: {ex}"); }

        try
        {
            Log("[Init] MigrateInitialPasswords 시작");
            AgentService.MigrateInitialPasswords();
            Log("[Init] MigrateInitialPasswords 완료");
        }
        catch (Exception ex) { Log($"[Init] MigrateInitialPasswords 실패: {ex}"); }

        try
        {
            Log("[Init] SyncAllPhotosToDb 시작");
            AgentService.SyncAllPhotosToDb();
            Log("[Init] SyncAllPhotosToDb 완료");
        }
        catch (Exception ex) { Log($"[Init] SyncAllPhotosToDb 실패: {ex}"); }



        txtEmail.KeyDown    += (_, e) => { if (e.Key == Key.Enter) txtPassword.Focus(); };
        txtPassword.KeyDown += (_, e) => { if (e.Key == Key.Enter) DoLogin(); };
    }

    private bool _isLoggingIn = false; // 중복 실행 방지

    private void Login_Click(object? sender, RoutedEventArgs e) => DoLogin();

    private async void DoLogin()
    {
        if (_isLoggingIn) { Log("[DoLogin] 중복 호출 차단"); return; }
        _isLoggingIn = true;

        try
        {
        Log("---------- 로그인 시도 ----------");
        Log($"[DoLogin] DB 모드: {(DbConnectionFactory.UseMariaDb ? "MariaDB" : "SQLite")}");

        // MariaDB 모드인데 연결 문자열이 없으면 차단
        if (DbConnectionFactory.UseMariaDb && string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
        {
            ShowError("서버 DB 연결 정보가 없습니다. appsettings.json을 확인하세요.");
            return;
        }

        string empId    = txtEmail?.Text?.Trim() ?? "";
        string password = txtPassword?.Text      ?? "";

        Log($"[DoLogin] 사번 입력값: '{empId}'");
        Log($"[DoLogin] 비밀번호 길이: {password.Length}자 / 비어있음: {string.IsNullOrEmpty(password)}");
        Log($"[DoLogin] 비밀번호 입력 해시: {AgentService.HashPassword(password)[..8]}...");

        if (string.IsNullOrEmpty(empId))    { ShowError("사번을 입력해주세요.");    Log("[DoLogin] 중단: 사번 없음"); return; }
        if (string.IsNullOrEmpty(password)) { ShowError("비밀번호를 입력해주세요."); Log("[DoLogin] 중단: 비밀번호 없음"); return; }

        // To Do 승인 동기화
        Log("[DoLogin] TodoService.SyncApprovalStatusAsync 시작");
        try
        {
            await ETA.Services.SERVICE1.TodoService.SyncApprovalStatusAsync();
            Log("[DoLogin] TodoService.SyncApprovalStatusAsync 완료");
        }
        catch (Exception ex)
        {
            Log($"[DoLogin] Sync 실패: {ex.Message}\n{ex.StackTrace}");
        }

        // DB 상태 덤프
        Log("[DoLogin] DB 계정 상태 확인 시작");
        AgentService.DebugCheckAccount(empId);

        // 로그인 검증
        Log("[DoLogin] ValidateLogin 호출");
        var (success, message, mustChangePw) = AgentService.ValidateLogin(empId, password);
        Log($"[DoLogin] ValidateLogin 결과 → success={success}, mustChangePw={mustChangePw}, message='{message}'");

        if (!success)
        {
            Log($"[DoLogin] 로그인 실패: {message}");
            ShowError(message);
            if (txtPassword != null) { txtPassword.Text = ""; txtPassword.Focus(); }
            return;
        }

        // 최초 비밀번호 강제 변경
        if (mustChangePw)
        {
            Log("[DoLogin] 최초 비밀번호 변경 팝업 표시");
            var dlg = new ChangePasswordWindow(empId, isForced: true);
            await dlg.ShowDialog(this);

            if (!dlg.IsChanged)
            {
                Log("[DoLogin] 비밀번호 변경 창 닫음 → 로그인 창으로 복귀");
                ShowError("💡 초기 비밀번호(123456) 변경 후 로그인해주세요.");
                if (txtPassword != null) txtPassword.Text = "";
                return;
            }
            Log("[DoLogin] 비밀번호 변경 완료");
        }

        Log("[DoLogin] 로그인 성공 → Lottie 재생");
        await AnimateLoginFormOut();
        PlayLottieAndNavigate(empId);
        }
        catch (Exception ex)
        {
            Log($"[DoLogin] 예외 발생: {ex.GetType().Name}: {ex.Message}");
            ShowError($"오류: {ex.Message}");
        }
        finally { _isLoggingIn = false; }
    }

    private async void PlayLottieAndNavigate(string empId)
    {
        // Lottie 재생 시작 (RepeatCount=-1 → 완료까지 반복 재생)
        try
        {
            var path = lottieView.Path;
            lottieView.AutoPlay = true;
            lottieView.Path = null;
            lottieView.Path = path;
            Log("[Lottie] 재생 시작");
        }
        catch (Exception ex) { Log($"[Lottie] 시작 오류: {ex.Message}"); }

        // 진행 오버레이 표시
        connectOverlay.IsVisible = true;

        try
        {
            // ── Step 1: DB 연결 확인 (서버 지연 시 Lottie 계속 반복) ──────────
            SetProgress("DB 연결 중...", 0);
            await Task.Run(() =>
            {
                using var conn = DbConnectionFactory.CreateConnection();
                conn.Open();
            });
            SetProgress("DB 연결 완료", 22);
            await Task.Delay(150);

            // ── Step 2: 처리시설 / 폐수 테이블 마이그레이션 ───────────────────
            SetProgress("테이블 초기화 중...", 22);
            await Task.Run(() => FacilityDbMigration.EnsureTables());
            SetProgress("테이블 준비 완료", 55);
            await Task.Delay(150);

            // ── Step 3: 견적 테이블 초기화 ────────────────────────────────────
            SetProgress("견적 데이터 준비 중...", 55);
            await Task.Run(() => QuotationService.EnsureQuotationIssueTable());
            SetProgress("데이터 준비 완료", 82);
            await Task.Delay(150);

            // ── Step 4: 완료 ──────────────────────────────────────────────────
            SetProgress("시스템 준비 완료", 96);
            await Task.Delay(300);
            SetProgress("✓ 완료!", 100);
            await Task.Delay(250);
        }
        catch (Exception ex)
        {
            Log($"[Init] 초기화 오류: {ex.Message}");
            SetProgress("⚠ 일부 항목 로드 실패 — 계속 진행", 100);
            await Task.Delay(700);
        }

        // Lottie 중단 후 메인 페이지 이동
        try { lottieView.AutoPlay = false; lottieView.Path = null; }
        catch { }

        CurrentUserManager.Instance.SetCurrentUser(empId);
        MainPage.CurrentEmployeeId = empId;

        var main = new MainPage();
        // 로그인 창과 같은 위치·크기로 열기 (이후 애니메이션 확장)
        main.WindowStartupLocation = WindowStartupLocation.Manual;
        main.Position    = this.Position;
        main.Width       = this.Width;
        main.Height      = this.Height;
        main.WindowState = this.WindowState;
        main.AnimateExpand(this.Width, this.Height, this.Position);
        main.Show();
        Close();
    }

    private void SetProgress(string status, int value)
    {
        txtConnectStatus.Text = status;
        pbConnect.Value       = value;
        txtConnectPct.Text    = $"{value}%";
        Log($"[Progress] {value}% — {status}");
    }

    private async void ChangePassword_Click(object? sender, RoutedEventArgs e)
    {
        string empId = txtEmail?.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(empId)) { ShowError("사번을 먼저 입력해주세요."); return; }

        var dlg = new ChangePasswordWindow(empId, isForced: false);
        await dlg.ShowDialog(this);

        if (dlg.IsChanged)
        {
            if (txtPassword != null) { txtPassword.Text = ""; txtPassword.Focus(); }
            if (txtError != null)
            {
                txtError.Foreground = Avalonia.Media.Brush.Parse("#10B981");
                txtError.Text       = "✓ 비밀번호가 변경되었습니다. 새 비밀번호로 로그인해주세요.";
                txtError.IsVisible  = true;
            }
        }
    }

    private void SignUp_Click(object? sender, PointerPressedEventArgs e)
        => new SignUpWindow().ShowDialog(this);

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        UpdateDbModeVisual();

        // MariaDB 연결 문자열이 없으면 서버 토글 비활성화
        if (string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
        {
            if (toggleDbMode != null) toggleDbMode.IsEnabled = false;
            if (txtDbStatus  != null) txtDbStatus.Text = "appsettings.json 없음";
            if (txtDbMode    != null) txtDbMode.Foreground = Avalonia.Media.Brush.Parse("#9CA3AF");
            Log("[DB모드] MariaDbConnectionString 미설정 → 서버 토글 비활성화");
        }
    }

    // ── DB 모드 토글 (커스텀 Border 토글) ─────────────────────────────────────
    private void DbModeToggle_Click(object? sender, PointerPressedEventArgs e)
    {
        bool useMariaDb = !DbConnectionFactory.UseMariaDb; // 현재 상태 반전

        // MariaDB 선택인데 연결 문자열이 없으면 차단
        if (useMariaDb && string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
        {
            ShowError("서버 DB 연결 설정(appsettings.json)이 없습니다.");
            return;
        }

        DbConnectionFactory.UseMariaDb = useMariaDb;
        UpdateDbModeVisual();
        Log($"[DB모드] {(useMariaDb ? "서버 DB (MariaDB)" : "로컬 DB (SQLite)")} 선택");
    }

    private void UpdateDbModeVisual()
    {
        bool on = DbConnectionFactory.UseMariaDb;

        // 트랙 색상
        if (toggleDbMode != null)
            toggleDbMode.Background = Avalonia.Media.Brush.Parse(on ? "#6366F1" : "#D1D5DB");

        // 녹 위치 (ON: 오른쪽, OFF: 왼쪽)
        if (toggleKnob != null)
            toggleKnob.Margin = new Avalonia.Thickness(on ? 25 : 3, 0, 0, 0);

        // 텍스트
        if (txtDbMode   != null) txtDbMode.Text   = on ? "서버 DB"  : "로컬 DB";
        if (txtDbStatus != null) txtDbStatus.Text = on ? "CHUNGHA DB (온라인)" : "SQLite (오프라인)";
    }

    // ── 이스터에그: 타이틀 5회 클릭 → 관리자 자동 로그인 ──────────────────
    private int _titleClickCount;
    private DateTime _titleLastClick = DateTime.MinValue;

    private async void Title_Click(object? sender, PointerPressedEventArgs e)
    {
        var now = DateTime.Now;
        if ((now - _titleLastClick).TotalSeconds > 3) _titleClickCount = 0;
        _titleLastClick = now;
        _titleClickCount++;

        if (_titleClickCount >= 5)
        {
            _titleClickCount = 0;
            Log("[EasterEgg] 관리자 자동 로그인");
            await AnimateLoginFormOut();
            PlayLottieAndNavigate("201000308");
        }
    }

    // ── 로그인 폼 페이드아웃 + 슬라이드업 애니메이션 ──────────────────────
    private async Task AnimateLoginFormOut()
    {
        var animation = new Animation
        {
            Duration = TimeSpan.FromMilliseconds(350),
            Easing = new CubicEaseIn(),
            FillMode = FillMode.Forward,
            Children =
            {
                new KeyFrame
                {
                    Cue = new Cue(0),
                    Setters =
                    {
                        new Setter(OpacityProperty, 1.0),
                        new Setter(TranslateTransform.YProperty, 0.0)
                    }
                },
                new KeyFrame
                {
                    Cue = new Cue(1),
                    Setters =
                    {
                        new Setter(OpacityProperty, 0.0),
                        new Setter(TranslateTransform.YProperty, -30.0)
                    }
                }
            }
        };
        await animation.RunAsync(loginForm);
        loginForm.IsVisible = false;
    }

    private async void ShowError(string msg)
    {
        if (txtError == null) return;
        txtError.Foreground = Brush.Parse("#EF4444");
        txtError.Opacity = 0;
        txtError.Text = msg;
        txtError.IsVisible = true;
        await Task.Delay(16); // 1프레임 대기 후 트랜지션 발동
        txtError.Opacity = 1;
    }
}
