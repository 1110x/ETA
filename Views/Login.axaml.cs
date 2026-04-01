using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
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
        loginForm.IsVisible = false;
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
        try
        {
            // Path 재할당으로 AutoPlay=True 상태에서 재로드 → 재생 트리거
            var path = lottieView.Path;
            lottieView.AutoPlay = true;
            lottieView.Path = null;
            lottieView.Path = path;
            Log("[Lottie] 재생 시작");
            await Task.Delay(4000);
        }
        catch (Exception ex) { Log($"[Lottie] 오류: {ex.Message}"); }

        CurrentUserManager.Instance.SetCurrentUser(empId);
        MainPage.CurrentEmployeeId = empId;

        var main = new MainPage();
        main.Show();
        Close();
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

    private void ShowError(string msg)
    {
        if (txtError == null) return;
        txtError.Foreground = Avalonia.Media.Brush.Parse("#EF4444");
        txtError.Text       = msg;
        txtError.IsVisible  = true;
    }
}
