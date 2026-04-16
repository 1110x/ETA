using Avalonia;
using Avalonia.Animation;
using Avalonia.Animation.Easings;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
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
        if (App.EnableLogging)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath, line + "\n");
            }
            catch { }
        }
    }

    public Login()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Log("========== 로그인 창 시작 ==========");

        txtEmail.KeyDown    += (_, e) => { if (e.Key == Key.Enter) txtPassword.Focus(); };
        txtPassword.KeyDown += (_, e) => { if (e.Key == Key.Enter) DoLogin(); };
    }

    // ── DB 선택 ComboBox ──────────────────────────────────────────────────────
    private record DbPeriodItem(string DbName, string Label)
    {
        public override string ToString() => Label;
    }

    private bool _dbSelectionChanging = false;

    private void CmbDb_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_dbSelectionChanging) return;
        if (cmbDb?.SelectedItem is DbPeriodItem item)
        {
            DbPathHelper.SetActiveDb(item.DbName);
            Log($"[DB선택] → {item.DbName}");
        }
    }

    private bool _isLoggingIn = false; // 중복 실행 방지

    private void Login_Click(object? sender, RoutedEventArgs e) => DoLogin();

    private async void DoLogin()
    {
        if (_isLoggingIn) { Log("[DoLogin] 중복 호출 차단"); return; }
        _isLoggingIn = true;
        pbLoginLoading.IsVisible = true;

        try
        {
        Log("---------- 로그인 시도 ----------");

        // MariaDB 연결 문자열 없으면 차단
        if (string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
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

        // To Do 승인 동기화 — 백그라운드 실행 (Microsoft API 지연이 로그인을 차단하지 않도록)
        _ = Task.Run(async () =>
        {
            try { await ETA.Services.SERVICE1.TodoService.SyncApprovalStatusAsync(); }
            catch { }
        });

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
        finally { _isLoggingIn = false; pbLoginLoading.IsVisible = false; }
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
            // ── Step 1: DB 연결 확인 ─────────────────────────────────────────
            SetProgress("DB 연결 중...", 0);
            await Task.Run(() =>
            {
                using var conn = DbConnectionFactory.CreateConnection();
                conn.Open();
            });
            SetProgress("DB 연결 완료", 14);
            await Task.Delay(120);

            // ── Step 2: 계정 DB 초기화 ───────────────────────────────────────
            SetProgress("계정 DB 초기화 중...", 14);
            await Task.Run(() =>
            {
                try { AgentService.MigrateAccountColumns(); }
                catch (Exception ex) { Log($"[Init] MigrateAccountColumns 실패: {ex.Message}"); }
            });
            SetProgress("계정 DB 초기화 완료", 28);
            await Task.Delay(120);

            // ── Step 3: 비밀번호 초기화 확인 ─────────────────────────────────
            SetProgress("비밀번호 초기화 확인 중...", 28);
            await Task.Run(() =>
            {
                try { AgentService.MigrateInitialPasswords(); }
                catch (Exception ex) { Log($"[Init] MigrateInitialPasswords 실패: {ex.Message}"); }
            });
            SetProgress("비밀번호 확인 완료", 42);
            await Task.Delay(120);

            // ── Step 4: 사진 데이터 동기화 ───────────────────────────────────
            SetProgress("사진 데이터 동기화 중...", 42);
            await Task.Run(() =>
            {
                try { AgentService.SyncAllPhotosToDb(); }
                catch (Exception ex) { Log($"[Init] SyncAllPhotosToDb 실패: {ex.Message}"); }
            });
            SetProgress("사진 동기화 완료", 56);
            await Task.Delay(120);

            // ── Step 5: Phase 1 DB 마이그레이션 (테이블명 통일) ──────────────────
            SetProgress("DB 마이그레이션 중... (Phase 1)", 60);
            try
            {
                await Task.Run(() => DbPhase1Migration.ExecutePhase1());
                SetProgress("DB 마이그레이션 완료 (Phase 1)", 65);
            }
            catch (Exception ex)
            {
                Log($"[Init] DbPhase1Migration 실패: {ex}");
                SetProgress($"⚠ DB 마이그레이션 실패: {ex.Message}", 65);
            }
            await Task.Delay(100);

            // ── Step 6: 처리시설 / 폐수 테이블 마이그레이션 ──────────────────
            SetProgress("테이블 초기화 중...", 68);
            try
            {
                await Task.Run(() => FacilityDbMigration.EnsureTables());
                SetProgress("테이블 준비 완료", 72);
            }
            catch (Exception ex)
            {
                Log($"[Init] FacilityDbMigration 실패: {ex}");
                SetProgress($"⚠ 테이블 초기화 실패: {ex.Message}", 72);
            }
            await Task.Delay(120);

            // ── Step 6: Phase 2 xlsm 데이터 마이그레이션 ────────────────────
            SetProgress("Phase 2 데이터 마이그레이션 중...", 75);
            try
            {
                await Task.Run(() => XlsmDataMigration.ExecutePhase2());
                SetProgress("Phase 2 마이그레이션 완료", 80);
            }
            catch (Exception ex)
            {
                Log($"[Init] Phase 2 마이그레이션 실패: {ex}");
                SetProgress($"⚠ Phase 2 마이그레이션 실패: {ex.Message}", 80);
            }
            await Task.Delay(100);

            // ── Step 7: 견적 테이블 초기화 ───────────────────────────────────
            SetProgress("견적 데이터 준비 중...", 82);
            try
            {
                await Task.Run(() => QuotationService.EnsureQuotationIssueTable());
                SetProgress("데이터 준비 완료", 88);
            }
            catch (Exception ex)
            {
                Log($"[Init] QuotationService 실패: {ex}");
                SetProgress($"⚠ 견적 초기화 실패: {ex.Message}", 88);
            }
            await Task.Delay(120);

            // ── Step 8: 처리시설 오늘 측정결과 자동 생성 ──────────────────────
            SetProgress("처리시설 측정결과 준비 중...", 88);
            try
            {
                int gen = await Task.Run(() => FacilityResultService.EnsureTodayMeasurementResults());
                SetProgress($"처리시설 준비 완료 ({gen}건)", 94);
                Log($"[Init] EnsureTodayMeasurementResults: {gen}건");
            }
            catch (Exception ex)
            {
                Log($"[Init] EnsureTodayMeasurementResults 실패: {ex.Message}");
                SetProgress("처리시설 준비 완료", 94);
            }
            await Task.Delay(120);

            // ── Step 8: 완료 ─────────────────────────────────────────────────
            SetProgress("시스템 준비 완료", 96);
            await Task.Delay(250);
            SetProgress("✓ 완료!", 100);
            await Task.Delay(200);
        }
        catch (Exception ex)
        {
            Log($"[Init] DB연결 오류: {ex}");
            SetProgress($"⚠ DB 연결 실패: {ex.Message}", 100);
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

    private async void OnLoaded(object? sender, RoutedEventArgs e)
    {
        // DB 선택 ComboBox 초기화 (가볍게 즉시 표시)
        _dbSelectionChanging = true;
        if (cmbDb != null)
        {
            cmbDb.ItemsSource = new[] { new DbPeriodItem(DbPathHelper.ConfigDbName, $"{DbPathHelper.ConfigDbName}  ★ 기존") };
            cmbDb.SelectedIndex = 0;
        }
        _dbSelectionChanging = false;
        // ETAS* DB 목록은 백그라운드에서 비동기 추가 (느려도 로그인 차단 안 함)
        _ = Task.Run(() =>
        {
            try
            {
                var etasDbs = DbRotationService.EnsureAndGetDbs();
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    if (cmbDb == null) return;
                    _dbSelectionChanging = true;
                    var items = new List<DbPeriodItem>
                        { new(DbPathHelper.ConfigDbName, $"{DbPathHelper.ConfigDbName}  ★ 기존") };
                    items.AddRange(etasDbs.Select(db =>
                        new DbPeriodItem(db, DbRotationService.GetLabelFromDbName(db))));
                    cmbDb.ItemsSource = items;
                    if (cmbDb.SelectedIndex < 0) cmbDb.SelectedIndex = 0;
                    _dbSelectionChanging = false;
                });
            }
            catch { /* NAS 연결 실패 시 무시 — eta_db로 계속 진행 */ }
        });

        // 로그인 폼 바로 표시 (시스템 초기화는 로그인 성공 후 진행)
        startupOverlay.IsVisible = false;
        loginForm.IsVisible = true;

        // 개발용: 자동 로그인
        if (txtEmail != null) txtEmail.Text = "201000308";
        if (txtPassword != null) txtPassword.Text = "1212xx!!AA";

        // 약간의 딜레이 후 자동 로그인
        await Task.Delay(500);
        DoLogin();
    }

    private void SetStartupProgress(string status, int value)
    {
        if (txtStartupStatus != null) txtStartupStatus.Text = status;
        if (pbStartup        != null) pbStartup.Value       = value;
        if (txtStartupPct    != null) txtStartupPct.Text    = $"{value}%";
        Log($"[Startup] {value}% — {status}");
    }

    // ── 이스터에그 비활성화 ────────────────────────────────────────────────
    private int _titleClickCount;
    private DateTime _titleLastClick = DateTime.MinValue;

    private void Title_Click(object? sender, PointerPressedEventArgs e)
    {
        // 관리자 자동 로그인 이스터에그 제거됨
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
