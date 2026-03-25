using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Services;
using System;
using System.IO;
using System.Diagnostics;

namespace ETA.Views;

/// <summary>
/// 비밀번호 변경 팝업.
/// isForced=true  → 최초 강제 변경 모드 (취소 불가, 변경 전 로그인 불가)
/// isForced=false → 일반 변경 모드 (취소 가능)
/// </summary>
public partial class ChangePasswordWindow : Window
{
    private readonly string _employeeId;
    private readonly bool   _isForced;

    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "eta_sync.log");

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

    /// <summary>변경 완료 여부 (부모 창에서 확인)</summary>
    public bool IsChanged { get; private set; } = false;

    // Avalonia XAML 로더용 기본 생성자 (AVLN3001 경고 제거)
    public ChangePasswordWindow() : this(string.Empty, false) { }

    public ChangePasswordWindow(string employeeId, bool isForced = false)
    {
        InitializeComponent();

        _employeeId = employeeId;
        _isForced   = isForced;

        Log($"[ChangePasswordWindow] 열림 - 사번={employeeId}, isForced={isForced}");

        // UI 모드 설정
        if (isForced)
        {
            txtTitle.Text       = "초기 비밀번호 변경 필요";
            txtSubtitle.Text    = "보안을 위해 비밀번호를 변경해야\n로그인할 수 있습니다.";
            bannerInitial.IsVisible = true;
            btnCancel.IsVisible = false;
        }

        txtNew.TextChanged += (_, _) => UpdateStrengthBar(txtNew.Text ?? "");

        txtCurrent.KeyDown += (_, e) => { if (e.Key == Avalonia.Input.Key.Enter) txtNew.Focus(); };
        txtNew.KeyDown     += (_, e) => { if (e.Key == Avalonia.Input.Key.Enter) txtConfirm.Focus(); };
        txtConfirm.KeyDown += (_, e) => { if (e.Key == Avalonia.Input.Key.Enter) DoChange(); };
    }

    // ── 변경 실행 ──────────────────────────────────────────────────────────
    private void Confirm_Click(object? sender, RoutedEventArgs e) => DoChange();

    private void DoChange()
    {
        Log("[DoChange] 변경하기 클릭");
        HideError();

        var current = txtCurrent.Text ?? "";
        var newPw   = txtNew.Text     ?? "";
        var confirm = txtConfirm.Text ?? "";

        Log($"[DoChange] 현재PW 길이={current.Length} / 새PW 길이={newPw.Length} / 확인PW 길이={confirm.Length}");

        if (string.IsNullOrEmpty(current))
        {
            ShowError("현재 비밀번호를 입력해주세요.");
            txtCurrent.Focus();
            return;
        }

        Log("[DoChange] ValidatePasswordRule 호출");
        var (valid, validMsg) = AgentService.ValidatePasswordRule(newPw);
        Log($"[DoChange] ValidatePasswordRule 결과: valid={valid}, msg='{validMsg}'");
        if (!valid)
        {
            ShowError(validMsg);
            txtNew.Focus();
            return;
        }

        if (newPw != confirm)
        {
            ShowError("새 비밀번호가 일치하지 않습니다.");
            txtConfirm.Focus();
            return;
        }

        if (current == newPw)
        {
            ShowError("새 비밀번호가 현재 비밀번호와 동일합니다.");
            txtNew.Focus();
            return;
        }

        Log($"[DoChange] AgentService.ChangePassword 호출 - 사번={_employeeId}");
        try
        {
            var result = AgentService.ChangePassword(_employeeId, current, newPw);
            Log($"[DoChange] ChangePassword 결과: success={result.success}, message='{result.message}'");

            if (!result.success)
            {
                ShowError(result.message);
                txtCurrent.Text = "";
                txtCurrent.Focus();
                return;
            }

            Log("[DoChange] 비밀번호 변경 성공 → 창 닫기");
            IsChanged = true;
            Close();
        }
        catch (Exception ex)
        {
            Log($"[DoChange] 예외 발생: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            ShowError($"오류가 발생했습니다: {ex.Message}");
        }
    }

    // ── 취소 ───────────────────────────────────────────────────────────────
    private void Cancel_Click(object? sender, RoutedEventArgs e)
    {
        if (_isForced) return;   // 강제 모드에서는 취소 버튼 동작 안 함
        IsChanged = false;
        Close();
    }

    // ── 강제 모드에서 창 닫기 처리 ────────────────────────────────────────
    protected override void OnClosing(WindowClosingEventArgs e)
    {
        // X버튼(UserClosing)으로 닫는 경우 → 그냥 닫힘 (IsChanged=false 유지)
        // 앱 종료 / 오너 창 닫힘 → 항상 허용
        // 프로그래밍적 Close() 호출(IsChanged=true 이후) → 허용
        base.OnClosing(e);
    }

    // ── 강도 표시 바 ───────────────────────────────────────────────────────
    private void UpdateStrengthBar(string pw)
    {
        int score = 0;
        if (pw.Length >= 6)  score++;
        if (pw.Length >= 10) score++;

        bool hasUpper   = false;
        bool hasSpecial = false;
        foreach (var c in pw)
        {
            if (char.IsUpper(c)) hasUpper   = true;
            if (!char.IsLetterOrDigit(c)) hasSpecial = true;
        }
        if (hasUpper)   score++;
        if (hasSpecial) score++;

        // score: 0~1 = 약함 / 2~3 = 보통 / 4 = 강함
        const string weak   = "#EF4444";
        const string medium = "#F59E0B";
        const string strong = "#10B981";
        const string empty  = "#E2E8F0";

        if (pw.Length == 0)
        {
            barWeak.Background   = Avalonia.Media.Brush.Parse(empty);
            barMedium.Background = Avalonia.Media.Brush.Parse(empty);
            barStrong.Background = Avalonia.Media.Brush.Parse(empty);
            txtStrength.Text     = "";
        }
        else if (score <= 1)
        {
            barWeak.Background   = Avalonia.Media.Brush.Parse(weak);
            barMedium.Background = Avalonia.Media.Brush.Parse(empty);
            barStrong.Background = Avalonia.Media.Brush.Parse(empty);
            txtStrength.Text     = "약함";
            txtStrength.Foreground = Avalonia.Media.Brush.Parse(weak);
        }
        else if (score <= 3)
        {
            barWeak.Background   = Avalonia.Media.Brush.Parse(medium);
            barMedium.Background = Avalonia.Media.Brush.Parse(medium);
            barStrong.Background = Avalonia.Media.Brush.Parse(empty);
            txtStrength.Text     = "보통";
            txtStrength.Foreground = Avalonia.Media.Brush.Parse(medium);
        }
        else
        {
            barWeak.Background   = Avalonia.Media.Brush.Parse(strong);
            barMedium.Background = Avalonia.Media.Brush.Parse(strong);
            barStrong.Background = Avalonia.Media.Brush.Parse(strong);
            txtStrength.Text     = "강함";
            txtStrength.Foreground = Avalonia.Media.Brush.Parse(strong);
        }
    }

    private void ShowError(string msg)
    {
        txtError.Text      = msg;
        txtError.IsVisible = true;
    }
    private void HideError() => txtError.IsVisible = false;
}
