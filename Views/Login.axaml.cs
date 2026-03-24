using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform;
using System;
using ETA.Services;

namespace ETA.Views;


public partial class Login : Window
{
    public Login()
    {
        InitializeComponent();

        // DB 마이그레이션 (최초 실행 시 컬럼 추가)
        AgentService.MigrateAccountColumns();

        // Enter 키 → 로그인
        txtPassword.KeyDown += (_, e) => { if (e.Key == Key.Enter) DoLogin(); };
        txtEmail.KeyDown    += (_, e) => { if (e.Key == Key.Enter) txtPassword.Focus(); };
    }

    // ── 로그인 버튼 ───────────────────────────────────────────────────────────
    private void Login_Click(object? sender, RoutedEventArgs e) => DoLogin();

    private async void DoLogin()
    {
        string empId    = txtEmail?.Text?.Trim() ?? "";
        string password = txtPassword?.Text      ?? "";

        if (string.IsNullOrEmpty(empId))
        {
            ShowError("사번을 입력해주세요.");
            return;
        }
        if (string.IsNullOrEmpty(password))
        {
            ShowError("비밀번호를 입력해주세요.");
            return;
        }

        // ── To Do 승인 상태 동기화 (백그라운드, 실패해도 진행) ──────────────
        try { await ETA.Services.TodoService.SyncApprovalStatusAsync(); }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Login] TodoSync 실패: {ex.Message}");
        }

        // ── DB 로그인 검증 ────────────────────────────────────────────────────
        var (success, message) = AgentService.ValidateLogin(empId, password);

        if (success)
        {
            var main = new MainPage();
            main.Show();
            Close();
        }
        else
        {
            ShowError(message);
            txtPassword.Text = "";
            txtPassword.Focus();
        }
    }

    // ── 비밀번호 표시 토글 ────────────────────────────────────────────────────
    private void OnShowPasswordChanged(object? sender, RoutedEventArgs e)
    {
        if (txtPassword != null && tglShowPassword != null)
            txtPassword.PasswordChar = tglShowPassword.IsChecked == true ? '\0' : '*';
    }

    // ── 회원가입 클릭 ─────────────────────────────────────────────────────────
    private void SignUp_Click(object? sender, PointerPressedEventArgs e)
    {
        var sw = new SignUpWindow();
        sw.ShowDialog(this);
    }

    // ── 오류 메시지 표시 ──────────────────────────────────────────────────────
    private void ShowError(string msg)
    {
        if (txtError != null)
        {
            txtError.Text      = msg;
            txtError.IsVisible = true;
        }
    }
}
