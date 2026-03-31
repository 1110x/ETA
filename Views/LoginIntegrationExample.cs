using Avalonia.Controls;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System.Threading.Tasks;

namespace ETA.Views;

/// <summary>
/// 로그인 성공 후 MainPage에 사용자 정보를 설정하는 예제
/// 
/// 사용 방법:
/// 1. 현재 Login.axaml.cs에서 로그인 성공 후 MainPage를 표시할 때 이 메서드를 호출
/// 2. 로그인 사용자의 ID를 전달하면 창 위치 저장/복원이 해당 사용자별로 관리됨
/// </summary>
public static class LoginIntegrationExample
{
    /// <summary>
    /// 로그인 완료 후 MainPage 초기화 (예제)
    /// </summary>
    public static void InitializeMainPageAfterLogin(string employeeId, string employeeName)
    {
        // 현재 사용자 설정
        CurrentUserManager.Instance.SetCurrentUser(employeeId);
        
        // MainPage 생성
        var mainPage = new MainPage
        {
            Title = $"ETA - {employeeName} ({employeeId})",
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            Width = 1400,
            Height = 800
        };
        
        // 👉 중요: 사용자별 레이아웃 복원을 위해 이 메서드 호출
        mainPage.UpdateCurrentUser(employeeId);
        
        // MainPage 표시
        mainPage.Show();
    }

    /// <summary>
    /// 로그인 창이 닫힌 후 메인 윈도우를 표시하는 예제
    /// (Login.axaml.cs의 로그인 성공 핸들러에 통합)
    /// </summary>
    public static async Task ShowMainPageAsync(Window? parent, string employeeId, string employeeName)
    {
        // 로그인 창 종료
        if (parent != null)
            parent.Close();
        
        // 짧은 딜레이 (UI 갱신용)
        await Task.Delay(500);
        
        // MainPage 표시
        InitializeMainPageAfterLogin(employeeId, employeeName);
    }
}

/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║                        LOGIN.AXAML.CS 통합 예제                           ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 * 
 * // 로그인 버튼 클릭 핸들러
 * private async void LoginButton_Click(object? sender, RoutedEventArgs e)
 * {
 *     string username = UsernameTextBox.Text ?? "";
 *     string password = PasswordTextBox.Text ?? "";
 *     
 *     // 로그인 검증
 *     var result = await AuthenticateUserAsync(username, password);
 *     
 *     if (result.IsSuccess)
 *     {
 *         // ✅ 로그인 성공 - MainPage 초기화 및 표시
 *         MainPage.CurrentEmployeeId = result.EmployeeId;
 *         await LoginIntegrationExample.ShowMainPageAsync(
 *             this as Window,
 *             result.EmployeeId,
 *             result.EmployeeName);
 *     }
 *     else
 *     {
 *         // ❌ 로그인 실패
 *         MessageBox.Show("로그인 실패: " + result.ErrorMessage);
 *     }
 * }
 * 
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║                    직접 호출 방법 (간단한 버전)                            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 * 
 * // 로그인 성공 후
 * MainPage mainPage = new MainPage();
 * mainPage.UpdateCurrentUser("emp_12345");  // 👈 이 한 줄이 중요!
 * mainPage.Show();
 * 
 * this.Close();  // 로그인 창 닫기
 */
