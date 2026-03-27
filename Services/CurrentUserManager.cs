using System;
using System.IO;
using System.Text.Json;

namespace ETA.Services;

/// <summary>
/// 현재 로그인한 사용자 정보 관리
/// </summary>
public class CurrentUserManager
{
    private static CurrentUserManager? _instance;
    private string _currentUserId = Environment.UserName ?? "DefaultUser";
    private static readonly object _lockObject = new object();

    public string CurrentUserId => _currentUserId;

    public static CurrentUserManager Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lockObject)
                {
                    _instance ??= new CurrentUserManager();
                }
            }
            return _instance;
        }
    }

    public void SetCurrentUser(string userId)
    {
        if (!string.IsNullOrWhiteSpace(userId))
        {
            _currentUserId = userId;
            System.Diagnostics.Debug.WriteLine($"[CurrentUserManager] 사용자 설정: {_currentUserId}");
        }
    }

    public void ResetToWindowsUser()
    {
        _currentUserId = Environment.UserName ?? "DefaultUser";
        System.Diagnostics.Debug.WriteLine($"[CurrentUserManager] Windows 사용자로 초기화: {_currentUserId}");
    }
}
