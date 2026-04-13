using System;
using System.IO;
using System.Text.Json;
using ETA.Services.Common;

namespace ETA.Services.Common;

/// <summary>
/// 현재 로그인한 사용자 정보 관리
/// </summary>
public class CurrentUserManager
{
    private static CurrentUserManager? _instance;
    private string _currentUserId   = Environment.UserName ?? "DefaultUser";
    private string _currentDept     = "";
    private static readonly object _lockObject = new object();

    public string CurrentUserId   => _currentUserId;
    /// <summary>로그인한 사용자의 부서명 (비어있으면 전체 접근)</summary>
    public string CurrentDepartment => _currentDept;

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
            _currentDept   = FetchDepartment(userId);
        }
    }

    private static string FetchDepartment(string userId)
    {
        try
        {
            using var conn = ETA.Services.Common.DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `부서` FROM `Agent` WHERE `사번` = @id";
            cmd.Parameters.AddWithValue("@id", userId);
            return cmd.ExecuteScalar()?.ToString()?.Trim() ?? "";
        }
        catch { return ""; }
    }

    public void ResetToWindowsUser()
    {
        _currentUserId = Environment.UserName ?? "DefaultUser";
    }
}
