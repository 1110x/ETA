using Avalonia.Controls;

namespace ETA.Services;

public static class AdminServices
{
    public static string[] GetAdminMenu()
    {
        return new string[]
        {
            "사용자 관리",
            "권한 관리",
            "로그 확인",
            "시스템 설정"
        };
    }
}
