namespace ETA.Services;

public static class AdminService
{
    public static string[] GetAdminItems()
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