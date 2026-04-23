using System;
using System.Security.Cryptography;
using System.Text;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public static class AuthService
{
    public static string HashPassword(string password)
    {
        using var sha = SHA256.Create();
        var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
        return Convert.ToHexString(bytes).ToLower();
    }

    public static (bool success, string message) ValidateLogin(string employeeId, string password)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `비밀번호`, `상태` FROM `Agent` WHERE `사번` = @id";
        var p = cmd.CreateParameter();
        p.ParameterName = "@id";
        p.Value = employeeId;
        cmd.Parameters.Add(p);

        using var r = cmd.ExecuteReader();
        if (!r.Read()) return (false, "등록되지 않은 사번입니다.");

        var dbPw = r.IsDBNull(0) ? "" : r.GetString(0);
        var status = r.IsDBNull(1) ? "" : r.GetString(1);

        switch (status)
        {
            case "approved": break;
            case "pending":  return (false, "관리자 승인 대기 중입니다.");
            case "rejected": return (false, "승인이 거부된 계정입니다.");
            default:         return (false, "승인되지 않은 계정입니다.");
        }

        if (string.IsNullOrEmpty(dbPw))
            return (false, "비밀번호가 설정되지 않은 계정입니다. 관리자에게 문의하세요.");

        if (dbPw != HashPassword(password))
            return (false, "비밀번호가 올바르지 않습니다.");

        return (true, "로그인 성공");
    }
}
