using System;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record HomeStats(int 수질, int 폐수, int 처리시설, int 전체);

public static class HomeStatsService
{
    public static HomeStats GetToday()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        int wq = CountToday(conn, "수질분석센터_결과", "채취일자");
        int wa = CountToday(conn, "비용부담금_결과",   "채취일자");
        int fc = CountToday(conn, "처리시설_결과",     "채취일자");

        return new HomeStats(wq, wa, fc, wq + wa + fc);
    }

    private static int CountToday(System.Data.Common.DbConnection conn, string table, string dateCol)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT COUNT(*) FROM `{table}` WHERE DATE(`{dateCol}`) = CURDATE()";
            var v = cmd.ExecuteScalar();
            return v is null or DBNull ? 0 : Convert.ToInt32(v);
        }
        catch { return 0; }
    }
}
