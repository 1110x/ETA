using System.Collections.Generic;
using System.Data.Common;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record TableStat(string Table, long Total, long Today, string DateCol);

public class ServerInfo
{
    public bool   IsConnected { get; set; }
    public string Host { get; set; } = "";
    public string Port { get; set; } = "";
    public string Database { get; set; } = "";
    public string User { get; set; } = "";
    public string Version { get; set; } = "";
    public List<TableStat> Tables { get; set; } = new();
}

public static class ServerInfoService
{
    private static readonly (string Table, string DateCol)[] _targets =
    {
        ("수질분석센터_결과", "채취일자"),
        ("비용부담금_결과",   "채취일자"),
        ("처리시설_결과",     "채취일자"),
        ("폐수배출업소",      ""),
        ("Agent",             ""),
        ("화합물별명",        ""),
    };

    public static ServerInfo GetInfo()
    {
        var info = new ServerInfo();

        var cs = DbPathHelper.MariaDbConnectionString ?? "";
        info.Host     = ExtractFromConnectionString(cs, "Server");
        info.Port     = ExtractFromConnectionString(cs, "Port");
        info.Database = ExtractFromConnectionString(cs, "Database");
        info.User     = ExtractFromConnectionString(cs, "User");

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            info.IsConnected = true;

            try
            {
                using var ver = conn.CreateCommand();
                ver.CommandText = "SELECT VERSION()";
                info.Version = ver.ExecuteScalar()?.ToString() ?? "";
            }
            catch { }

            foreach (var (table, dateCol) in _targets)
            {
                long total = CountAll(conn, table);
                long today = string.IsNullOrEmpty(dateCol) ? 0 : CountToday(conn, table, dateCol);
                info.Tables.Add(new TableStat(table, total, today, dateCol));
            }
        }
        catch { }

        return info;
    }

    private static long CountAll(DbConnection conn, string table)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT COUNT(*) FROM `{table}`";
            var v = cmd.ExecuteScalar();
            return v is null or System.DBNull ? 0 : System.Convert.ToInt64(v);
        }
        catch { return 0; }
    }

    private static long CountToday(DbConnection conn, string table, string dateCol)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT COUNT(*) FROM `{table}` WHERE DATE(`{dateCol}`) = CURDATE()";
            var v = cmd.ExecuteScalar();
            return v is null or System.DBNull ? 0 : System.Convert.ToInt64(v);
        }
        catch { return 0; }
    }

    private static string ExtractFromConnectionString(string cs, string key)
    {
        if (string.IsNullOrEmpty(cs)) return "";
        foreach (var part in cs.Split(';'))
        {
            var eq = part.IndexOf('=');
            if (eq <= 0) continue;
            var k = part.Substring(0, eq).Trim();
            if (string.Equals(k, key, System.StringComparison.OrdinalIgnoreCase))
                return part.Substring(eq + 1).Trim();
        }
        return "";
    }
}
