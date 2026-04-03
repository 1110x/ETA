using System;
using System.Collections.Generic;
using System.Data.Common;
using ETA.Models;

namespace ETA.Services.Common;

public static class ReagentHistoryService
{
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `시약이력` (
                Id      INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                시약Id  INTEGER NOT NULL DEFAULT 0,
                일자    TEXT NOT NULL DEFAULT '',
                입고    INTEGER DEFAULT 0,
                출고    INTEGER DEFAULT 0,
                재고    INTEGER DEFAULT 0,
                사용중  INTEGER DEFAULT 0
            )";
        cmd.ExecuteNonQuery();
    }

    public static List<ReagentHistory> GetByReagentId(int reagentId)
    {
        var list = new List<ReagentHistory>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,시약Id,일자,입고,출고,재고,사용중 FROM `시약이력` WHERE 시약Id=@id ORDER BY 일자 ASC";
        cmd.Parameters.AddWithValue("@id", reagentId);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static List<ReagentHistory> GetByDate(string date)
    {
        var list = new List<ReagentHistory>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,시약Id,일자,입고,출고,재고,사용중 FROM `시약이력` WHERE 일자=@d ORDER BY 시약Id ASC";
        cmd.Parameters.AddWithValue("@d", date);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static bool Exists(int reagentId, string date)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM `시약이력` WHERE 시약Id=@id AND 일자=@d";
        cmd.Parameters.AddWithValue("@id", reagentId);
        cmd.Parameters.AddWithValue("@d", date);
        return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
    }

    public static bool Insert(ReagentHistory item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `시약이력` (시약Id,일자,입고,출고,재고,사용중)
            VALUES (@시약Id,@일자,@입고,@출고,@재고,@사용중)";
        cmd.Parameters.AddWithValue("@시약Id", item.시약Id);
        cmd.Parameters.AddWithValue("@일자",  item.일자);
        cmd.Parameters.AddWithValue("@입고",  item.입고);
        cmd.Parameters.AddWithValue("@출고",  item.출고);
        cmd.Parameters.AddWithValue("@재고",  item.재고);
        cmd.Parameters.AddWithValue("@사용중", item.사용중);
        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        return rows > 0;
    }

    public static int DeleteByReagentId(int reagentId)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `시약이력` WHERE 시약Id=@id";
        cmd.Parameters.AddWithValue("@id", reagentId);
        return cmd.ExecuteNonQuery();
    }

    private static ReagentHistory Map(DbDataReader r) => new()
    {
        Id     = r.IsDBNull(0) ? 0 : r.GetInt32(0),
        시약Id = r.IsDBNull(1) ? 0 : r.GetInt32(1),
        일자   = r.IsDBNull(2) ? "" : r.GetValue(2)?.ToString() ?? "",
        입고   = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3)),
        출고   = r.IsDBNull(4) ? 0 : Convert.ToInt32(r.GetValue(4)),
        재고   = r.IsDBNull(5) ? 0 : Convert.ToInt32(r.GetValue(5)),
        사용중 = r.IsDBNull(6) ? 0 : Convert.ToInt32(r.GetValue(6)),
    };
}
