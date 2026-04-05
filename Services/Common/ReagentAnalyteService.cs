using System;
using System.Collections.Generic;
using System.Data.Common;
using ETA.Models;

namespace ETA.Services.Common;

public static class ReagentAnalyteService
{
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `시약_분석항목` (
                Id           INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                시약Id       INTEGER NOT NULL DEFAULT 0,
                분석항목     TEXT NOT NULL DEFAULT '',
                시료당소요량  DOUBLE DEFAULT 0
            )";
        cmd.ExecuteNonQuery();
    }

    public static List<ReagentAnalyte> GetByReagentId(int reagentId)
    {
        var list = new List<ReagentAnalyte>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,시약Id,분석항목,시료당소요량 FROM `시약_분석항목` WHERE 시약Id=@id ORDER BY 분석항목 ASC";
        cmd.Parameters.AddWithValue("@id", reagentId);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static bool Insert(ReagentAnalyte item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `시약_분석항목` (시약Id,분석항목,시료당소요량)
            VALUES (@시약Id,@분석항목,@시료당소요량)";
        cmd.Parameters.AddWithValue("@시약Id", item.시약Id);
        cmd.Parameters.AddWithValue("@분석항목", item.분석항목);
        cmd.Parameters.AddWithValue("@시료당소요량", item.시료당소요량);
        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        return rows > 0;
    }

    public static bool Update(ReagentAnalyte item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `시약_분석항목` SET 분석항목=@분석항목, 시료당소요량=@시료당소요량 WHERE Id=@id";
        cmd.Parameters.AddWithValue("@분석항목", item.분석항목);
        cmd.Parameters.AddWithValue("@시료당소요량", item.시료당소요량);
        cmd.Parameters.AddWithValue("@id", item.Id);
        return cmd.ExecuteNonQuery() > 0;
    }

    public static bool Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `시약_분석항목` WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    private static ReagentAnalyte Map(DbDataReader r) => new()
    {
        Id           = r.IsDBNull(0) ? 0 : r.GetInt32(0),
        시약Id       = r.IsDBNull(1) ? 0 : r.GetInt32(1),
        분석항목     = r.IsDBNull(2) ? "" : r.GetValue(2)?.ToString() ?? "",
        시료당소요량  = r.IsDBNull(3) ? 0 : Convert.ToDouble(r.GetValue(3)),
    };
}
