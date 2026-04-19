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
                시료당소요량  DOUBLE DEFAULT 0,
                일일소요량    DOUBLE DEFAULT 0
            )";
        cmd.ExecuteNonQuery();

        // 기존 테이블 마이그레이션
        try
        {
            using var alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE `시약_분석항목` ADD COLUMN `일일소요량` DOUBLE DEFAULT 0";
            alter.ExecuteNonQuery();
        }
        catch { /* 이미 존재 */ }
    }

    public static List<ReagentAnalyte> GetByReagentId(int reagentId)
    {
        var list = new List<ReagentAnalyte>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,시약Id,분석항목,시료당소요량,일일소요량 FROM `시약_분석항목` WHERE 시약Id=@id ORDER BY 분석항목 ASC";
        cmd.Parameters.AddWithValue("@id", reagentId);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    /// <summary>특정 분석항목에 연결된 시약 Id 집합 반환.</summary>
    public static HashSet<int> GetReagentIdsByAnalyte(string analyte)
    {
        var set = new HashSet<int>();
        if (string.IsNullOrWhiteSpace(analyte)) return set;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 시약Id FROM `시약_분석항목` WHERE 분석항목=@a";
        cmd.Parameters.AddWithValue("@a", analyte);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            if (!r.IsDBNull(0)) set.Add(r.GetInt32(0));
        return set;
    }

    /// <summary>분석정보 테이블의 Analyte 목록 (Category 순). 시약/초자 필터용.</summary>
    public static List<string> GetAllAnalytes()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT DISTINCT `Analyte` FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte` <> '' ORDER BY `Category`, `{DbConnectionFactory.RowId}`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var s = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(s)) list.Add(s);
            }
        }
        catch { }
        return list;
    }

    /// <summary>분석정보의 (Analyte, Category) 쌍 목록. 콤보 뱃지 표시용.</summary>
    public static List<(string Analyte, string Category)> GetAllAnalytesWithCategory()
    {
        var list = new List<(string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `Analyte`, COALESCE(`Category`,'') FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte` <> '' ORDER BY `Category`, `{DbConnectionFactory.RowId}`";
            using var r = cmd.ExecuteReader();
            var seen = new HashSet<string>();
            while (r.Read())
            {
                var ana = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                var cat = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(ana)) continue;
                if (seen.Add(ana)) list.Add((ana, cat));
            }
        }
        catch { }
        return list;
    }

    public static bool Insert(ReagentAnalyte item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `시약_분석항목` (시약Id,분석항목,시료당소요량,일일소요량)
            VALUES (@시약Id,@분석항목,@시료당소요량,@일일소요량)";
        cmd.Parameters.AddWithValue("@시약Id", item.시약Id);
        cmd.Parameters.AddWithValue("@분석항목", item.분석항목);
        cmd.Parameters.AddWithValue("@시료당소요량", item.시료당소요량);
        cmd.Parameters.AddWithValue("@일일소요량", item.일일소요량);
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
        cmd.CommandText = @"UPDATE `시약_분석항목` SET 분석항목=@분석항목, 시료당소요량=@시료당소요량, 일일소요량=@일일소요량 WHERE Id=@id";
        cmd.Parameters.AddWithValue("@분석항목", item.분석항목);
        cmd.Parameters.AddWithValue("@시료당소요량", item.시료당소요량);
        cmd.Parameters.AddWithValue("@일일소요량", item.일일소요량);
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
        일일소요량    = r.IsDBNull(4) ? 0 : Convert.ToDouble(r.GetValue(4)),
    };
}
