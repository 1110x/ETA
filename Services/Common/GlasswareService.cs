using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using ETA.Models;

namespace ETA.Services.Common;

public static class GlasswareService
{
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `초자` (
                Id     INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                품목명 TEXT    NOT NULL DEFAULT '',
                용도   TEXT             DEFAULT '',
                규격   TEXT             DEFAULT '',
                재질   TEXT             DEFAULT '유리',
                수량   INTEGER          DEFAULT 0,
                단가   REAL             DEFAULT 0,
                비고   TEXT             DEFAULT '',
                등록일 TEXT             DEFAULT '',
                상태   TEXT             DEFAULT '정상'
            )";
        cmd.ExecuteNonQuery();
    }

    public static List<Glassware> GetAll()
    {
        var list = new List<Glassware>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,품목명,용도,규격,재질,수량,단가,비고,등록일,상태 FROM `초자` ORDER BY Id DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static List<Glassware> GetByStatus(string status)
    {
        var list = new List<Glassware>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id,품목명,용도,규격,재질,수량,단가,비고,등록일,상태 FROM `초자` WHERE 상태=@s ORDER BY Id DESC";
        cmd.Parameters.AddWithValue("@s", status);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static bool Insert(Glassware item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `초자` (품목명,용도,규격,재질,수량,단가,비고,등록일,상태)
            VALUES (@품목명,@용도,@규격,@재질,@수량,@단가,@비고,@등록일,@상태)";
        SetParams(cmd, item);
        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        Debug.WriteLine($"[GlasswareService] INSERT → {item.품목명}");
        return rows > 0;
    }

    public static bool Update(Glassware item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `초자` SET
                품목명=@품목명, 용도=@용도, 규격=@규격, 재질=@재질,
                수량=@수량, 단가=@단가, 비고=@비고, 등록일=@등록일, 상태=@상태
            WHERE Id=@id";
        SetParams(cmd, item);
        cmd.Parameters.AddWithValue("@id", item.Id);
        return cmd.ExecuteNonQuery() > 0;
    }

    public static bool Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `초자` WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[GlasswareService] DELETE id={id} → {rows}행");
        return rows > 0;
    }

    private static void SetParams(DbCommand cmd, Glassware item)
    {
        cmd.Parameters.AddWithValue("@품목명", item.품목명 ?? "");
        cmd.Parameters.AddWithValue("@용도",   item.용도   ?? "");
        cmd.Parameters.AddWithValue("@규격",   item.규격   ?? "");
        cmd.Parameters.AddWithValue("@재질",   item.재질   ?? "유리");
        cmd.Parameters.AddWithValue("@수량",   item.수량);
        cmd.Parameters.AddWithValue("@단가",   (double)item.단가);
        cmd.Parameters.AddWithValue("@비고",   item.비고   ?? "");
        cmd.Parameters.AddWithValue("@등록일", item.등록일 ?? "");
        cmd.Parameters.AddWithValue("@상태",   item.상태   ?? "정상");
    }

    private static Glassware Map(DbDataReader r) => new()
    {
        Id     = r.IsDBNull(0) ? 0  : r.GetInt32(0),
        품목명 = S(r, 1),
        용도   = S(r, 2),
        규격   = S(r, 3),
        재질   = S(r, 4),
        수량   = r.IsDBNull(5) ? 0  : Convert.ToInt32(r.GetValue(5)),
        단가   = r.IsDBNull(6) ? 0m : Convert.ToDecimal(r.GetValue(6)),
        비고   = S(r, 7),
        등록일 = S(r, 8),
        상태   = S(r, 9),
    };

    private static string S(DbDataReader r, int i) =>
        r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
}
