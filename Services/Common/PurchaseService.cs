using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.Common;

public static class PurchaseService
{
    private static readonly string LogFile = System.IO.Path.Combine("Logs", "Purchase.log");
    private static void Log(string msg)
    {
        try
        {
            System.IO.Directory.CreateDirectory("Logs");
            System.IO.File.AppendAllText(LogFile, $"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            System.Diagnostics.Debug.WriteLine($"[Purchase] {msg}");
        }
        catch { }
    }

    // ── 테이블 자동 생성 ──────────────────────────────────────────────────────
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `물품구매` (
                Id      INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                구분    TEXT    NOT NULL DEFAULT '',
                품목    TEXT    NOT NULL DEFAULT '',
                수량    INTEGER NOT NULL DEFAULT 1,
                비고    TEXT             DEFAULT '',
                요청자  TEXT             DEFAULT '',
                요청일  TEXT             DEFAULT '',
                상태    TEXT             DEFAULT '대기'
            )";
        cmd.ExecuteNonQuery();
    }

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<PurchaseItem> GetAll() => GetAllInternal(deletedOnly: false);
    public static List<PurchaseItem> GetAllDeleted() => GetAllInternal(deletedOnly: true);

    private static List<PurchaseItem> GetAllInternal(bool deletedOnly)
    {
        var list   = new List<PurchaseItem>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        var where = deletedOnly ? "WHERE `삭제여부`=1"
                                : "WHERE (`삭제여부` IS NULL OR `삭제여부`=0)";
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT _id, 구분, 품목, 수량, 비고, 요청자, 요청일, 상태
            FROM `물품구매`
            {where}
            ORDER BY _id DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var item = new PurchaseItem
            {
                Id    = r.IsDBNull(0) ? 0    : r.GetInt32(0),
                구분  = S(r, 1),
                품목  = S(r, 2),
                수량  = r.IsDBNull(3) ? 1    : r.GetInt32(3),
                비고  = S(r, 4),
                요청자 = S(r, 5),
                상태  = S(r, 7),
            };
            if (DateTime.TryParse(S(r, 6), out var dt)) item.요청일 = dt;
            list.Add(item);
        }

        return list;
    }

    // ── 월별 조회 ────────────────────────────────────────────────────────────
    public static List<PurchaseItem> GetByMonth(int year, int month, bool deletedOnly = false)
    {
        var list   = new List<PurchaseItem>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        var ym = $"{year:D4}-{month:D2}";
        var dateFmt = DbConnectionFactory.DateFmt("요청일", "%Y-%m");
        var trashWhere = deletedOnly
            ? "`삭제여부`=1"
            : "(`삭제여부` IS NULL OR `삭제여부`=0)";

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT _id, 구분, 품목, 수량, 비고, 요청자, 요청일, 상태
            FROM `물품구매`
            WHERE {dateFmt} = @ym AND {trashWhere}
            ORDER BY _id DESC";
        cmd.Parameters.AddWithValue("@ym", ym);

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var item = new PurchaseItem
            {
                Id    = r.IsDBNull(0) ? 0 : r.GetInt32(0),
                구분  = S(r, 1), 품목 = S(r, 2),
                수량  = r.IsDBNull(3) ? 1 : r.GetInt32(3),
                비고  = S(r, 4), 요청자 = S(r, 5), 상태 = S(r, 7),
            };
            if (DateTime.TryParse(S(r, 6), out var dt)) item.요청일 = dt;
            list.Add(item);
        }
        return list;
    }

    // ── 데이터가 있는 연도/월 목록 조회 (트리 구성용) ─────────────────────────
    public static List<(int Year, int Month, int Count)> GetMonthSummary()
    {
        var result = new List<(int, int, int)>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        var dateFmtY = DbConnectionFactory.DateFmt("요청일", "%Y");
        var dateFmtM = DbConnectionFactory.DateFmt("요청일", "%m");
        cmd.CommandText = $@"
            SELECT {dateFmtY} AS yr,
                   {dateFmtM} AS mo,
                   COUNT(*)              AS cnt
            FROM `물품구매`
            GROUP BY yr, mo
            ORDER BY yr DESC, mo DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            if (int.TryParse(r.GetString(0), out int y) &&
                int.TryParse(r.GetString(1), out int m))
                result.Add((y, m, r.GetInt32(2)));
        }
        return result;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────────────
    public static bool Insert(PurchaseItem item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `물품구매` (구분, 품목, 수량, 비고, 요청자, 요청일, 상태)
            VALUES (@구분, @품목, @수량, @비고, @요청자, @요청일, @상태)";

        cmd.Parameters.AddWithValue("@구분",  item.구분  ?? "");
        cmd.Parameters.AddWithValue("@품목",  item.품목  ?? "");
        cmd.Parameters.AddWithValue("@수량",  item.수량);
        cmd.Parameters.AddWithValue("@비고",  item.비고  ?? "");
        cmd.Parameters.AddWithValue("@요청자", item.요청자 ?? "");
        cmd.Parameters.AddWithValue("@요청일", item.요청일.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@상태",  item.상태  ?? "대기");

        int rows = cmd.ExecuteNonQuery();

        // 방금 삽입된 Id 가져오기
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }

        return rows > 0;
    }

    // ── 상태 변경 ─────────────────────────────────────────────────────────────
    // ── 수정 ─────────────────────────────────────────────────────────────────
    public static bool Update(int id, string 구분, string 품목, int 수량,
                              string 요청자, string 비고)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `물품구매` SET
                구분=@구분, 품목=@품목, 수량=@수량,
                요청자=@요청자, 비고=@비고
            WHERE _id=@id";

        cmd.Parameters.AddWithValue("@구분",   구분   ?? "");
        cmd.Parameters.AddWithValue("@품목",   품목   ?? "");
        cmd.Parameters.AddWithValue("@수량",   수량);
        cmd.Parameters.AddWithValue("@요청자", 요청자 ?? "");
        cmd.Parameters.AddWithValue("@비고",   비고   ?? "");
        cmd.Parameters.AddWithValue("@id",     id);

        int rows = cmd.ExecuteNonQuery();
        Log($"Update(id={id}, {구분}/{품목} x{수량}) → rows={rows}");
        return rows > 0;
    }

    public static bool UpdateStatus(int id, string status)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `물품구매` SET 상태=@상태 WHERE _id=@id";
        cmd.Parameters.AddWithValue("@상태", status);
        cmd.Parameters.AddWithValue("@id",   id);

        int rows = cmd.ExecuteNonQuery();
        Log($"UpdateStatus(id={id}, status={status}) → rows={rows}");
        return rows > 0;
    }

    // ── 소프트 삭제 ──────────────────────────────────────────────────────────
    public static bool Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `물품구매` SET `삭제여부`=1 WHERE _id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 복원 ─────────────────────────────────────────────────────────────────
    public static bool Restore(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `물품구매` SET `삭제여부`=0 WHERE _id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    private static string S(DbDataReader r, int i)
        => r.IsDBNull(i) ? "" : r.GetString(i) ?? "";
}
