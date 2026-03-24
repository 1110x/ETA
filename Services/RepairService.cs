using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;

namespace ETA.Services;

public static class RepairService
{
    private static string GetDatabasePath()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dir  = Path.Combine(root, "Data");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "eta.db");
    }

    // ── 테이블 자동 생성 ──────────────────────────────────────────────────
    private static void EnsureTable(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            CREATE TABLE IF NOT EXISTS ""보수요청"" (
                Id       INTEGER PRIMARY KEY AUTOINCREMENT,
                구분     TEXT    NOT NULL DEFAULT '',
                장비명   TEXT    NOT NULL DEFAULT '',
                증상     TEXT             DEFAULT '',
                위치     TEXT             DEFAULT '',
                요청자   TEXT             DEFAULT '',
                요청일   TEXT             DEFAULT '',
                완료예정일 TEXT           DEFAULT '',
                처리내용 TEXT             DEFAULT '',
                비고     TEXT             DEFAULT '',
                상태     TEXT             DEFAULT '대기'
            )";
        cmd.ExecuteNonQuery();
    }

    // ── 전체 조회 ─────────────────────────────────────────────────────────
    public static List<RepairItem> GetAll()
    {
        var list   = new List<RepairItem>();
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT Id, 구분, 장비명, 증상, 위치, 요청자, 요청일, 완료예정일, 처리내용, 비고, 상태
            FROM ""보수요청"" ORDER BY Id DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Read(r));
        return list;
    }

    // ── 연/월 요약 (트리용) ───────────────────────────────────────────────
    public static List<(int Year, int Month, int Count)> GetMonthSummary()
    {
        var result = new List<(int, int, int)>();
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT strftime('%Y', 요청일) AS yr,
                   strftime('%m', 요청일) AS mo,
                   COUNT(*)              AS cnt
            FROM ""보수요청""
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

    // ── 월별 조회 ─────────────────────────────────────────────────────────
    public static List<RepairItem> GetByMonth(int year, int month)
    {
        var list   = new List<RepairItem>();
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsureTable(conn);

        var ym = $"{year:D4}-{month:D2}";
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT Id, 구분, 장비명, 증상, 위치, 요청자, 요청일, 완료예정일, 처리내용, 비고, 상태
            FROM ""보수요청""
            WHERE strftime('%Y-%m', 요청일) = @ym
            ORDER BY Id DESC";
        cmd.Parameters.AddWithValue("@ym", ym);

        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Read(r));
        return list;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────
    public static bool Insert(RepairItem item)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO ""보수요청""
                (구분, 장비명, 증상, 위치, 요청자, 요청일, 완료예정일, 처리내용, 비고, 상태)
            VALUES
                (@구분, @장비명, @증상, @위치, @요청자, @요청일, @완료예정일, @처리내용, @비고, @상태)";

        SetParams(cmd, item);
        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = "SELECT last_insert_rowid()";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        Debug.WriteLine($"[RepairService] INSERT → {item.장비명}");
        return rows > 0;
    }

    // ── 수정 ─────────────────────────────────────────────────────────────
    public static bool Update(RepairItem item)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE ""보수요청"" SET
                구분=@구분, 장비명=@장비명, 증상=@증상, 위치=@위치,
                요청자=@요청자, 요청일=@요청일, 완료예정일=@완료예정일,
                처리내용=@처리내용, 비고=@비고, 상태=@상태
            WHERE Id=@id";

        SetParams(cmd, item);
        cmd.Parameters.AddWithValue("@id", item.Id);
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 상태 변경 ─────────────────────────────────────────────────────────
    public static bool UpdateStatus(int id, string status)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE ""보수요청"" SET 상태=@상태 WHERE Id=@id";
        cmd.Parameters.AddWithValue("@상태", status);
        cmd.Parameters.AddWithValue("@id",   id);
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────
    public static bool Delete(int id)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM ""보수요청"" WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[RepairService] DELETE id={id} → {rows}행");
        return rows > 0;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static RepairItem Read(SqliteDataReader r) => new()
    {
        Id       = r.IsDBNull(0) ? 0 : r.GetInt32(0),
        구분     = S(r, 1),
        장비명   = S(r, 2),
        증상     = S(r, 3),
        위치     = S(r, 4),
        요청자   = S(r, 5),
        요청일   = DateTime.TryParse(S(r, 6), out var d)  ? d  : DateTime.Today,
        완료예정일 = DateTime.TryParse(S(r, 7), out var d2) ? d2 : null,
        처리내용 = S(r, 8),
        비고     = S(r, 9),
        상태     = S(r, 10),
    };

    private static void SetParams(SqliteCommand cmd, RepairItem item)
    {
        cmd.Parameters.AddWithValue("@구분",     item.구분     ?? "");
        cmd.Parameters.AddWithValue("@장비명",   item.장비명   ?? "");
        cmd.Parameters.AddWithValue("@증상",     item.증상     ?? "");
        cmd.Parameters.AddWithValue("@위치",     item.위치     ?? "");
        cmd.Parameters.AddWithValue("@요청자",   item.요청자   ?? "");
        cmd.Parameters.AddWithValue("@요청일",   item.요청일.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@완료예정일", item.완료예정일.HasValue
            ? item.완료예정일.Value.ToString("yyyy-MM-dd")
            : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@처리내용", item.처리내용 ?? "");
        cmd.Parameters.AddWithValue("@비고",     item.비고     ?? "");
        cmd.Parameters.AddWithValue("@상태",     item.상태     ?? "대기");
    }

    private static string S(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
}
