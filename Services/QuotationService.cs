using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;
using System.Linq;

namespace ETA.Services;

public static class QuotationService
{
    private static string GetDatabasePath()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dir  = Path.Combine(root, "Data");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "eta.db");
    }

    // ── 계약업체 조회 ─────────────────────────────────────────────────────
    public static List<Contract> GetContractCompanies(bool activeOnly)
    {
        var list   = new List<Contract>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return list;

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        if (activeOnly)
        {
            var today = DateTime.Today.ToString("yyyy-MM-dd");
            cmd.CommandText = @"
                SELECT C_CompanyName, C_Abbreviation, C_ContractType,
                       C_ContractStart, C_ContractEnd, C_FacilityType, C_CategoryType
                FROM ""계약 DB""
                WHERE date(C_ContractStart) <= @today
                  AND date(C_ContractEnd)   >= @today
                ORDER BY C_CompanyName ASC";
            cmd.Parameters.AddWithValue("@today", today);
        }
        else
        {
            cmd.CommandText = @"
                SELECT C_CompanyName, C_Abbreviation, C_ContractType,
                       C_ContractStart, C_ContractEnd, C_FacilityType, C_CategoryType
                FROM ""계약 DB""
                ORDER BY C_CompanyName ASC";
        }

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var c = new Contract
            {
                C_CompanyName  = S(r, 0),
                C_Abbreviation = S(r, 1),
                C_ContractType = S(r, 2),
                C_FacilityType = S(r, 5),
                C_CategoryType = S(r, 6),
            };
            if (DateTime.TryParse(S(r, 3), out var s)) c.C_ContractStart = s;
            if (DateTime.TryParse(S(r, 4), out var e)) c.C_ContractEnd   = e;
            list.Add(c);
        }
        return list;
    }

    // ── 견적발행내역 전체 조회 ────────────────────────────────────────────
    public static List<QuotationIssue> GetAllIssues()
    {
        var list   = new List<QuotationIssue>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return list;

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return list;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT rowid,
                   ""견적발행일자"",
                   ""업체명"",
                   ""약칭"",
                   ""시료명"",
                   ""견적번호"",
                   ""적용구분"",
                   ""합계 금액""
            FROM ""견적발행내역""
            ORDER BY ""견적발행일자"" DESC, rowid DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new QuotationIssue
            {
                Id       = r.IsDBNull(0) ? 0 : Convert.ToInt32(r.GetValue(0)),
                발행일   = S(r, 1),
                업체명   = S(r, 2),
                약칭     = S(r, 3),
                시료명   = S(r, 4),
                견적번호 = S(r, 5),
                견적구분 = S(r, 6),
                총금액   = Dec(r, 7),
            });
        }
        return list;
    }

    // ── rowid 기준 단일 행 전체 컬럼 조회 ────────────────────────────────
    /// <summary>
    /// 컬럼명을 Trim() 해서 Dictionary 에 저장.
    /// 중복 방지를 위해 이미 존재하는 키는 덮어쓰지 않음.
    /// </summary>
    public static Dictionary<string, string> GetIssueRow(int rowid)
    {
        // OrdinalIgnoreCase + Trim 기반으로 조회되도록
        var dict   = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return dict;

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return dict;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT * FROM ""견적발행내역"" WHERE rowid = @id LIMIT 1";
        cmd.Parameters.AddWithValue("@id", rowid);

        using var r = cmd.ExecuteReader();
        if (r.Read())
        {
            for (int i = 0; i < r.FieldCount; i++)
            {
                // ★ 컬럼명 Trim — 앞뒤 공백 제거
                var colName = r.GetName(i).Trim();
                var val     = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString()?.Trim() ?? "";
                // 중복 키는 첫 번째 값 유지
                if (!dict.ContainsKey(colName))
                    dict[colName] = val;
            }
        }
        Debug.WriteLine($"[Quotation] GetIssueRow rowid={rowid} → {dict.Count}컬럼");
        return dict;
    }

    // ── INSERT ────────────────────────────────────────────────────────────
    public static bool Insert(QuotationIssue issue)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return false;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO ""견적발행내역""
                (""견적발행일자"", ""업체명"", ""약칭"", ""시료명"",
                 ""견적번호"",   ""적용구분"", ""합계 금액"")
            VALUES (@date,@company,@abbr,@sample,@no,@type,@amount)";
        cmd.Parameters.AddWithValue("@date",    issue.발행일   ?? DateTime.Today.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@company", issue.업체명   ?? "");
        cmd.Parameters.AddWithValue("@abbr",    issue.약칭     ?? "");
        cmd.Parameters.AddWithValue("@sample",  issue.시료명   ?? "");
        cmd.Parameters.AddWithValue("@no",      issue.견적번호 ?? "");
        cmd.Parameters.AddWithValue("@type",    issue.견적구분 ?? "");
        cmd.Parameters.AddWithValue("@amount",  issue.총금액);

        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = "SELECT last_insert_rowid()";
            issue.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        return rows > 0;
    }

    // ── DELETE ────────────────────────────────────────────────────────────
    public static bool Delete(int rowid)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM ""견적발행내역"" WHERE rowid = @id";
        cmd.Parameters.AddWithValue("@id", rowid);
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static bool TableExists(SqliteConnection conn, string name)
    {
        using var c = conn.CreateCommand();
        c.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@n";
        c.Parameters.AddWithValue("@n", name);
        return Convert.ToInt32(c.ExecuteScalar()) > 0;
    }

    private static string S(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";

    private static decimal Dec(SqliteDataReader r, int i)
    {
        if (r.IsDBNull(i)) return 0;
        // 통화기호(₩, ,) 제거 후 파싱
        var raw = r.GetValue(i)?.ToString() ?? "";
        raw = raw.Replace("₩","").Replace(",","").Trim();
        return decimal.TryParse(raw,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : 0;
    }
}
