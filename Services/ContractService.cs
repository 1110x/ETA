using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class ContractService
{
    // ── DB 경로 (AgentService 와 동일 방식) ──────────────────────────────────
    public static string GetDatabasePath()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dir  = Path.Combine(root, "Data");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "eta.db");
    }

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<Contract> GetAllContracts()
    {
        var list   = new List<Contract>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) { Debug.WriteLine("❌ DB 없음"); return list; }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                   C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                   C_Address, C_Representative, C_FacilityType, C_CategoryType,
                   C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email
            FROM ""계약 DB""
            ORDER BY C_CompanyName ASC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var c = new Contract
            {
                C_CompanyName    = S(r, "C_CompanyName"),
                OriginalCompanyName = S(r, "C_CompanyName"),
                C_Abbreviation   = S(r, "C_Abbreviation"),
                C_ContractType   = S(r, "C_ContractType"),
                C_Address        = S(r, "C_Address"),
                C_Representative = S(r, "C_Representative"),
                C_FacilityType   = S(r, "C_FacilityType"),
                C_CategoryType   = S(r, "C_CategoryType"),
                C_MainProduct    = S(r, "C_MainProduct"),
                C_ContactPerson  = S(r, "C_ContactPerson"),
                C_PhoneNumber    = S(r, "C_PhoneNumber"),
                C_Email          = S(r, "C_Email"),
                C_ContractDays   = NullInt(r, "C_ContractDays"),
                C_ContractAmountVATExcluded = NullDecimal(r, "C_ContractAmountVATExcluded"),
            };

            var startStr = S(r, "C_ContractStart");
            if (DateTime.TryParse(startStr, out var start)) c.C_ContractStart = start;

            var endStr = S(r, "C_ContractEnd");
            if (DateTime.TryParse(endStr, out var end)) c.C_ContractEnd = end;

            list.Add(c);
        }

        Debug.WriteLine($"[Contract] 로드 {list.Count}건");
        return list;
    }

    // ── 수정 ─────────────────────────────────────────────────────────────────
    public static bool Update(Contract contract)
    {
        if (string.IsNullOrEmpty(contract.OriginalCompanyName))
        {
            Debug.WriteLine("❌ OriginalCompanyName 없음"); return false;
        }

        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE ""계약 DB"" SET
                C_CompanyName=@name, C_ContractStart=@start, C_ContractEnd=@end,
                C_ContractDays=@days, C_ContractAmountVATExcluded=@amount,
                C_Abbreviation=@abbr, C_ContractType=@type, C_Address=@addr,
                C_Representative=@rep, C_FacilityType=@ftype, C_CategoryType=@ctype,
                C_MainProduct=@prod, C_ContactPerson=@contact,
                C_PhoneNumber=@phone, C_Email=@email
            WHERE C_CompanyName=@original";

        SetParams(cmd, contract);
        cmd.Parameters.AddWithValue("@original", contract.OriginalCompanyName);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Contract UPDATE] {rows}행 → {contract.C_CompanyName}");
        if (rows > 0) { contract.OriginalCompanyName = contract.C_CompanyName; return true; }
        return false;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────────
    public static bool Insert(Contract contract)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO ""계약 DB""
                (C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                 C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                 C_Address, C_Representative, C_FacilityType, C_CategoryType,
                 C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email)
            VALUES
                (@name, @start, @end, @days, @amount, @abbr, @type, @addr,
                 @rep, @ftype, @ctype, @prod, @contact, @phone, @email)";

        SetParams(cmd, contract);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Contract INSERT] {rows}행 → {contract.C_CompanyName}");
        return rows > 0;
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────────
    public static bool Delete(Contract contract)
    {
        if (string.IsNullOrEmpty(contract.C_CompanyName)) return false;

        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM ""계약 DB"" WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", contract.C_CompanyName);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Contract DELETE] {rows}행 → {contract.C_CompanyName}");
        return rows > 0;
    }

    // ── 공통 파라미터 ─────────────────────────────────────────────────────────
    private static void SetParams(SqliteCommand cmd, Contract c)
    {
        cmd.Parameters.AddWithValue("@name",    c.C_CompanyName    ?? "");
        cmd.Parameters.AddWithValue("@start",   c.C_ContractStart.HasValue
                                                    ? c.C_ContractStart.Value.ToString("yyyy-MM-dd")
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@end",     c.C_ContractEnd.HasValue
                                                    ? c.C_ContractEnd.Value.ToString("yyyy-MM-dd")
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@days",    c.C_ContractDays.HasValue
                                                    ? c.C_ContractDays.Value
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@amount",  c.C_ContractAmountVATExcluded.HasValue
                                                    ? c.C_ContractAmountVATExcluded.Value
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@abbr",    c.C_Abbreviation   ?? "");
        cmd.Parameters.AddWithValue("@type",    c.C_ContractType   ?? "");
        cmd.Parameters.AddWithValue("@addr",    c.C_Address        ?? "");
        cmd.Parameters.AddWithValue("@rep",     c.C_Representative ?? "");
        cmd.Parameters.AddWithValue("@ftype",   c.C_FacilityType   ?? "");
        cmd.Parameters.AddWithValue("@ctype",   c.C_CategoryType   ?? "");
        cmd.Parameters.AddWithValue("@prod",    c.C_MainProduct    ?? "");
        cmd.Parameters.AddWithValue("@contact", c.C_ContactPerson  ?? "");
        cmd.Parameters.AddWithValue("@phone",   c.C_PhoneNumber    ?? "");
        cmd.Parameters.AddWithValue("@email",   c.C_Email          ?? "");
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────
    private static string S(SqliteDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) ?? "" : ""; }
        catch { return ""; }
    }
    private static int? NullInt(SqliteDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? null : r.GetInt32(i); }
        catch { return null; }
    }
    private static decimal? NullDecimal(SqliteDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? null : r.GetDecimal(i); }
        catch { return null; }
    }
}