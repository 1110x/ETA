using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.Common;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class ContractService
{
    // ── DB 경로 (AgentService 와 동일 방식) ──────────────────────────────────

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<Contract> GetAllContracts()
    {
        var list   = new List<Contract>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Debug.WriteLine("❌ DB 없음"); return list; }

        using var conn = DbConnectionFactory.CreateConnection();
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

        using var conn = DbConnectionFactory.CreateConnection();
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
        using var conn = DbConnectionFactory.CreateConnection();
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

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM ""계약 DB"" WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", contract.C_CompanyName);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Contract DELETE] {rows}행 → {contract.C_CompanyName}");
        return rows > 0;
    }

    // ── 분석단가 테이블 컬럼 목록 조회 (FS100 이후 컬럼만) ───────────────────
    /// <summary>
    /// 분석단가 테이블의 PRAGMA table_info 로 전체 컬럼을 읽은 뒤,
    /// 이름이 "FS" 로 시작하고 번호가 100 이상인 컬럼만 순서대로 반환합니다.
    /// </summary>
public static List<string> GetUnitPriceColumns()
    {
        var result = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) 
        { 
            Debug.WriteLine("❌ DB 없음"); 
            return result; 
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"PRAGMA table_info(""분석단가"")";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            int cid = r.GetInt32(0);                    // 컬럼 순번 (0=ES, 1=Category, 2=Analyte)
            string colName = r.IsDBNull(1) ? "" : r.GetString(1);

            if (cid >= 3 && !string.IsNullOrWhiteSpace(colName))
            {
                result.Add(colName);
            }
        }

        Debug.WriteLine($"[Contract] 분석단가 컬럼 로드 완료 → {result.Count}개 (FS100~FS25)");
        return result;
    }

    // ── 분석단가 특정 컬럼 데이터 조회 ──────────────────────────────────────
    /// <summary>
    /// 분석단가 테이블에서 행 키(첫 번째 컬럼)와 지정된 컬럼 값을 함께 반환합니다.
    /// Key   = 첫 번째 컬럼 값 (항목명 등)
    /// Value = columnName 컬럼 값
    /// </summary>
    public static List<KeyValuePair<string, string>> GetUnitPriceColumnData(string columnName)
    {
        var result = new List<KeyValuePair<string, string>>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Debug.WriteLine("❌ DB 없음"); return result; }

        // 컬럼명 화이트리스트 검증 (SQL 인젝션 방지 — PRAGMA로 실제 컬럼 목록과 대조)
        var validColumns = GetAllUnitPriceColumnNames();
        if (!validColumns.Contains(columnName))
        {
            Debug.WriteLine($"❌ 유효하지 않은 컬럼명: {columnName}");
            return result;
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 첫 번째 컬럼(항목 키)을 동적으로 파악
        string keyColumn;
        using (var pragmaCmd = conn.CreateCommand())
        {
            pragmaCmd.CommandText = @"PRAGMA table_info(""분석단가"")";
            using var pr = pragmaCmd.ExecuteReader();
            keyColumn = pr.Read() ? (pr.IsDBNull(1) ? "" : pr.GetString(1)) : "";
        }

        if (string.IsNullOrEmpty(keyColumn))
        {
            Debug.WriteLine("❌ 분석단가 첫 번째 컬럼을 확인할 수 없음");
            return result;
        }

        using var cmd = conn.CreateCommand();
        // 컬럼명은 검증된 것만 사용하므로 직접 삽입 (파라미터 바인딩 불가)
        cmd.CommandText = $@"SELECT ""{keyColumn}"", ""{columnName}"" FROM ""분석단가""";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var key = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
            var val = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString() ?? "";
            result.Add(new KeyValuePair<string, string>(key, val));
        }

        Debug.WriteLine($"[Contract] 분석단가[{columnName}] {result.Count}행 로드");
        return result;
    }

    // ── 분석단가 전체 컬럼명 조회 (내부 화이트리스트용) ─────────────────────
    private static List<string> GetAllUnitPriceColumnNames()
    {
        var result = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"PRAGMA table_info(""분석단가"")";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            if (!r.IsDBNull(1)) result.Add(r.GetString(1));
        }
        return result;
    }

    // ── 공통 파라미터 ─────────────────────────────────────────────────────────
    private static void SetParams(DbCommand cmd, Contract c)
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
    private static string S(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) ?? "" : ""; }
        catch { return ""; }
    }
    private static int? NullInt(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? null : r.GetInt32(i); }
        catch { return null; }
    }
    private static decimal? NullDecimal(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? null : r.GetDecimal(i); }
        catch { return null; }
    }
}