using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;

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
                   ""담당자"",
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
                담당자   = S(r, 7),
                총금액   = Dec(r, 8),
            });
        }
        return list;
    }

    // ── 업체명으로 최근 견적 담당자 조회 ───────────────────────────────
    public static string GetLatestManagerForCompany(string companyName)
    {
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return "";

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return "";

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT ""담당자""
            FROM ""견적발행내역""
            WHERE ""업체명"" = @company
              AND ""담당자"" IS NOT NULL
              AND TRIM(""담당자"") != ''
            ORDER BY ""견적발행일자"" DESC, rowid DESC
            LIMIT 1";
        cmd.Parameters.AddWithValue("@company", companyName);

        using var r = cmd.ExecuteReader();
        if (r.Read())
        {
            return S(r, 0);
        }
        return "";
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

    // ── 분석단가 테이블 컬럼헤더에서 적용구분 목록 조회 ─────────────────────
    /// <summary>
    /// 분석단가 테이블의 컬럼 중 Analyte, Category, ES 등 고정 컬럼을 제외한
    /// 나머지 컬럼헤더 (FS100, FS100+, FS56 ...) 를 반환
    /// </summary>
    public static List<string> GetContractTypes()
    {
        var list   = new List<string>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return list;

        // 분석단가 테이블에서 제외할 고정 컬럼
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Analyte", "Category", "ES", "unit", "Unit", "Method", "방법",
            "비고", "Note", "단위", "분류",
        };

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            if (!TableExists(conn, "분석단가"))
            {
                Debug.WriteLine("[GetContractTypes] 분석단가 테이블 없음");
                return list;
            }

            using var pragma = conn.CreateCommand();
            pragma.CommandText = @"PRAGMA table_info(""분석단가"")";
            using var r = pragma.ExecuteReader();
            while (r.Read())
            {
                var col = r.GetString(1).Trim();  // 컬럼명
                if (!fixedCols.Contains(col))
                    list.Add(col);
            }

            Debug.WriteLine($"[GetContractTypes] 분석단가 컬럼 → {list.Count}개: {string.Join(", ", list)}");
        }
        catch (Exception ex) { Debug.WriteLine($"[GetContractTypes] {ex.Message}"); }
        return list;
    }

    // ── 분석단가 테이블에서 적용구분 컬럼 단가 조회 ──────────────────────────
    public static Dictionary<string, decimal> GetPricesByColumn(string columnName)
    {
        var dict   = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return dict;
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            using var pragma = conn.CreateCommand();
            pragma.CommandText = @"PRAGMA table_info(""분석단가"")";
            var cols = new List<string>();
            using (var pr = pragma.ExecuteReader())
                while (pr.Read()) cols.Add(pr.GetString(1));

            var match = cols.FirstOrDefault(c =>
                string.Equals(c.Trim(), columnName.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null) { Debug.WriteLine($"[Prices] '{columnName}' 컬럼 없음"); return dict; }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT ""Analyte"", ""{match}"" FROM ""분석단가""";
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                var analyte = dr.IsDBNull(0) ? "" : dr.GetValue(0)?.ToString() ?? "";
                if (string.IsNullOrEmpty(analyte)) continue;
                if (!dr.IsDBNull(1) && decimal.TryParse(
                    dr.GetValue(1)?.ToString(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var p))
                    dict[analyte] = p;
            }
            Debug.WriteLine($"[Prices] {columnName} → {dict.Count}개");
        }
        catch (Exception ex) { Debug.WriteLine($"[Prices] {ex.Message}"); }
        return dict;
    }

    // ── INSERT (분석항목 포함) ────────────────────────────────────────────
    public static bool Insert(QuotationIssue issue,
        Dictionary<string, (int Qty, decimal Price)>? itemData = null)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return false;

        // ── 테이블 실제 컬럼 목록 조회 (항목 컬럼 검증용) ──────────────
        var tableCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = @"PRAGMA table_info(""견적발행내역"")";
            using var pr = pragma.ExecuteReader();
            while (pr.Read()) tableCols.Add(pr.GetString(1).Trim());
        }

        // ── 저장할 항목 필터링 (테이블에 존재하는 컬럼만) ───────────────
        var validItems = itemData?
            .Where(kv => tableCols.Contains(kv.Key) && kv.Value.Qty > 0)
            .ToList() ?? new();

        // ── INSERT 쿼리 동적 생성 ────────────────────────────────────────
        var colList = new List<string>
        {
            "\"견적발행일자\"", "\"업체명\"", "\"약칭\"", "\"시료명\"",
            "\"견적번호\"", "\"적용구분\"", "\"담당자\"", "\"합계 금액\""
        };
        var paramList = new List<string>
        {
            "@date", "@company", "@abbr", "@sample",
            "@no", "@type", "@manager", "@amount"
        };

        foreach (var kv in validItems)
        {
            var name = kv.Key;
            colList.Add($"\"{name}\"");
            colList.Add($"\"{name}단가\"");
            colList.Add($"\"{name}소계\"");
            paramList.Add($"@qty_{ToParamName(name)}");
            paramList.Add($"@price_{ToParamName(name)}");
            paramList.Add($"@sub_{ToParamName(name)}");
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO ""견적발행내역""
                ({string.Join(", ", colList)})
            VALUES
                ({string.Join(", ", paramList)})";

        cmd.Parameters.AddWithValue("@date",    issue.발행일   ?? DateTime.Today.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@company", issue.업체명   ?? "");
        cmd.Parameters.AddWithValue("@abbr",    issue.약칭     ?? "");
        cmd.Parameters.AddWithValue("@sample",  issue.시료명   ?? "");
        cmd.Parameters.AddWithValue("@no",      issue.견적번호 ?? "");
        cmd.Parameters.AddWithValue("@type",    issue.견적구분 ?? "");
        cmd.Parameters.AddWithValue("@manager", issue.담당자   ?? "");
        cmd.Parameters.AddWithValue("@amount",  issue.총금액);

        foreach (var kv in validItems)
        {
            var p = ToParamName(kv.Key);
            var sub = kv.Value.Qty * kv.Value.Price;
            cmd.Parameters.AddWithValue($"@qty_{p}",   kv.Value.Qty);
            cmd.Parameters.AddWithValue($"@price_{p}", kv.Value.Price);
            cmd.Parameters.AddWithValue($"@sub_{p}",   sub);
        }

        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = "SELECT last_insert_rowid()";
            issue.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        Debug.WriteLine($"[Insert] {rows}행 → {issue.견적번호}  항목{validItems.Count}개");
        return rows > 0;
    }

    // 파라미터명용 안전 문자열 변환 (특수문자 제거)
    private static string ToParamName(string name)
        => System.Text.RegularExpressions.Regex.Replace(name, @"[^a-zA-Z0-9가-힣]", "_");

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
