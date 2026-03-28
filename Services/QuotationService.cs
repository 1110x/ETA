using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.Common;
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

    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [QService] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    // ── 계약업체 조회 ─────────────────────────────────────────────────────
    public static List<Contract> GetContractCompanies(bool activeOnly)
    {
        var list   = new List<Contract>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        if (activeOnly)
        {
            var today = DateTime.Today.ToString("yyyy-MM-dd");
            cmd.CommandText = @"
                SELECT C_CompanyName, C_Abbreviation, C_ContractType,
                       C_ContractStart, C_ContractEnd, C_FacilityType, C_CategoryType
                FROM `계약 DB`
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
                FROM `계약 DB`
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return list;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT _id,
                   `견적발행일자`,
                   `업체명`,
                   `약칭`,
                   `시료명`,
                   `견적번호`,
                   `적용구분`,
                   `담당자`,
                   `합계 금액`
            FROM `견적발행내역`
            ORDER BY `견적발행일자` DESC, _id DESC";

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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return "";

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return "";

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT `담당자`
            FROM `견적발행내역`
            WHERE `업체명` = @company
              AND `담당자` IS NOT NULL
              AND TRIM(`담당자`) != ''
            ORDER BY `견적발행일자` DESC, _id DESC
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return dict;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return dict;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT * FROM `견적발행내역` WHERE _id = @id LIMIT 1";
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        // 분석단가 테이블에서 제외할 고정 컬럼
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Analyte", "Category", "ES", "unit", "Unit", "Method", "방법",
            "비고", "Note", "단위", "분류",
        };

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            if (!TableExists(conn, "분석단가"))
            {
                Debug.WriteLine("[GetContractTypes] 분석단가 테이블 없음");
                return list;
            }

            foreach (var col in DbConnectionFactory.GetColumnNames(conn, "분석단가"))
            {
                if (!fixedCols.Contains(col.Trim()))
                    list.Add(col.Trim());
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return dict;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var cols = DbConnectionFactory.GetColumnNames(conn, "분석단가");

            var match = cols.FirstOrDefault(c =>
                string.Equals(c.Trim(), columnName.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null) { Debug.WriteLine($"[Prices] '{columnName}' 컬럼 없음"); return dict; }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT `Analyte`, `{match}` FROM `분석단가`";
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
        Log($"Insert 시작: DbPathHelper.DbPath={DbPathHelper.DbPath}  존재={File.Exists(DbPathHelper.DbPath)}");
        Log($"  issue: 발행일={issue.발행일} 업체명={issue.업체명} 번호={issue.견적번호} 구분={issue.견적구분}");
        Log($"  itemData 입력: {itemData?.Count ?? 0}개");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        bool tableExists = TableExists(conn, "견적발행내역");
        Log($"  TableExists(견적발행내역)={tableExists}");
        if (!tableExists) return false;

        // ── 누락 컬럼 마이그레이션 ──────────────────────────────────────
        MigrateIssueColumns(conn);

        // ── 테이블 실제 컬럼 목록 조회 (항목 컬럼 검증용) ──────────────
        var tableCols = new HashSet<string>(DbConnectionFactory.GetColumnNames(conn, "견적발행내역"), StringComparer.OrdinalIgnoreCase);
        Log($"  마이그레이션 후 테이블 컬럼 수={tableCols.Count}");

        // ── 저장할 항목 필터링 (분석항목 컬럼만: 본체·단가·소계 모두 존재해야 함) ──
        var validItems = itemData?
            .Where(kv => tableCols.Contains(kv.Key)
                      && tableCols.Contains(kv.Key + "단가")
                      && tableCols.Contains(kv.Key + "소계")
                      && kv.Value.Qty > 0)
            .ToList() ?? new();
        Log($"  validItems(Qty>0, 컬럼존재)={validItems.Count}개");

        // ── 고정 컬럼도 존재 여부 확인 후 동적 추가 ─────────────────────
        var fixedCols = new List<(string Col, string Param, object? Value)>
        {
            ("견적발행일자", "@date",    issue.발행일   ?? DateTime.Today.ToString("yyyy-MM-dd")),
            ("업체명",       "@company", issue.업체명   ?? ""),
            ("약칭",         "@abbr",    issue.약칭     ?? ""),
            ("시료명",       "@sample",  issue.시료명   ?? ""),
            ("견적번호",     "@no",      issue.견적번호 ?? ""),
            ("적용구분",     "@type",    issue.견적구분 ?? ""),
            ("담당자",       "@manager", issue.담당자   ?? ""),
            ("합계 금액",    "@amount",  issue.총금액),
        };
        var activeFix = fixedCols.Where(f => tableCols.Contains(f.Col)).ToList();
        Log($"  고정컬럼 삽입 대상: {string.Join(", ", activeFix.Select(f => f.Col))}");

        // ── INSERT 쿼리 동적 생성 ────────────────────────────────────────
        var colList   = activeFix.Select(f => $"\"{f.Col}\"").ToList();
        var paramList = activeFix.Select(f => f.Param).ToList();

        foreach (var kv in validItems)
        {
            var name = kv.Key;
            colList.Add($"`{name}`");
            colList.Add($"`{name}단가`");
            colList.Add($"`{name}소계`");
            paramList.Add($"@qty_{ToParamName(name)}");
            paramList.Add($"@price_{ToParamName(name)}");
            paramList.Add($"@sub_{ToParamName(name)}");
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO `견적발행내역`
                ({string.Join(", ", colList)})
            VALUES
                ({string.Join(", ", paramList)})";

        // 고정 컬럼 파라미터 (존재하는 것만)
        foreach (var (_, param, value) in activeFix)
            cmd.Parameters.AddWithValue(param, value ?? "");

        foreach (var kv in validItems)
        {
            var p = ToParamName(kv.Key);
            var sub = kv.Value.Qty * kv.Value.Price;
            cmd.Parameters.AddWithValue($"@qty_{p}",   kv.Value.Qty);
            cmd.Parameters.AddWithValue($"@price_{p}", kv.Value.Price);
            cmd.Parameters.AddWithValue($"@sub_{p}",   sub);
        }

        int rows;
        try
        {
            rows = cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            Log($"  ExecuteNonQuery 예외: {ex.GetType().Name}: {ex.Message}");
            Log($"  SQL: {cmd.CommandText[..Math.Min(300, cmd.CommandText.Length)]}");
            throw;
        }

        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            issue.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        Log($"  ExecuteNonQuery rows={rows}  new_id={issue.Id}");
        return rows > 0;
    }

    // 파라미터명용 안전 문자열 변환 (특수문자 제거)
    private static string ToParamName(string name)
        => System.Text.RegularExpressions.Regex.Replace(name, @"[^a-zA-Z0-9가-힣]", "_");

    // ── DELETE ────────────────────────────────────────────────────────────
    public static bool Delete(int rowid)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM `견적발행내역` WHERE _id = @id";
        cmd.Parameters.AddWithValue("@id", rowid);
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    // ── 견적발행내역 누락 컬럼 자동 추가 ─────────────────────────────────────
    private static void MigrateIssueColumns(DbConnection conn)
    {
        // 추가해야 할 컬럼 목록: (컬럼명, 타입, 기본값)
        var needed = new[]
        {
            ("담당자", "TEXT", "''"),
            ("약칭",   "TEXT", "''"),
        };

        var existing = new HashSet<string>(DbConnectionFactory.GetColumnNames(conn, "견적발행내역"), StringComparer.OrdinalIgnoreCase);

        foreach (var (col, type, def) in needed)
        {
            if (existing.Contains(col)) continue;
            try
            {
                using var alter = conn.CreateCommand();
                alter.CommandText = $@"ALTER TABLE `견적발행내역` ADD COLUMN `{col}` {type} DEFAULT {def}";
                alter.ExecuteNonQuery();
                Log($"  마이그레이션: '{col}' 컬럼 추가 완료");
            }
            catch (Exception ex)
            {
                Log($"  마이그레이션 실패 '{col}': {ex.Message}");
            }
        }
    }

    private static bool TableExists(DbConnection conn, string name)
        => DbConnectionFactory.TableExists(conn, name);

    private static string S(DbDataReader r, int i)
        => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";

    private static decimal Dec(DbDataReader r, int i)
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
