using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.Common;
using ETA.Models;
using System.Diagnostics;
using ETA.Services.Common;
using ClosedXML.Excel;

namespace ETA.Services.SERVICE1;

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
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Quotation.log"));

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

    // ── 테이블 초기화 ─────────────────────────────────────────────────────
    /// <summary>견적발행내역 테이블이 없으면 기본 컬럼으로 생성한다.</summary>
    public static void EnsureQuotationIssueTable()
    {

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"CREATE TABLE IF NOT EXISTS `견적발행내역` (
                `견적발행일자`   VARCHAR(20)  NULL,
                `업체명`         VARCHAR(200) NULL,
                `약칭`           VARCHAR(100) NULL,
                `시료명`         VARCHAR(200) NULL,
                `견적번호`       VARCHAR(100) NULL,
                `적용구분`       VARCHAR(50)  NULL,
                `담당자`         VARCHAR(100) NULL,
                `담당자연락처`   VARCHAR(100) NULL,
                `담당자 e-Mail`  VARCHAR(200) NULL,
                `합계 금액`      DECIMAL(18,2) NULL DEFAULT 0,
                `거래명세서번호` VARCHAR(50)  NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
        try { cmd.ExecuteNonQuery(); Log("EnsureQuotationIssueTable: 테이블 생성(또는 이미 존재)"); }
        catch (Exception ex) { Log($"EnsureQuotationIssueTable 오류: {ex.Message}"); }
    }

    // ── 견적발행내역 전체 조회 ────────────────────────────────────────────
    /// <summary>견적발행내역에 거래명세서번호 컬럼이 없으면 추가</summary>
    public static void EnsureTradeStatementColumn()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return;
        var cols = DbConnectionFactory.GetColumnNames(conn, "견적발행내역");
        if (cols.Contains("거래명세서번호")) return;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "ALTER TABLE `견적발행내역` ADD COLUMN `거래명세서번호` VARCHAR(50) NULL DEFAULT NULL";
        try { cmd.ExecuteNonQuery(); } catch (Exception ex) { Debug.WriteLine($"[Quotation] 컬럼추가 오류: {ex.Message}"); }
    }

    public static List<QuotationIssue> GetAllIssues()
    {
        var list   = new List<QuotationIssue>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return list;

        // 거래명세서번호 컬럼 존재 여부 확인
        var cols = DbConnectionFactory.GetColumnNames(conn, "견적발행내역");
        bool hasTrade = cols.Contains("거래명세서번호");

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT {DbConnectionFactory.RowId},
                   `견적발행일자`,
                   `업체명`,
                   `약칭`,
                   `시료명`,
                   `견적번호`,
                   `적용구분`,
                   `담당자`,
                   `합계 금액`
                   {(hasTrade ? ", `거래명세서번호`" : "")}
            FROM `견적발행내역`
            ORDER BY `견적발행일자` DESC, {DbConnectionFactory.RowId} DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new QuotationIssue
            {
                Id            = r.IsDBNull(0) ? 0 : Convert.ToInt32(r.GetValue(0)),
                발행일        = S(r, 1),
                업체명        = S(r, 2),
                약칭          = S(r, 3),
                시료명        = S(r, 4),
                견적번호      = S(r, 5),
                견적구분      = S(r, 6),
                담당자        = S(r, 7),
                총금액        = Dec(r, 8),
                거래명세서번호 = hasTrade ? S(r, 9) : "",
            });
        }
        return list;
    }

    // ── 업체명으로 최근 견적 담당자 조회 ───────────────────────────────
    public static string GetLatestManagerForCompany(string companyName)
    {

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
            ORDER BY `견적발행일자` DESC, {DbConnectionFactory.RowId} DESC
            LIMIT 1";
        cmd.Parameters.AddWithValue("@company", companyName);

        using var r = cmd.ExecuteReader();
        if (r.Read())
        {
            return S(r, 0);
        }
        return "";
    }

    // ── 업체명으로 과거 담당자 목록 중복제거 조회 ────────────────────────
    public static List<string> GetDistinctManagersForCompany(string companyName)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!TableExists(conn, "견적발행내역")) return list;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `담당자`
                FROM `견적발행내역`
                WHERE `업체명` = @company
                  AND `담당자` IS NOT NULL
                  AND TRIM(`담당자`) != ''
                ORDER BY `담당자` ASC";
            cmd.Parameters.AddWithValue("@company", companyName);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(S(r, 0));
        }
        catch { }
        return list;
    }

    // ── 업체+담당자명으로 연락처/이메일 조회 (최신 발행건 기준) ────────────
    public static (string Phone, string Email) GetManagerContactInfo(string companyName, string managerName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!TableExists(conn, "견적발행내역")) return ("", "");

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT `담당자연락처`, `담당자 e-Mail`
                FROM `견적발행내역`
                WHERE `업체명` = @company
                  AND `담당자` = @manager
                  AND (`담당자연락처` IS NOT NULL AND TRIM(`담당자연락처`) != ''
                    OR `담당자 e-Mail` IS NOT NULL AND TRIM(`담당자 e-Mail`) != '')
                ORDER BY `견적발행일자` DESC
                LIMIT 1";
            cmd.Parameters.AddWithValue("@company", companyName);
            cmd.Parameters.AddWithValue("@manager", managerName);
            using var r = cmd.ExecuteReader();
            if (r.Read()) return (S(r, 0), S(r, 1));
        }
        catch { }
        return ("", "");
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

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!TableExists(conn, "견적발행내역")) return dict;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM `견적발행내역` WHERE {DbConnectionFactory.RowId} = @id LIMIT 1";
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

    // ── 계약 DB 에서 적용구분(계약구분) 고유값 목록 조회 ──────────────────────
    /// <summary>계약 DB 의 C_ContractType 고유값. 없으면 기본값 반환.</summary>
    public static List<string> GetContractTypes()
    {
        var list = new List<string>();

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT C_ContractType FROM `계약 DB` " +
                              "WHERE C_ContractType IS NOT NULL AND C_ContractType <> '' " +
                              "ORDER BY C_ContractType ASC";
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                var v = dr.IsDBNull(0) ? "" : dr.GetString(0);
                if (!string.IsNullOrWhiteSpace(v)) list.Add(v);
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[GetContractTypes] {ex.Message}"); }

        if (list.Count == 0)
            list.AddRange(new[] { "위탁", "용역", "구매", "기타" });

        return list;
    }

    // ── 계약 DB 에서 업체별 단가 조회 ────────────────────────────────────────
    /// <summary>계약 DB 에서 특정 업체의 분석항목 단가를 반환합니다.</summary>
    public static Dictionary<string, decimal> GetPricesByCompany(string companyName)
    {
        var dict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(companyName)) return dict;

        try
        {
            foreach (var (analyte, priceStr) in ContractService.GetContractPrices(companyName))
            {
                var clean = (priceStr ?? "").Replace(",", "").Trim();
                if (decimal.TryParse(clean, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var d) && d > 0)
                    dict[analyte] = d;
            }
            Debug.WriteLine($"[Prices] {companyName} -> {dict.Count}건");
        }
        catch (Exception ex) { Debug.WriteLine($"[GetPricesByCompany] {ex.Message}"); }
        return dict;
    }
    // ── INSERT (분석항목 포함) ────────────────────────────────────────────
    public static bool Insert(QuotationIssue issue,
        Dictionary<string, (int Qty, decimal Price)>? itemData = null)
    {
        Log($"  issue: 발행일={issue.발행일} 업체명={issue.업체명} 번호={issue.견적번호} 구분={issue.견적구분}");
        Log($"  itemData 입력: {itemData?.Count ?? 0}개");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        bool tableExists = TableExists(conn, "견적발행내역");
        Log($"  TableExists(견적발행내역)={tableExists}");
        if (!tableExists)
        {
            Log("  테이블 없음 → EnsureQuotationIssueTable() 호출");
            EnsureQuotationIssueTable();
        }

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
            ("담당자연락처",  "@phone",   issue.담당자연락처 ?? ""),
            ("담당자 e-Mail", "@email",   issue.담당자이메일 ?? ""),
            ("합계 금액",    "@amount",  issue.총금액),
        };
        var activeFix = fixedCols.Where(f => tableCols.Contains(f.Col)).ToList();
        Log($"  고정컬럼 삽입 대상: {string.Join(", ", activeFix.Select(f => f.Col))}");

        // ── INSERT 쿼리 동적 생성 ────────────────────────────────────────
        var colList   = activeFix.Select(f => $"`{f.Col}`").ToList();
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

    // ── 체크박스 변경 저장 — 체크=기존값 유지(없으면 수량1), 해제=NULL ────────
    public static bool UpdateIssueAnalytes(
        int rowId,
        IEnumerable<string> allAnalytes,
        IEnumerable<string> checkedAnalytes)
    {
        var all        = allAnalytes.Select(a => a.Trim()).Where(a => a.Length > 0).ToList();
        var checkedSet = new HashSet<string>(checkedAnalytes.Select(a => a.Trim()),
                             StringComparer.OrdinalIgnoreCase);
        if (all.Count == 0) return true;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var currentRow = GetIssueRow(rowId);
            var tableCols  = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "견적발행내역"),
                StringComparer.OrdinalIgnoreCase);

            var setParts = new List<string>();
            int idx = 0;
            var pvals = new List<(string p, object? v)>();

            foreach (var a in all)
            {
                if (!tableCols.Contains(a)) continue;
                if (checkedSet.Contains(a))
                {
                    // 기존 수량 유지, 없으면 1
                    var qty = currentRow.TryGetValue(a, out var eq) && IsNonZero(eq) ? eq : "1";
                    setParts.Add($"`{a}` = @p{idx}");
                    pvals.Add(($"@p{idx}", qty)); idx++;

                    if (tableCols.Contains(a + "단가"))
                    {
                        var pr = currentRow.TryGetValue(a + "단가", out var ep) ? ep ?? "0" : "0";
                        setParts.Add($"`{a}단가` = @p{idx}");
                        pvals.Add(($"@p{idx}", pr)); idx++;
                    }
                    if (tableCols.Contains(a + "소계"))
                    {
                        var sb = currentRow.TryGetValue(a + "소계", out var es) ? es ?? "0" : "0";
                        setParts.Add($"`{a}소계` = @p{idx}");
                        pvals.Add(($"@p{idx}", sb)); idx++;
                    }
                }
                else
                {
                    setParts.Add($"`{a}` = NULL");
                    if (tableCols.Contains(a + "단가")) setParts.Add($"`{a}단가` = NULL");
                    if (tableCols.Contains(a + "소계")) setParts.Add($"`{a}소계` = NULL");
                }
            }
            if (setParts.Count == 0) return true;

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"UPDATE `견적발행내역` SET {string.Join(", ", setParts)} " +
                $"WHERE {DbConnectionFactory.RowId} = @id";
            foreach (var (p, v) in pvals)
                cmd.Parameters.AddWithValue(p, v ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@id", rowId);
            int rows = cmd.ExecuteNonQuery();
            Log($"UpdateIssueAnalytes rowId={rowId} → checked={checkedSet.Count} rows={rows}");
            return true;
        }
        catch (Exception ex) { Log($"UpdateIssueAnalytes 오류: {ex.Message}"); return false; }
    }

    private static bool IsNonZero(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        var clean = val.Replace("₩", "").Replace(",", "").Trim();
        return decimal.TryParse(clean,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var d) ? d != 0 : true;
    }

    // ── DELETE ────────────────────────────────────────────────────────────
    public static bool Delete(int rowid)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"DELETE FROM `견적발행내역` WHERE {DbConnectionFactory.RowId} = @id";
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
            ("담당자연락처",  "TEXT", "''"),
            ("담당자 e-Mail", "TEXT", "''"),
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
        var raw = r.GetValue(i)?.ToString() ?? "";
        raw = raw.Replace("₩","").Replace(",","").Trim();
        return decimal.TryParse(raw,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : 0;
    }

    // ── 거래명세서 ────────────────────────────────────────────────────────

    private const int MaxQuotationNoCols = 10;  // 견적번호1 ~ 견적번호10

    /// <summary>거래명세서발행내역 기본 테이블 생성 보장 (견적발행내역과 유사한 구조)</summary>
    public static void EnsureTradeStatementTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 테이블 기본 생성
        using (var cmd = conn.CreateCommand())
        {
            string quotNoCols = string.Concat(Enumerable.Range(1, MaxQuotationNoCols).Select(i =>
                $"\n                    `견적번호{i}`      VARCHAR(100) NULL,"));

            cmd.CommandText = $@"CREATE TABLE IF NOT EXISTS `거래명세서발행내역` (
                    `발행일`         VARCHAR(20)   NULL,
                    `업체명`         VARCHAR(200)  NULL,
                    `약칭`           VARCHAR(100)  NULL,
                    `거래명세서번호` VARCHAR(100)  NULL,{quotNoCols}
                    `공급가액`       DECIMAL(20,2) NULL DEFAULT 0,
                    `부가세`         DECIMAL(20,2) NULL DEFAULT 0,
                    `합계금액`       DECIMAL(20,2) NULL DEFAULT 0
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
            try { cmd.ExecuteNonQuery(); Log("EnsureTradeStatementTable: 생성(또는 이미 존재)"); }
            catch (Exception ex) { Log($"EnsureTradeStatementTable 오류: {ex.Message}"); }
        }

        // 2. 기존 테이블에 누락된 고정 컬럼 마이그레이션
        var existCols = new HashSet<string>(DbConnectionFactory.GetColumnNames(conn, "거래명세서발행내역"),
                            StringComparer.OrdinalIgnoreCase);
        var needed = new List<(string col, string type)>
        {
            ("약칭", "VARCHAR(100)"),
        };
        for (int i = 1; i <= MaxQuotationNoCols; i++)
            needed.Add(($"견적번호{i}", "VARCHAR(100)"));

        foreach (var (col, type) in needed)
        {
            if (existCols.Contains(col)) continue;
            using var alt = conn.CreateCommand();
            alt.CommandText = $"ALTER TABLE `거래명세서발행내역` ADD COLUMN `{col}` {type} NULL";
            try { alt.ExecuteNonQuery(); Log($"  거래명세서 마이그레이션: '{col}' 추가"); }
            catch { }
        }
    }

    /// <summary>거래명세서발행내역에 분석항목 컬럼이 없으면 동적으로 추가 (수량/단가/소계)</summary>
    private static void EnsureTradeStatementItemColumns(DbConnection conn,
        IEnumerable<string> itemNames)
    {
        var existCols = new HashSet<string>(DbConnectionFactory.GetColumnNames(conn, "거래명세서발행내역"),
                            StringComparer.OrdinalIgnoreCase);
        foreach (var name in itemNames)
        {
            foreach (var suffix in new[] { "", "단가", "소계" })
            {
                var col = name + suffix;
                if (existCols.Contains(col)) continue;
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE `거래명세서발행내역` ADD COLUMN `{col}` DECIMAL(20,2) NULL DEFAULT 0";
                try { cmd.ExecuteNonQuery(); existCols.Add(col); }
                catch { }
            }
        }
    }

    /// <summary>거래명세서발행내역에 레코드 삽입. 삽입된 rowid 반환 (-1=실패)</summary>
    /// <param name="itemData">항목명 → (수량, 단가, 소계) — null이면 항목 컬럼 미삽입</param>
    public static int InsertTradeStatement(string companyName, string abbreviation,
        string statementNo, IEnumerable<string> quotationNos, decimal supplyAmt, decimal vat,
        decimal total, Dictionary<string, (decimal qty, decimal unitPrice, decimal subtotal)>? itemData = null)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        if (!TableExists(conn, "거래명세서발행내역")) return -1;

        // 항목 컬럼 보장
        if (itemData != null && itemData.Count > 0)
            EnsureTradeStatementItemColumns(conn, itemData.Keys);

        var tableCols = new HashSet<string>(
            DbConnectionFactory.GetColumnNames(conn, "거래명세서발행내역"),
            StringComparer.OrdinalIgnoreCase);

        // 고정 컬럼 구성
        var quotNos = quotationNos.ToList();
        var colList   = new List<string>();
        var paramList = new List<string>();
        var @params   = new Dictionary<string, object?>();

        void Add(string col, string param, object? val)
        {
            if (!tableCols.Contains(col)) return;
            colList.Add($"`{col}`");
            paramList.Add(param);
            @params[param] = val;
        }

        Add("발행일",         "@dt",   DateTime.Today.ToString("yyyy-MM-dd"));
        Add("업체명",         "@co",   companyName);
        Add("약칭",           "@abbr", abbreviation);
        Add("거래명세서번호", "@no",   statementNo);
        for (int i = 0; i < MaxQuotationNoCols; i++)
            Add($"견적번호{i + 1}", $"@qno{i + 1}", i < quotNos.Count ? quotNos[i] : (object?)DBNull.Value);
        Add("공급가액", "@sup", supplyAmt);
        Add("부가세",   "@vat", vat);
        Add("합계금액", "@tot", total);

        // 항목 컬럼 구성
        if (itemData != null)
        {
            foreach (var kv in itemData)
            {
                var n = kv.Key;
                if (tableCols.Contains(n))          { colList.Add($"`{n}`");      paramList.Add($"@q_{ToParamName(n)}");   @params[$"@q_{ToParamName(n)}"]   = kv.Value.qty; }
                if (tableCols.Contains(n + "단가")) { colList.Add($"`{n}단가`"); paramList.Add($"@up_{ToParamName(n)}"); @params[$"@up_{ToParamName(n)}"] = kv.Value.unitPrice; }
                if (tableCols.Contains(n + "소계")) { colList.Add($"`{n}소계`"); paramList.Add($"@sub_{ToParamName(n)}"); @params[$"@sub_{ToParamName(n)}"] = kv.Value.subtotal; }
            }
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO `거래명세서발행내역`
                ({string.Join(", ", colList)})
            VALUES
                ({string.Join(", ", paramList)})";
        foreach (var kv in @params)
            cmd.Parameters.AddWithValue(kv.Key, kv.Value ?? DBNull.Value);

        try
        {
            cmd.ExecuteNonQuery();
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = "SELECT LAST_INSERT_ID()";
            return Convert.ToInt32(idCmd.ExecuteScalar());
        }
        catch (Exception ex) { Debug.WriteLine($"[거래명세서] Insert 오류: {ex.Message}"); return -1; }
    }

    /// <summary>선택된 견적 목록의 거래명세서번호를 DB에 설정</summary>
    public static void SetTradeStatementNo(IEnumerable<int> rowids, string statementNo)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        foreach (var id in rowids)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `견적발행내역` SET `거래명세서번호`=@no WHERE {DbConnectionFactory.RowId}=@id";
            cmd.Parameters.AddWithValue("@no", statementNo);
            cmd.Parameters.AddWithValue("@id", id);
            try { cmd.ExecuteNonQuery(); } catch (Exception ex) { Debug.WriteLine($"[거래명세서] SetNo 오류: {ex.Message}"); }
        }
    }

    /// <summary>선택된 견적의 거래명세서번호를 DB에서 제거하고, 거래명세서발행내역 레코드도 삭제</summary>
    public static bool DeleteTradeStatement(string statementNo)
    {
        if (string.IsNullOrEmpty(statementNo)) return false;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 견적발행내역에서 거래명세서번호 제거
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "UPDATE `견적발행내역` SET `거래명세서번호`=NULL WHERE `거래명세서번호`=@no";
            cmd.Parameters.AddWithValue("@no", statementNo);
            cmd.ExecuteNonQuery();
        }

        // 2. 거래명세서발행내역 레코드 삭제
        if (TableExists(conn, "거래명세서발행내역"))
        {
            using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = "DELETE FROM `거래명세서발행내역` WHERE `거래명세서번호`=@no";
            cmd2.Parameters.AddWithValue("@no", statementNo);
            cmd2.ExecuteNonQuery();
        }

        Log($"거래명세서 삭제: {statementNo}");
        return true;
    }

    /// <summary>
    /// 선택된 견적들의 항목별 합산 데이터를 거래명세서 INSERT용 dict로 반환한다.
    /// key: 항목명, value: (합산수량, 단가, 합산소계)
    /// </summary>
    public static Dictionary<string, (decimal qty, decimal unitPrice, decimal subtotal)>
        BuildTradeStatementItemData(IEnumerable<QuotationIssue> issues)
    {
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "_id","rowid","견적발행일자","업체명","약칭","대표자",
            "견적요청담당","담당자","담당자연락처","담당자 e-Mail",
            "시료명","견적번호","적용구분","적용구분_코드",
            "합계 금액","부가세","총합계","비고","거래명세서번호"
        };

        var map = new Dictionary<string, (decimal qty, decimal unitPrice, decimal subtotal)>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var issue in issues)
        {
            var row = GetIssueRow(issue.Id);
            foreach (var kv in row)
            {
                var col = kv.Key.Trim();
                if (fixedCols.Contains(col)) continue;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (!decimal.TryParse(kv.Value.Replace(",",""), out var qty) || qty == 0) continue;

                decimal up = 0;
                if (row.TryGetValue(col + "단가", out var upStr))
                    decimal.TryParse(upStr.Replace(",",""), out up);
                decimal sub = qty * up;

                if (map.ContainsKey(col))
                    map[col] = (map[col].qty + qty, up, map[col].subtotal + sub);
                else
                    map[col] = (qty, up, sub);
            }
        }
        return map;
    }

    /// <summary>선택된 견적들의 항목별 합산 수량/금액 목록을 반환한다 (Show2 미리보기용).</summary>
    public static List<(string 항목, int 수량, decimal 금액)> AggregateIssueItems(
        IEnumerable<QuotationIssue> issues)
    {
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "_id","rowid","견적발행일자","업체명","약칭","대표자",
            "견적요청담당","담당자","담당자연락처","담당자 e-Mail",
            "시료명","견적번호","적용구분","적용구분_코드",
            "합계 금액","부가세","총합계","비고","거래명세서번호"
        };

        var itemMap = new Dictionary<string, (decimal qty, decimal unitPrice, decimal subtotal)>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var issue in issues)
        {
            var row = GetIssueRow(issue.Id);
            foreach (var kv in row)
            {
                var col = kv.Key.Trim();
                if (fixedCols.Contains(col)) continue;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;

                if (!decimal.TryParse(kv.Value.Replace(",",""), out var qty) || qty == 0) continue;

                decimal up = 0;
                if (row.TryGetValue(col + "단가", out var upStr))
                    decimal.TryParse(upStr.Replace(",",""), out up);
                decimal sub = qty * up;

                if (itemMap.ContainsKey(col))
                    itemMap[col] = (itemMap[col].qty + qty, up, itemMap[col].subtotal + sub);
                else
                    itemMap[col] = (qty, up, sub);
            }
        }

        return itemMap
            .Where(kv => kv.Value.qty > 0)
            .Select(kv => (항목: kv.Key, 수량: (int)kv.Value.qty, 금액: kv.Value.subtotal))
            .OrderBy(x => x.항목)
            .ToList();
    }

    /// <summary>견적서 1건을 템플릿 기반으로 Excel 파일에 출력한다.</summary>
    /// <returns>(성공 여부, 오류 메시지)</returns>
    public static (bool ok, string msg) ExportQuotation(QuotationIssue issue, string savePath)
    {
        try
        {
            // 항목 데이터 로드
            var row = GetIssueRow(issue.Id);
            var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id","rowid","견적발행일자","업체명","약칭","대표자",
                "견적요청담당","담당자","담당자연락처","담당자 e-Mail",
                "시료명","견적번호","적용구분","적용구분_코드",
                "합계 금액","부가세","총합계","비고","거래명세서번호"
            };

            var items = new List<(string name, decimal qty, decimal unitPrice)>();
            foreach (var kv in row)
            {
                var col = kv.Key.Trim();
                if (fixedCols.Contains(col)) continue;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (!decimal.TryParse(kv.Value.Replace(",",""), out var qty) || qty == 0) continue;
                decimal up = 0;
                if (row.TryGetValue(col + "단가", out var upStr))
                    decimal.TryParse(upStr.Replace(",",""), out up);
                items.Add((col, qty, up));
            }
            items.Sort((a, b) => string.Compare(a.name, b.name, StringComparison.OrdinalIgnoreCase));

            // 템플릿 선택 (33개 이하 → 견적서1, 34개 이상 → 견적서2)
            bool useTemplate2 = items.Count > 33;
            string tplPath = TemplatePath(useTemplate2 ? "견적서2.xlsx" : "견적서1.xlsx");
            if (!File.Exists(tplPath))
                return (false, $"템플릿 파일을 찾을 수 없습니다: {tplPath}");

            using var wb = new XLWorkbook(tplPath);
            var ws = wb.Worksheet("견적서");

            // ── 헤더 입력 ────────────────────────────────────────────────
            ws.Cell(3, 3).Value = issue.업체명;
            ws.Cell(4, 3).Value = "";
            ws.Cell(5, 3).Value = issue.담당자;
            ws.Cell(6, 3).Value = issue.담당자연락처;
            ws.Cell(7, 3).Value = issue.담당자이메일;
            ws.Cell(8, 3).Value = issue.시료명;
            ws.Cell(3, 8).Value = issue.견적번호;
            ws.Cell(4, 8).Value = issue.발행일;

            // ── 항목 입력 ────────────────────────────────────────────────
            if (!useTemplate2)
            {
                for (int i = 0; i < Math.Min(items.Count, 33); i++)
                {
                    int r = 11 + i;
                    ws.Cell(r, 1).Value = i + 1;
                    ws.Cell(r, 4).Value = items[i].name;
                    ws.Cell(r, 8).Value = (double)items[i].qty;
                    ws.Cell(r, 9).Value = (double)items[i].unitPrice;
                }
            }
            else
            {
                int leftCount  = Math.Min(items.Count, 33);
                int rightCount = Math.Min(items.Count - leftCount, 36);
                for (int i = 0; i < leftCount; i++)
                {
                    int r = 11 + i;
                    ws.Cell(r, 1).Value = i + 1;
                    ws.Cell(r, 4).Value = items[i].name;
                    ws.Cell(r, 8).Value = (double)items[i].qty;
                    ws.Cell(r, 9).Value = (double)items[i].unitPrice;
                }
                for (int i = 0; i < rightCount; i++)
                {
                    int r = 4 + i;
                    ws.Cell(r, 11).Value = leftCount + i + 1;
                    ws.Cell(r, 14).Value = items[leftCount + i].name;
                    ws.Cell(r, 18).Value = (double)items[leftCount + i].qty;
                    ws.Cell(r, 19).Value = (double)items[leftCount + i].unitPrice;
                }
            }

            wb.SaveAs(savePath);
            Log($"ExportQuotation: {(useTemplate2 ? "견적서2" : "견적서1")}, 항목={items.Count}개, 저장={savePath}");
            return (true, "");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ExportQuotation] 오류: {ex.Message}");
            return (false, ex.Message);
        }
    }

    private static string TemplatePath(string fileName)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        return Path.Combine(root, "Data", "Templates", fileName);
    }

    /// <summary>선택된 견적들의 데이터를 합산해 거래명세서 Excel을 생성한다 (템플릿 사용).</summary>
    /// <returns>(성공 여부, 오류 메시지, 공급가액, 부가세, 합계)</returns>
    public static (bool ok, string msg, decimal supply, decimal vat, decimal total) ExportTradingStatement(
        IEnumerable<QuotationIssue> issues, string savePath)
    {
        var issueList = issues.ToList();
        if (issueList.Count == 0) return (false, "선택된 견적이 없습니다.", 0, 0, 0);

        try
        {
            // ── 항목 집계 ────────────────────────────────────────────────────
            var itemMap = BuildTradeStatementItemData(issueList);
            var items   = itemMap
                .Where(kv => kv.Value.qty > 0)
                .OrderBy(kv => kv.Key)
                .ToList();

            decimal supply     = items.Sum(kv => kv.Value.subtotal);
            decimal vat        = Math.Round(supply * 0.1m, 0);
            decimal grandTotal = supply + vat;

            var first = issueList.First();
            string companyName = first.업체명;
            string abbr        = first.약칭;
            string manager     = first.담당자;
            string phone       = first.담당자연락처;
            string email       = first.담당자이메일;
            string sampleNames = string.Join(", ", issueList.Select(i => i.시료명).Where(s => !string.IsNullOrEmpty(s)).Distinct());
            string quotNos     = string.Join(", ", issueList.Select(i => i.견적번호));

            // ── 템플릿 선택 ──────────────────────────────────────────────────
            // 템플릿1: 항목 33개 이하 (항목행 11~43, J열 소계)
            // 템플릿2: 항목 34개 이상 (좌: 11~43 / 우: 4~39, 최대 69개)
            bool useTemplate2 = items.Count > 33;
            string tplFile    = useTemplate2 ? "거래명세서2.xlsx" : "거래명세서1.xlsx";
            string tplPath    = TemplatePath(tplFile);
            if (!File.Exists(tplPath))
                return (false, $"템플릿 파일을 찾을 수 없습니다: {tplPath}", 0, 0, 0);

            using var wb = new XLWorkbook(tplPath);
            var ws = wb.Worksheet("거래명세서");

            // ── 헤더 입력 ────────────────────────────────────────────────────
            // 좌측 헤더 (거래명세서1/2 공통, A~J 섹션)
            ws.Cell(3, 3).Value = companyName;    // 업체명
            ws.Cell(4, 3).Value = "";              // 대표자 (정보 없음)
            ws.Cell(5, 3).Value = manager;         // 견적요청담당자
            ws.Cell(6, 3).Value = phone;           // 담당자연락처
            ws.Cell(7, 3).Value = email;           // 담당자 e-Mail
            ws.Cell(8, 3).Value = sampleNames;     // 시료명

            // 우측 헤더
            ws.Cell(3, 8).Value = "";                                    // 거래명세서번호 — 발행 후 DB에서 채번되므로 빈값
            ws.Cell(4, 8).Value = DateTime.Today.ToString("yyyy-MM-dd"); // 발행일자

            // ── 항목 입력 ────────────────────────────────────────────────────
            if (!useTemplate2)
            {
                // 템플릿1: 행 11~43 (최대 33행), H=수량, I=단가, J=수식(소계)
                for (int i = 0; i < Math.Min(items.Count, 33); i++)
                {
                    int row = 11 + i;
                    var (name, (qty, up, _)) = items[i];
                    ws.Cell(row, 1).Value = i + 1;   // 번호
                    ws.Cell(row, 4).Value = name;     // 시험항목 (D열, D:E 병합)
                    ws.Cell(row, 8).Value = (double)qty;  // H 수량
                    ws.Cell(row, 9).Value = (double)up;   // I 단가
                }
            }
            else
            {
                // 템플릿2: 좌 11~43(33개), 우 4~39(36개)
                int leftCount  = Math.Min(items.Count, 33);
                int rightCount = Math.Min(items.Count - leftCount, 36);

                for (int i = 0; i < leftCount; i++)
                {
                    int row = 11 + i;
                    var (name, (qty, up, _)) = items[i];
                    ws.Cell(row, 1).Value = i + 1;
                    ws.Cell(row, 4).Value = name;
                    ws.Cell(row, 8).Value = (double)qty;
                    ws.Cell(row, 9).Value = (double)up;
                }
                for (int i = 0; i < rightCount; i++)
                {
                    int row = 4 + i;
                    var (name, (qty, up, _)) = items[leftCount + i];
                    ws.Cell(row, 11).Value = leftCount + i + 1;  // K 번호
                    ws.Cell(row, 14).Value = name;                // N 시험항목 (N:O 병합)
                    ws.Cell(row, 18).Value = (double)qty;         // R 수량
                    ws.Cell(row, 19).Value = (double)up;          // S 단가
                }
            }

            wb.SaveAs(savePath);
            Log($"ExportTradingStatement: {tplFile} 사용, 항목={items.Count}개, 저장={savePath}");
            return (true, "", supply, vat, grandTotal);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[거래명세서] Excel 생성 오류: {ex.Message}");
            return (false, ex.Message, 0, 0, 0);
        }
    }
}
