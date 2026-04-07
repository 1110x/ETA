using System.Data;
using System.Data.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class MeasurerService
{
    // ── 메모리 캐시 ──────────────────────────────────────────────────────────
    private static List<string>? _companiesCache;
    private static List<(string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명)>? _allDataCache;
    private static List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)>? _analysisItemsCache;
    public static void InvalidateCache() { _companiesCache = null; _allDataCache = null; }
    public static void InvalidateAnalysisItemsCache() => _analysisItemsCache = null;

    // ── 테이블 생성 + 컬럼 마이그레이션 ─────────────────────────────────────
    public static void EnsureMeasurerTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 기본 테이블 생성
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS 측정인_채취지점 (
                    Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    계약번호   TEXT,
                    약칭       TEXT,
                    계약기간   TEXT,
                    업체명     TEXT NOT NULL,
                    채취지점명 TEXT NOT NULL,
                    UNIQUE(업체명, 채취지점명)
                );";
            cmd.ExecuteNonQuery();
        }

        // 기존 DB 마이그레이션 — 컬럼이 없으면 추가
        foreach (var col in new[] { ("계약번호", "TEXT"), ("약칭", "TEXT"), ("계약기간", "TEXT"),
                                    ("계약금액", "TEXT"), ("견적구분", "TEXT") })
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE 측정인_채취지점 ADD COLUMN {col.Item1} {col.Item2}";
                cmd.ExecuteNonQuery();
            }
            catch { /* 이미 존재하면 무시 */ }
        }
    }

    // ── 분석장비 테이블 생성 / 마이그레이션 ─────────────────────────────────
    public static void EnsureEquipmentTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS 측정인_분석장비 (
                Id       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                장비명   TEXT NOT NULL,
                코드값   TEXT NOT NULL,
                UNIQUE(코드값)
            );";
        cmd.ExecuteNonQuery();
    }

    // ── 분석장비 일괄 저장 (UPSERT) ─────────────────────────────────────────
    public static int SaveEquipments(List<(string 장비명, string 코드값)> items)
    {
        EnsureEquipmentTable();
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var (name, code) in items)
            {
                if (string.IsNullOrWhiteSpace(code)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(new[] { "코드값" }, new[] { "장비명" });
                cmd.CommandText = $"INSERT INTO 측정인_분석장비 (장비명, 코드값) VALUES (@name, @code) {upsert};";
                cmd.Parameters.AddWithValue("@name", name.Trim());
                cmd.Parameters.AddWithValue("@code", code.Trim());
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
            Debug.WriteLine($"[MeasurerService] SaveEquipments: {count}/{items.Count}개 저장");
        }
        catch (Exception ex) { txn.Rollback(); Debug.WriteLine($"[MeasurerService] SaveEquipments 오류: {ex.Message}"); throw; }
        return count;
    }

    // ── 분석장비 전체 조회 ───────────────────────────────────────────────────
    public static List<(string 장비명, string 코드값)> GetAllEquipments()
    {
        EnsureEquipmentTable();
        var list = new List<(string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 장비명, 코드값 FROM 측정인_분석장비 ORDER BY 장비명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1)));
        return list;
    }

    // ── 분석항목 테이블 생성 / 마이그레이션 ─────────────────────────────────
    public static void EnsureAnalysisItemTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS 측정인_분석항목 (
                Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                분야       TEXT NOT NULL DEFAULT '',
                항목구분   TEXT NOT NULL DEFAULT '',
                항목명     TEXT NOT NULL,
                코드값     TEXT NOT NULL,
                select2id  TEXT,
                UNIQUE(코드값)
            );";
        cmd.ExecuteNonQuery();
        // 기존 테이블 마이그레이션
        try { cmd.CommandText = "ALTER TABLE 측정인_분석항목 ADD COLUMN 분야 TEXT NOT NULL DEFAULT ''"; cmd.ExecuteNonQuery(); } catch { }
        try { cmd.CommandText = "ALTER TABLE 측정인_분석항목 ADD COLUMN 항목구분 TEXT NOT NULL DEFAULT ''"; cmd.ExecuteNonQuery(); } catch { }
    }

    // ── 분석항목 일괄 저장 (UPSERT) ─────────────────────────────────────────
    /// <param name="items">(분야, 항목구분, 항목명, 코드값, select2id) 목록</param>
    public static int SaveAnalysisItems(List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)> items)
    {
        EnsureAnalysisItemTable();
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var (field, category, name, code, s2id) in items)
            {
                if (string.IsNullOrWhiteSpace(code)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(
                    new[] { "코드값" },
                    new[] { "분야", "항목구분", "항목명", "select2id" });
                cmd.CommandText =
                    $"""
                    INSERT INTO 측정인_분석항목 (분야, 항목구분, 항목명, 코드값, select2id)
                    VALUES (@field, @category, @name, @code, @s2id)
                    {upsert};
                    """;
                cmd.Parameters.AddWithValue("@field",    field.Trim());
                cmd.Parameters.AddWithValue("@category", category.Trim());
                cmd.Parameters.AddWithValue("@name",     name.Trim());
                cmd.Parameters.AddWithValue("@code",     code.Trim());
                cmd.Parameters.AddWithValue("@s2id",     s2id.Trim());
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
            Debug.WriteLine($"[MeasurerService] SaveAnalysisItems: {{count}}/{{items.Count}}개 저장");
        }
        catch (Exception ex)
        {
            txn.Rollback();
            Debug.WriteLine($"[MeasurerService] SaveAnalysisItems 오류: {{ex.Message}}");
            throw;
        }
        return count;
    }

    // ── 분석항목 전체 조회 ───────────────────────────────────────────────────
    public static List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)> GetAllAnalysisItems()
    {
        if (_analysisItemsCache != null) return _analysisItemsCache;
        EnsureAnalysisItemTable();
        var list = new List<(string, string, string, string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 분야, 항목구분, 항목명, 코드값, select2id FROM 측정인_분석항목 ORDER BY 분야 ASC, 항목구분 ASC, 항목명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.IsDBNull(0) ? "" : r.GetString(0),
                      r.IsDBNull(1) ? "" : r.GetString(1),
                      r.GetString(2),
                      r.GetString(3),
                      r.IsDBNull(4) ? "" : r.GetString(4)));
        _analysisItemsCache = list;
        return list;
    }

    // ── 전체 데이터 수동 초기화 (필요 시 직접 호출 — 스크래핑 자동 흐름에서는 사용하지 않음) ──
    // SaveSamplingPoints가 UPSERT를 사용하므로 일반적으로 호출 불필요.
    public static void ClearAll()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM 측정인_채취지점";
        cmd.ExecuteNonQuery();
        InvalidateCache();
    }

    // ── 채취지점 데이터 저장 ─────────────────────────────────────────────────
    public static void SaveSamplingPoints(
        string companyName,
        List<string> points,
        string 계약번호 = "",
        string 계약기간 = "")
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var point in points.Where(p => !string.IsNullOrWhiteSpace(p)))
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(
                    new[] { "업체명", "채취지점명" },
                    new[] { "계약번호", "계약기간" });
                cmd.CommandText =
                    $"""
                    INSERT INTO 측정인_채취지점 (계약번호, 계약기간, 업체명, 채취지점명)
                    VALUES (@contractNo, @period, @company, @point)
                    {upsert};
                    """;
                cmd.Parameters.AddWithValue("@contractNo", 계약번호.Trim());
                cmd.Parameters.AddWithValue("@period",     계약기간.Trim());
                cmd.Parameters.AddWithValue("@company",    companyName.Trim());
                cmd.Parameters.AddWithValue("@point",      point.Trim());
                cmd.ExecuteNonQuery();
            }
            txn.Commit();
            InvalidateCache();
        }
        catch (Exception ex)
        {
            txn.Rollback();
            Debug.WriteLine($"[MeasurerService] 저장 오류: {ex.Message}");
            throw;
        }
    }

    // ── 업체별 관리 정보 조회 ────────────────────────────────────────────────
    public static (string 약칭, string 계약금액, string 견적구분) GetCompanyInfo(string companyName)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT IFNULL(약칭,''), IFNULL(계약금액,''), IFNULL(견적구분,'')
            FROM 측정인_채취지점 WHERE 업체명 = @company LIMIT 1";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        return r.Read()
            ? (r.GetString(0), r.GetString(1), r.GetString(2))
            : ("", "", "");
    }

    // ── 업체별 관리 정보 저장 ────────────────────────────────────────────────
    public static void UpdateCompanyInfo(string companyName, string 약칭, string 계약금액, string 견적구분)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE 측정인_채취지점
            SET 약칭 = @alias, 계약금액 = @amount, 견적구분 = @quotType
            WHERE 업체명 = @company";
        cmd.Parameters.AddWithValue("@alias",    약칭.Trim());
        cmd.Parameters.AddWithValue("@amount",   계약금액.Trim());
        cmd.Parameters.AddWithValue("@quotType", 견적구분.Trim());
        cmd.Parameters.AddWithValue("@company",  companyName.Trim());
        cmd.ExecuteNonQuery();
    }

    // ── 계약 DB에서 약칭을 측정인_채취지점으로 동기화 ──────────────────────────
    public static void SyncAbbrFromContractDb()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        // 약칭이 비어 있고 계약 DB에 같은 업체명이 있으면 약칭 복사
        cmd.CommandText = @"
            UPDATE 측정인_채취지점
            SET 약칭 = (
                SELECT C_Abbreviation
                FROM `계약 DB`
                WHERE `계약 DB`.C_CompanyName = 측정인_채취지점.업체명
                  AND C_Abbreviation IS NOT NULL
                  AND C_Abbreviation != ''
                LIMIT 1
            )
            WHERE (약칭 IS NULL OR 약칭 = '')
              AND EXISTS (
                SELECT 1 FROM `계약 DB`
                WHERE `계약 DB`.C_CompanyName = 측정인_채취지점.업체명
                  AND C_Abbreviation IS NOT NULL
                  AND C_Abbreviation != ''
              )";
        cmd.ExecuteNonQuery();
    }

    // ── 저장된 업체명 목록 조회 ──────────────────────────────────────────────
    public static List<string> GetCompanies()
    {
        if (_companiesCache != null) return _companiesCache;
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 업체명 FROM 측정인_채취지점 ORDER BY 업체명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        _companiesCache = list;
        return list;
    }

    // ── 특정 업체의 채취지점명 목록 조회 ─────────────────────────────────────
    public static List<string> GetSamplingPoints(string companyName)
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 채취지점명 FROM 측정인_채취지점 WHERE 업체명 = @company ORDER BY 채취지점명 ASC";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 업체 목록을 Contract 형태로 반환 (견적/의뢰서 Content4용) ─────────────
    public static List<ETA.Models.Contract> GetCompaniesAsContracts()
    {
        var map = new Dictionary<string, ETA.Models.Contract>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 업체명, IFNULL(약칭,''), IFNULL(계약기간,'')
            FROM 측정인_채취지점
            GROUP BY 업체명
            ORDER BY 업체명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            string name   = r.GetString(0);
            string abbr   = r.GetString(1);
            string period = r.GetString(2);

            DateTime? endDate = ParsePeriodEnd(period);
            map[name] = new ETA.Models.Contract
            {
                C_CompanyName  = name,
                C_Abbreviation = abbr,
                C_ContractEnd  = endDate,
            };
        }
        return [.. map.Values];
    }

    // "YYYY-MM-DD ~ YYYY-MM-DD" 또는 "YYYY.MM.DD~YYYY.MM.DD" 등 다양한 형식에서 종료일 파싱
    private static DateTime? ParsePeriodEnd(string period)
    {
        if (string.IsNullOrWhiteSpace(period)) return null;
        // 구분자 기준으로 뒷부분 추출
        var sep = new[] { " ~ ", "~", " - ", "–", "—" };
        string? endStr = null;
        foreach (var s in sep)
        {
            int idx = period.IndexOf(s, StringComparison.Ordinal);
            if (idx >= 0) { endStr = period[(idx + s.Length)..].Trim(); break; }
        }
        endStr ??= period.Trim();
        // 날짜 형식 파싱 시도
        var fmts = new[] { "yyyy-MM-dd", "yyyy.MM.dd", "yyyy/MM/dd", "yyyyMMdd" };
        foreach (var fmt in fmts)
            if (DateTime.TryParseExact(endStr, fmt, null,
                    System.Globalization.DateTimeStyles.None, out var dt))
                return dt;
        return null;
    }

    // ── 전체 데이터 조회 ──────────────────────────────────────────────────────
    public static List<(string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명)> GetAllData()
    {
        if (_allDataCache != null) return _allDataCache;
        var list = new List<(string, string, string, string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 계약번호, IFNULL(약칭,''), IFNULL(계약기간,''), 업체명, 채취지점명
            FROM 측정인_채취지점
            ORDER BY 업체명, 채취지점명";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4)));
        _allDataCache = list;
        return list;
    }

    /// <summary>
    /// 채취지점명을 기준으로 계약번호를 직접 조회합니다.
    /// sampleName이 DB의 채취지점명과 일치하는 행의 계약번호를 반환합니다.
    /// </summary>
    public static (string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명) FindContractBySamplingPoint(
        string sampleName,
        string abbr = "")
    {
        var all = GetAllData();
        if (all.Count == 0 || string.IsNullOrWhiteSpace(sampleName))
            return ("", "", "", "", "");

        string normSample = NormalizeCompany(sampleName);
        string normAbbr   = NormalizeCompany(abbr);

        var best = all
            .Select(item =>
            {
                string itemSite = NormalizeCompany(item.채취지점명);
                string itemAbbr = NormalizeCompany(item.약칭);

                int score = 0;
                if (itemSite == normSample) score += 1000;
                else if (itemSite.Contains(normSample, StringComparison.OrdinalIgnoreCase)
                      || normSample.Contains(itemSite, StringComparison.OrdinalIgnoreCase))
                    score += 500;

                if (!string.IsNullOrWhiteSpace(normAbbr))
                {
                    if (itemAbbr == normAbbr) score += 200;
                    else if (itemAbbr.Contains(normAbbr, StringComparison.OrdinalIgnoreCase)
                          || normAbbr.Contains(itemAbbr, StringComparison.OrdinalIgnoreCase))
                        score += 80;
                }
                if (!string.IsNullOrWhiteSpace(item.계약번호)) score += 20;

                return (item, score);
            })
            .Where(x => x.score > 0)
            .OrderByDescending(x => x.score)
            .FirstOrDefault();

        return best.score > 0 ? best.item : ("", "", "", "", "");
    }

    public static (string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명) FindBestContract(
        string companyName,
        string workSite,
        string abbr = "")
    {
        var all = GetAllData();
        if (all.Count == 0)
            return ("", "", "", "", "");

        string normCompany = NormalizeCompany(companyName);
        string normSite = NormalizeCompany(workSite);
        string normAbbr = NormalizeCompany(abbr);

        int Score(string source, string query, int exactScore, int containsScore)
        {
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(query)) return 0;
            if (source == query) return exactScore;
            if (source.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                query.Contains(source, StringComparison.OrdinalIgnoreCase))
                return containsScore;
            return 0;
        }

        var best = all
            .Select(item =>
            {
                string itemCompany = NormalizeCompany(item.업체명);
                string itemSite = NormalizeCompany(item.채취지점명);
                string itemAbbr = NormalizeCompany(item.약칭);

                int score = 0;
                score += Score(itemCompany, normCompany, 500, 280);
                score += Score(itemSite, normSite, 320, 180);
                score += Score(itemAbbr, normAbbr, 120, 60);

                if (!string.IsNullOrWhiteSpace(normCompany) && !string.IsNullOrWhiteSpace(normSite))
                {
                    if ((itemCompany + itemSite).Contains(normCompany + normSite, StringComparison.OrdinalIgnoreCase) ||
                        (normCompany + normSite).Contains(itemCompany + itemSite, StringComparison.OrdinalIgnoreCase))
                        score += 140;
                }

                if (!string.IsNullOrWhiteSpace(item.계약번호))
                    score += 20;

                return (item, score);
            })
            .OrderByDescending(x => x.score)
            .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.item.계약번호))
            .FirstOrDefault();

        return best.score > 0 ? best.item : ("", "", "", "", "");
    }

    private static string NormalizeCompany(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "";

        return name
            .Replace("㈜", "주")
            .Replace("(주)", "주")
            .Replace("（주）", "주")
            .Replace("주식회사", "주")
            .Replace("유한회사", "유")
            .Replace("㈔", "사")
            .Replace("(사)", "사")
            .Replace(" ", "")
            .Replace("-", "")
            .Replace("·", "")
            .Replace("(", "")
            .Replace(")", "")
            .Replace("（", "")
            .Replace("）", "")
            .ToLowerInvariant()
            .Trim();
    }
}
