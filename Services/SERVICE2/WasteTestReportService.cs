using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 배출업소 시험성적서 서비스
/// — 비용부담금_결과 테이블에서 날짜별 결과 조회
/// — 이전 N회 평균 대비 추세 계산 (단계별: ▲ ▲▲ ↑)
/// </summary>
public static class WasteTestReportService
{
    // ── 날짜 목록 (채수일 distinct, 내림차순) ─────────────────────────────
    public static List<string> GetDates(List<string> groups)
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        if (groups.Count == 0)
        {
            cmd.CommandText = "SELECT DISTINCT 채수일 FROM `비용부담금_결과` ORDER BY 채수일 DESC";
        }
        else
        {
            var placeholders = string.Join(",", groups.Select((_, i) => $"@g{i}"));
            cmd.CommandText = $"SELECT DISTINCT 채수일 FROM `비용부담금_결과` WHERE 구분 IN ({placeholders}) ORDER BY 채수일 DESC";
            for (int i = 0; i < groups.Count; i++)
                cmd.Parameters.AddWithValue($"@g{i}", groups[i]);
        }
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var raw = r.GetValue(0);
            string d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
            if (!string.IsNullOrEmpty(d)) list.Add(d);
        }
        return list;
    }

    // ── 날짜별 구분별 건수 ───────────────────────────────────────────────
    public static Dictionary<string, Dictionary<string, int>> GetDateGroupCounts()
    {
        var result = new Dictionary<string, Dictionary<string, int>>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 채수일, 구분, COUNT(*) FROM `비용부담금_결과` GROUP BY 채수일, 구분";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var raw = r.GetValue(0);
            string d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
            string g = r.GetString(1);
            int cnt = r.GetInt32(2);
            if (string.IsNullOrEmpty(d)) continue;
            if (!result.ContainsKey(d)) result[d] = new Dictionary<string, int>();
            result[d][g] = cnt;
        }
        return result;
    }

    // ── 업체명 → 약칭 매핑 ──────────────────────────────────────────────
    public static Dictionary<string, string> GetCompanyAbbreviations()
    {
        var map = new Dictionary<string, string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        foreach (var table in new[] { "폐수배출업소", "여수_폐수배출업소" })
        {
            if (!DbConnectionFactory.TableExists(conn, table)) continue;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT 업체명, COALESCE(약칭, '') FROM `{table}`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var name = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                var abbr = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(abbr))
                    map[name] = abbr;
            }
        }
        return map;
    }

    // ── 특정 날짜의 전체 결과 조회 ────────────────────────────────────────
    public static List<WasteTestRow> GetByDate(string 채수일, List<string> groups)
    {
        var list = new List<WasteTestRow>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        var where = "채수일=@d";
        if (groups.Count > 0)
        {
            var placeholders = string.Join(",", groups.Select((_, i) => $"@g{i}"));
            where += $" AND 구분 IN ({placeholders})";
        }
        cmd.CommandText = $@"
            SELECT Id, 순서, SN, 업체명, 관리번호, 구분,
                   BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols, 비고, 확인자
            FROM `비용부담금_결과`
            WHERE {where}
            ORDER BY 구분, 순서, 업체명";
        cmd.Parameters.AddWithValue("@d", 채수일);
        for (int i = 0; i < groups.Count; i++)
            cmd.Parameters.AddWithValue($"@g{i}", groups[i]);
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new WasteTestRow
            {
                Id       = Convert.ToInt32(r["Id"]),
                순서     = r["순서"] is DBNull ? 0 : Convert.ToInt32(r["순서"]),
                SN       = r["SN"]?.ToString() ?? "",
                업체명   = r["업체명"]?.ToString() ?? "",
                관리번호 = r["관리번호"]?.ToString() ?? "",
                구분     = r["구분"]?.ToString() ?? "",
                BOD      = r["BOD"]?.ToString() ?? "",
                TOC      = r["TOC"]?.ToString() ?? "",
                SS       = r["SS"]?.ToString() ?? "",
                TN       = r["T-N"]?.ToString() ?? "",
                TP       = r["T-P"]?.ToString() ?? "",
                NHexan   = r["N-Hexan"]?.ToString() ?? "",
                Phenols  = r["Phenols"]?.ToString() ?? "",
                비고     = r["비고"]?.ToString() ?? "",
                확인자   = r["확인자"]?.ToString() ?? "",
            });
        }
        return list;
    }

    // ── 추세 계산: 업체별 이전 N회 평균 대비 변화율 반환 ──────────────────
    public static Dictionary<string, Dictionary<string, TrendInfo>> GetTrends(
        string 채수일, List<string> groups, int historyCount = 10)
    {
        var result = new Dictionary<string, Dictionary<string, TrendInfo>>();
        var currentRows = GetByDate(채수일, groups);
        if (currentRows.Count == 0) return result;

        var items = new[] { "BOD", "TOC", "SS", "TN", "TP", "NHexan", "Phenols" };

        var allDates = GetDates(groups);
        int idx = allDates.IndexOf(채수일);
        if (idx < 0) return result;
        var prevDates = allDates.Skip(idx + 1).Take(historyCount * 2).ToList();

        var historyByCompany = new Dictionary<string, List<WasteTestRow>>();
        foreach (var d in prevDates)
        {
            var rows = GetByDate(d, groups);
            foreach (var row in rows)
            {
                if (!historyByCompany.ContainsKey(row.업체명))
                    historyByCompany[row.업체명] = new List<WasteTestRow>();
                historyByCompany[row.업체명].Add(row);
            }
        }

        foreach (var cur in currentRows)
        {
            var trends = new Dictionary<string, TrendInfo>();
            if (!historyByCompany.TryGetValue(cur.업체명, out var history))
            {
                result[cur.업체명] = trends;
                continue;
            }

            var recent = history.Take(historyCount).ToList();
            if (recent.Count == 0) { result[cur.업체명] = trends; continue; }

            foreach (var item in items)
            {
                var curVal = GetValue(cur, item);
                if (curVal == null) continue;

                var prevVals = recent.Select(r => GetValue(r, item))
                    .Where(v => v.HasValue).Select(v => v!.Value).ToList();
                if (prevVals.Count == 0) continue;

                double avg = prevVals.Average();
                if (avg <= 0) continue;

                double pctChange = (curVal.Value - avg) / avg * 100.0;
                trends[item] = new TrendInfo
                {
                    Value = curVal.Value,
                    Average = avg,
                    PctChange = pctChange,
                };
            }
            result[cur.업체명] = trends;
        }
        return result;
    }

    private static double? GetValue(WasteTestRow row, string item) => item switch
    {
        "BOD"     => TryParse(row.BOD),
        "TOC"     => TryParse(row.TOC),
        "SS"      => TryParse(row.SS),
        "TN"      => TryParse(row.TN),
        "TP"      => TryParse(row.TP),
        "NHexan"  => TryParse(row.NHexan),
        "Phenols" => TryParse(row.Phenols),
        _         => null,
    };

    private static double? TryParse(string? s) =>
        string.IsNullOrWhiteSpace(s) ? null : double.TryParse(s, out var v) ? v : null;

    /// <summary>항목별 표시 자릿수 포맷 (VBA 기준)</summary>
    internal static string GetFormat(string item) => item switch
    {
        "BOD" or "TOC" or "SS" or "NHexan" => "F1",  // 소수 1자리
        "TN" or "TP" or "Phenols"          => "F3",  // 소수 3자리
        _                                  => "F1",
    };
}

// ── 추세 정보 ────────────────────────────────────────────────────────────
public class TrendInfo
{
    public double Value     { get; set; }
    public double Average   { get; set; }
    public double PctChange { get; set; } // 양수=상승, 음수=하락
}

// ── 결과 행 모델 ────────────────────────────────────────────────────────
public class WasteTestRow
{
    public int    Id       { get; set; }
    public int    순서     { get; set; }
    public string SN       { get; set; } = "";
    public string 업체명   { get; set; } = "";
    public string 관리번호 { get; set; } = "";
    public string 구분     { get; set; } = "";
    public string BOD      { get; set; } = "";
    public string TOC      { get; set; } = "";
    public string SS       { get; set; } = "";
    public string TN       { get; set; } = "";
    public string TP       { get; set; } = "";
    public string NHexan   { get; set; } = "";
    public string Phenols  { get; set; } = "";
    public string 비고     { get; set; } = "";
    public string 확인자   { get; set; } = "";
}
