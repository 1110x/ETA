using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteDataService
{
    // ── 업체명으로 전 항목 분석결과 조회 (폐수의뢰및결과 테이블, 날짜 오름차순) ──
    public static List<WasteAnalysisResult> GetResults(string 업체명)
    {
        var list = new List<WasteAnalysisResult>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 채수일, BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols
            FROM `폐수의뢰및결과`
            WHERE 업체명 = @n
            ORDER BY 채수일 ASC";
        cmd.Parameters.AddWithValue("@n", 업체명);
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var dateRaw = r.IsDBNull(0) ? "" : r.GetValue(0);
            string date = dateRaw is DateTime dt ? dt.ToString("yyyy-MM-dd") : dateRaw?.ToString() ?? "";
            if (string.IsNullOrEmpty(date)) continue;

            list.Add(new WasteAnalysisResult
            {
                채수일   = date,
                BOD      = ParseDouble(r, 1),
                TOC_TCIC = ParseDouble(r, 2),
                SS       = ParseDouble(r, 3),
                TN       = ParseDouble(r, 4),
                TP       = ParseDouble(r, 5),
                NHexan   = ParseDouble(r, 6),
                Phenols  = ParseDouble(r, 7),
            });
        }
        return list;
    }

    private static double? ParseDouble(System.Data.Common.DbDataReader r, int i)
    {
        if (r.IsDBNull(i)) return null;
        var s = r.GetValue(i)?.ToString() ?? "";
        return string.IsNullOrWhiteSpace(s) ? null : double.TryParse(s, out var v) ? v : null;
    }

    // ── DATA 테이블 업체명 일괄 변경 (8개 테이블 전체) ───────────────────────
    public static int RenameCompany(string oldName, string newName)
    {
        if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return 0;
        int total = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        foreach (var table in new[]
        {
            "BOD_DATA", "SS_DATA", "TN_DATA", "TP_DATA",
            "TOC_TCIC_DATA", "TOC_NPOC_DATA", "Phenols_DATA", "NHexan_DATA"
        })
        {
            using var cmd = conn.CreateCommand();
            // 이름 변경 (중복 키 충돌 시 건너뜀)
            cmd.CommandText = $"UPDATE IGNORE `{table}` SET 업체명=@newName WHERE 업체명 COLLATE utf8mb4_general_ci = @oldName COLLATE utf8mb4_general_ci";
            cmd.Parameters.AddWithValue("@newName", newName);
            cmd.Parameters.AddWithValue("@oldName", oldName);
            total += cmd.ExecuteNonQuery();

            // UPDATE IGNORE로 건너뛴 중복 행 삭제 (동일 데이터가 newName으로 이미 존재)
            using var delCmd = conn.CreateCommand();
            delCmd.CommandText = $"DELETE FROM `{table}` WHERE 업체명 COLLATE utf8mb4_general_ci = @oldName COLLATE utf8mb4_general_ci";
            delCmd.Parameters.AddWithValue("@oldName", oldName);
            delCmd.ExecuteNonQuery();
        }
        return total;
    }

    // ── DATA 테이블 전체 업체명 목록 ─────────────────────────────────────────
    public static List<string> GetDataCompanyNames()
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT DISTINCT 업체명 FROM (
                SELECT 업체명 COLLATE utf8mb4_general_ci AS 업체명 FROM `BOD_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `SS_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `TN_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `TP_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `TOC_TCIC_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `TOC_NPOC_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `Phenols_DATA`
                UNION SELECT 업체명 COLLATE utf8mb4_general_ci FROM `NHexan_DATA`
            ) t
            ORDER BY 업체명";
        using var r = cmd.ExecuteReader();
        while (r.Read()) if (!r.IsDBNull(0)) list.Add(r.GetString(0));
        return list;
    }

    // ── 업체별 TOC 결과 (TC-IC + NPOC, *_DATA 테이블에서 직접 조회) ─────
    public static List<WasteAnalysisResult> GetTocResults(string 업체명)
    {
        // 날짜 → (TCIC, NPOC) 매핑
        var map = new Dictionary<string, (double? tcic, double? npoc)>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // TOC(TC-IC) from TOC_TCIC_DATA
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT 분석일, `검량선_a` FROM `TOC_TCIC_DATA` WHERE 업체명=@n ORDER BY 분석일";
            cmd.Parameters.AddWithValue("@n", 업체명);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var raw = r.IsDBNull(0) ? "" : r.GetValue(0);
                string d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
                if (string.IsNullOrEmpty(d)) continue;
                double? v = r.IsDBNull(1) ? null : ParseStr(r.GetValue(1)?.ToString());
                if (!map.ContainsKey(d)) map[d] = (null, null);
                var cur = map[d];
                map[d] = (v ?? cur.tcic, cur.npoc);
            }
        }

        // TOC(NPOC) from TOC_NPOC_DATA
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT 분석일, `검량선_a` FROM `TOC_NPOC_DATA` WHERE 업체명=@n ORDER BY 분석일";
            cmd.Parameters.AddWithValue("@n", 업체명);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var raw = r.IsDBNull(0) ? "" : r.GetValue(0);
                string d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
                if (string.IsNullOrEmpty(d)) continue;
                double? v = r.IsDBNull(1) ? null : ParseStr(r.GetValue(1)?.ToString());
                if (!map.ContainsKey(d)) map[d] = (null, null);
                var cur = map[d];
                map[d] = (cur.tcic, v ?? cur.npoc);
            }
        }

        var list = new List<WasteAnalysisResult>();
        foreach (var (date, (tcic, npoc)) in map.OrderBy(kv => kv.Key))
        {
            list.Add(new WasteAnalysisResult
            {
                채수일    = date,
                TOC_TCIC = tcic,
                TOC_NPOC = npoc,
            });
        }
        return list;
    }

    private static double? ParseStr(string? s)
        => string.IsNullOrWhiteSpace(s) ? null : double.TryParse(s, out var v) ? v : null;

    private static void Fetch(DbConnection conn, string sql, string 업체명, Action<string, double?> set)
    {
        var rows = new List<(string d, double? v)>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@n", 업체명);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var raw = r.IsDBNull(0) ? "" : r.GetValue(0);
                var d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
                double? v = r.IsDBNull(1) ? null : Convert.ToDouble(r.GetValue(1));
                if (!string.IsNullOrEmpty(d)) rows.Add((d, v));
            }
        }
        foreach (var (d, v) in rows) set(d, v);
    }
}
