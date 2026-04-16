using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteDataService
{
    // ── 업체명으로 전 항목 분석결과 조회 (비용부담금_결과 테이블, 날짜 오름차순) ──
    public static List<WasteAnalysisResult> GetResults(string 업체명)
    {
        var list = new List<WasteAnalysisResult>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 채수일, BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols
            FROM `비용부담금_결과`
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

    // ── 시험기록부 테이블 업체명 일괄 변경 ───────────────────────────────────
    public static int RenameCompany(string oldName, string newName)
    {
        if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return 0;
        int total = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var tables = new List<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA=DATABASE() AND COLUMN_NAME='업체명'
                    AND TABLE_NAME LIKE '%_시험기록부'";
                using var r = cmd.ExecuteReader();
                while (r.Read()) tables.Add(r.GetString(0));
            }
            foreach (var table in tables)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"UPDATE IGNORE `{table}` SET 업체명=@newName WHERE 업체명 COLLATE utf8mb4_general_ci = @oldName COLLATE utf8mb4_general_ci";
                cmd.Parameters.AddWithValue("@newName", newName);
                cmd.Parameters.AddWithValue("@oldName", oldName);
                total += cmd.ExecuteNonQuery();
                using var delCmd = conn.CreateCommand();
                delCmd.CommandText = $"DELETE FROM `{table}` WHERE 업체명 COLLATE utf8mb4_general_ci = @oldName COLLATE utf8mb4_general_ci";
                delCmd.Parameters.AddWithValue("@oldName", oldName);
                delCmd.ExecuteNonQuery();
            }
        }
        catch { }
        return total;
    }

    // ── 시험기록부 테이블 전체 업체명 목록 ───────────────────────────────────
    public static List<string> GetDataCompanyNames()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var tables = new List<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA=DATABASE() AND COLUMN_NAME='업체명'
                    AND TABLE_NAME LIKE '%_시험기록부'";
                using var r = cmd.ExecuteReader();
                while (r.Read()) tables.Add(r.GetString(0));
            }
            if (tables.Count == 0) return list;
            var unions = string.Join("\nUNION ", tables.Select(t => $"SELECT 업체명 COLLATE utf8mb4_general_ci AS 업체명 FROM `{t}`"));
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT DISTINCT 업체명 FROM ({unions}) t WHERE 업체명 IS NOT NULL AND 업체명 <> '' ORDER BY 업체명";
                using var r = cmd.ExecuteReader();
                while (r.Read()) if (!r.IsDBNull(0)) list.Add(r.GetString(0));
            }
        }
        catch { }
        return list;
    }

    // ── 업체별 TOC 결과 ───────────────────────────────────────────────────────
    public static List<WasteAnalysisResult> GetTocResults(string 업체명)
    {
        var list = new List<WasteAnalysisResult>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 분석일, 검량선_a FROM `총_유기탄소_시험기록부` WHERE 업체명=@n ORDER BY 분석일";
            cmd.Parameters.AddWithValue("@n", 업체명);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var raw = r.IsDBNull(0) ? "" : r.GetValue(0);
                string d = raw is DateTime dt ? dt.ToString("yyyy-MM-dd") : raw?.ToString() ?? "";
                if (string.IsNullOrEmpty(d)) continue;
                double? v = r.IsDBNull(1) ? null : ParseStr(r.GetValue(1)?.ToString());
                list.Add(new WasteAnalysisResult { 채수일 = d, TOC_TCIC = v });
            }
        }
        catch { }
        return list;
    }

    private static double? ParseStr(string? s)
        => string.IsNullOrWhiteSpace(s) ? null : double.TryParse(s, out var v) ? v : null;
}
