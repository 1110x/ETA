using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteDataService
{
    // ── 업체명으로 전 항목 분석결과 조회 (날짜 오름차순) ─────────────────────
    public static List<WasteAnalysisResult> GetResults(string 업체명)
    {
        var dict = new SortedDictionary<string, WasteAnalysisResult>();
        WasteAnalysisResult Get(string d) =>
            dict.TryGetValue(d, out var r) ? r : (dict[d] = new WasteAnalysisResult { 채수일 = d });

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        Fetch(conn, "SELECT 채수일, 결과 FROM `BOD_DATA`      WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).BOD      = v);
        Fetch(conn, "SELECT 채수일, 결과 FROM `SS_DATA`       WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).SS       = v);
        Fetch(conn, "SELECT 채수일, 농도  FROM `TN_DATA`      WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).TN       = v);
        Fetch(conn, "SELECT 채수일, 농도  FROM `TP_DATA`      WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).TP       = v);
        Fetch(conn, "SELECT 채수일, 검량선_a FROM `TOC_TCIC_DATA` WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).TOC_TCIC = v);
        Fetch(conn, "SELECT 채수일, 검량선_a FROM `TOC_NPOC_DATA` WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).TOC_NPOC = v);
        Fetch(conn, "SELECT 채수일, 농도  FROM `Phenols_DATA` WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).Phenols  = v);
        Fetch(conn, "SELECT 채수일, 결과 FROM `NHexan_DATA`   WHERE 업체명=@n ORDER BY 채수일",
              업체명, (d, v) => Get(d).NHexan   = v);

        return [.. dict.Values];
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
