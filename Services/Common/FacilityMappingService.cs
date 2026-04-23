using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ETA.Services.Common;

/// <summary>
/// 처리시설 매칭 — 외부 시스템(WAYBLE / 측정인.kr) 과 ETA 처리시설_마스터 매핑
/// 테이블: 처리시설_매핑
///   (external_system, external_site_cd, external_sample_cd) UNIQUE
/// </summary>
public static class FacilityMappingService
{
    public record Row(
        int Id,
        string ExternalSystem,   // "WAYBLE" | "측정인"
        string ExternalSiteCd,   // WAYBLE siteCd / 측정인 처리시설코드
        string ExternalSiteNm,   // WAYBLE siteNm / 측정인 처리시설명
        string ExternalSampleCd, // WAYBLE sampleCategory
        string ExternalSampleNm, // WAYBLE sampleCategoryNm
        string EtaSiseol,        // ETA 시설명
        string EtaSiryo,         // ETA 시료명
        string RegisteredAt);

    public static void EnsureTable()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS `처리시설_매핑` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `external_system`   VARCHAR(50)  NOT NULL,
                    `external_site_cd`  VARCHAR(100) NOT NULL DEFAULT '',
                    `external_site_nm`  VARCHAR(200) NOT NULL DEFAULT '',
                    `external_sample_cd` VARCHAR(100) NOT NULL DEFAULT '',
                    `external_sample_nm` VARCHAR(200) NOT NULL DEFAULT '',
                    `시설명`            VARCHAR(200) NOT NULL DEFAULT '',
                    `시료명`            VARCHAR(200) NOT NULL DEFAULT '',
                    `등록일시`          VARCHAR(30),
                    UNIQUE KEY uq_mapping (`external_system`, `external_site_cd`, `external_sample_cd`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.EnsureTable] {ex.Message}"); }
    }

    public static List<Row> GetAllForMaster(string 시설명, string 시료명)
    {
        var list = new List<Row>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT id, external_system, external_site_cd, external_site_nm,
                                       external_sample_cd, external_sample_nm, `시설명`, `시료명`,
                                       COALESCE(등록일시,'')
                                FROM `처리시설_매핑`
                                WHERE `시설명`=@s AND `시료명`=@r
                                ORDER BY external_system, id";
            cmd.Parameters.AddWithValue("@s", 시설명 ?? "");
            cmd.Parameters.AddWithValue("@r", 시료명 ?? "");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new Row(
                    r.GetInt32(0), r.GetString(1), r.GetString(2), r.GetString(3),
                    r.GetString(4), r.GetString(5), r.GetString(6), r.GetString(7),
                    r.GetString(8)));
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.GetAllForMaster] {ex.Message}"); }
        return list;
    }

    /// <summary>UPSERT — (external_system, external_site_cd, external_sample_cd) 기준</summary>
    public static void Save(string system, string siteCd, string siteNm, string sampleCd, string sampleNm,
                            string 시설명, string 시료명)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO `처리시설_매핑`
                (external_system, external_site_cd, external_site_nm,
                 external_sample_cd, external_sample_nm, `시설명`, `시료명`, 등록일시)
                VALUES (@sys, @scd, @snm, @pcd, @pnm, @fac, @smp, @ts)
                ON DUPLICATE KEY UPDATE
                    external_site_nm = VALUES(external_site_nm),
                    external_sample_nm = VALUES(external_sample_nm),
                    `시설명` = VALUES(`시설명`),
                    `시료명` = VALUES(`시료명`),
                    등록일시 = VALUES(등록일시)";
            cmd.Parameters.AddWithValue("@sys", system ?? "");
            cmd.Parameters.AddWithValue("@scd", siteCd ?? "");
            cmd.Parameters.AddWithValue("@snm", siteNm ?? "");
            cmd.Parameters.AddWithValue("@pcd", sampleCd ?? "");
            cmd.Parameters.AddWithValue("@pnm", sampleNm ?? "");
            cmd.Parameters.AddWithValue("@fac", 시설명 ?? "");
            cmd.Parameters.AddWithValue("@smp", 시료명 ?? "");
            cmd.Parameters.AddWithValue("@ts",  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.Save] {ex.Message}"); }
    }

    public static void Delete(int id)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `처리시설_매핑` WHERE id=@id";
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.Delete] {ex.Message}"); }
    }

    /// <summary>모든 매핑 반환 (Show4 카탈로그 뷰용)</summary>
    public static List<Row> GetAllMappings()
    {
        var list = new List<Row>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT id, external_system, external_site_cd, external_site_nm,
                                       external_sample_cd, external_sample_nm, `시설명`, `시료명`,
                                       COALESCE(등록일시,'')
                                FROM `처리시설_매핑`
                                ORDER BY external_system, external_site_cd, external_sample_cd, id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new Row(
                    r.GetInt32(0), r.GetString(1), r.GetString(2), r.GetString(3),
                    r.GetString(4), r.GetString(5), r.GetString(6), r.GetString(7),
                    r.GetString(8)));
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.GetAllMappings] {ex.Message}"); }
        return list;
    }

    /// <summary>처리시설_마스터 에서 모든 (시설명, 시료명) 쌍 반환</summary>
    public static List<(string 시설명, string 시료명)> GetAllMasterPairs()
    {
        var list = new List<(string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `시설명`, `시료명` FROM `처리시설_마스터` ORDER BY `시설명`, id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var f = r.IsDBNull(0) ? "" : r.GetString(0);
                var s = r.IsDBNull(1) ? "" : r.GetString(1);
                if (!string.IsNullOrWhiteSpace(f) || !string.IsNullOrWhiteSpace(s))
                    list.Add((f, s));
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[FacilityMappingService.GetAllMasterPairs] {ex.Message}"); }
        return list;
    }
}
