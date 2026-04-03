using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.Common;
using ETA.Models;
using System.Diagnostics;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteCompanyService
{

    private static void EnsureAbbrevColumn(DbConnection conn)
    {
        if (!DbConnectionFactory.ColumnExists(conn, "폐수배출업소", "약칭"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE `폐수배출업소` ADD COLUMN `약칭` TEXT DEFAULT ''";
            try { cmd.ExecuteNonQuery(); } catch { }
        }
        if (!DbConnectionFactory.ColumnExists(conn, "폐수배출업소", "비용부담금_업체명"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE `폐수배출업소` ADD COLUMN `비용부담금_업체명` TEXT DEFAULT ''";
            try { cmd.ExecuteNonQuery(); } catch { }
        }
    }

    public static List<WasteCompany> GetAllItems()
    {
        var items = new List<WasteCompany>();

        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath))
        {
            Debug.WriteLine("❌ DB 파일 없음");
            return items;
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);

        // ── 율촌/세풍: 기존 폐수배출업소 테이블 (여수 제외) ──────────────────
        using var cmd1 = conn.CreateCommand();
        cmd1.CommandText = @"
            SELECT
                프로젝트,
                프로젝트명,
                관리번호,
                업체명,
                사업자번호,
                COALESCE(약칭, '') AS 약칭,
                COALESCE(비용부담금_업체명, '') AS 비용부담금_업체명
            FROM `폐수배출업소`
            WHERE `프로젝트명` NOT LIKE '%여수%'
            ORDER BY 관리번호 ASC";
        using (var reader = cmd1.ExecuteReader())
        {
            while (reader.Read())
            {
                items.Add(new WasteCompany
                {
                    프로젝트          = GetStringOrEmpty(reader, "프로젝트"),
                    프로젝트명        = GetStringOrEmpty(reader, "프로젝트명"),
                    관리번호          = GetStringOrEmpty(reader, "관리번호"),
                    업체명            = GetStringOrEmpty(reader, "업체명"),
                    사업자번호        = GetStringOrEmpty(reader, "사업자번호"),
                    약칭              = GetStringOrEmpty(reader, "약칭"),
                    비용부담금_업체명 = GetStringOrEmpty(reader, "비용부담금_업체명"),
                    Original업체명    = GetStringOrEmpty(reader, "업체명"),
                });
            }
        }

        // ── 여수: 여수_폐수배출업소 테이블 ───────────────────────────────────
        if (DbConnectionFactory.TableExists(conn, "여수_폐수배출업소"))
        {
            using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = @"
                SELECT id, 업체명, COALESCE(약칭, '') AS 약칭
                FROM `여수_폐수배출업소`
                ORDER BY id ASC";
            using var r2 = cmd2.ExecuteReader();
            while (r2.Read())
            {
                var name = r2.IsDBNull(1) ? "" : r2.GetString(1).Trim();
                var abbr = r2.IsDBNull(2) ? "" : r2.GetString(2).Trim();
                items.Add(new WasteCompany
                {
                    프로젝트          = "여수",
                    프로젝트명        = "여수",
                    관리번호          = r2.GetValue(0).ToString() ?? "",
                    업체명            = name,
                    사업자번호        = "",
                    약칭              = abbr,
                    비용부담금_업체명 = "",
                    Original업체명    = name,
                });
            }
        }

        Debug.WriteLine($"📊 로드 완료: {items.Count}개 업체");
        return items;
    }

    private static string GetStringOrEmpty(DbDataReader reader, string col)
    {
        try
        {
            int i = reader.GetOrdinal(col);
            return !reader.IsDBNull(i) ? reader.GetString(i) ?? "" : "";
        }
        catch
        {
            return "";
        }
    }

    // ── 비용부담금_업체명 저장 ────────────────────────────────────────────────
    public static void UpdateCostName(string 관리번호, string costName)
    {
        if (string.IsNullOrEmpty(관리번호)) return;
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `폐수배출업소` SET `비용부담금_업체명`=@n WHERE `관리번호`=@id";
        cmd.Parameters.AddWithValue("@n",  costName);
        cmd.Parameters.AddWithValue("@id", 관리번호);
        cmd.ExecuteNonQuery();
    }

    // ── 비용부담금_업체명 전체 조회 (Show4 파란색 표시용) ─────────────────────
    public static HashSet<string> GetAllCostNames()
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return set;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `비용부담금_업체명` FROM `폐수배출업소` WHERE `비용부담금_업체명` IS NOT NULL AND `비용부담금_업체명` != ''";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var val = r.GetString(0);
            foreach (var part in val.Split(','))
                if (!string.IsNullOrWhiteSpace(part)) set.Add(part.Trim());
        }
        return set;
    }

    // 필요하면 이 메서드도 사용하세요 (Page에서 ExecuteDirectUpdate 대신 추천)
    public static bool Update(WasteCompany wasteCompany)
    {
        if (string.IsNullOrEmpty(wasteCompany.Original업체명))
        {
            Debug.WriteLine("❌ Original업체명 없음 → 저장 스킵");
            return false;
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // ── 여수: 여수_폐수배출업소 테이블 ───────────────────────────────────
        if (wasteCompany.프로젝트명 == "여수")
        {
            if (!DbConnectionFactory.TableExists(conn, "여수_폐수배출업소"))
            {
                Debug.WriteLine("❌ 여수_폐수배출업소 테이블 없음");
                return false;
            }
            EnsureYeosuAbbrevColumn(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                UPDATE `여수_폐수배출업소`
                SET 업체명 = @업체명, 약칭 = @약칭
                WHERE 업체명 = @Original업체명";
            cmd.Parameters.AddWithValue("@업체명",        wasteCompany.업체명 ?? "");
            cmd.Parameters.AddWithValue("@약칭",          wasteCompany.약칭 ?? "");
            cmd.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);
            int rows = cmd.ExecuteNonQuery();
            Debug.WriteLine($"[UPDATE 여수] {rows}행 업데이트 (업체명: {wasteCompany.Original업체명})");
            if (rows > 0)
                wasteCompany.Original업체명 = wasteCompany.업체명 ?? "";
            return rows > 0;
        }

        // ── 율촌/세풍: 폐수배출업소 테이블 ───────────────────────────────────
        using var cmd2 = conn.CreateCommand();
        EnsureAbbrevColumn(conn);

        cmd2.CommandText = @"
            UPDATE `폐수배출업소`
            SET
                업체명       = @업체명,
                프로젝트     = @프로젝트,
                프로젝트명   = @프로젝트명,
                사업자번호   = @사업자번호,
                약칭         = @약칭
            WHERE 업체명 = @Original업체명";

        cmd2.Parameters.AddWithValue("@업체명",       wasteCompany.업체명 ?? "");
        cmd2.Parameters.AddWithValue("@프로젝트",     wasteCompany.프로젝트 ?? "");
        cmd2.Parameters.AddWithValue("@프로젝트명",   wasteCompany.프로젝트명 ?? "");
        // 사업자번호 컬럼이 INTEGER 타입인 경우 빈 문자열 거부 → NULL로 전달
        var 사업자번호Val = string.IsNullOrEmpty(wasteCompany.사업자번호)
            ? (object)DBNull.Value
            : wasteCompany.사업자번호;
        cmd2.Parameters.AddWithValue("@사업자번호", 사업자번호Val);
        cmd2.Parameters.AddWithValue("@약칭",         wasteCompany.약칭 ?? "");
        cmd2.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);

        int rows2 = cmd2.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE] {rows2}행 업데이트 (업체명: {wasteCompany.Original업체명})");

        if (rows2 > 0)
            wasteCompany.Original업체명 = wasteCompany.업체명 ?? "";  // 키 갱신

        return rows2 > 0;
    }

    private static void EnsureYeosuAbbrevColumn(DbConnection conn)
    {
        if (!DbConnectionFactory.ColumnExists(conn, "여수_폐수배출업소", "약칭"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE `여수_폐수배출업소` ADD COLUMN `약칭` TEXT DEFAULT ''";
            try { cmd.ExecuteNonQuery(); } catch { }
        }
    }
}