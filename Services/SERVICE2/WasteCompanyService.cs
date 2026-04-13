using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteCompanyService
{

    private static readonly (string col, string def)[] _extraCols =
    [
        ("약칭",           "TEXT DEFAULT ''"),
        ("비용부담금_업체명", "TEXT DEFAULT ''"),
        ("BOD",           "TEXT DEFAULT ''"),
        ("TOC",           "TEXT DEFAULT ''"),
        ("SS",            "TEXT DEFAULT ''"),
        ("`T-N`",         "TEXT DEFAULT ''"),
        ("`T-P`",         "TEXT DEFAULT ''"),
        ("Phenols",       "TEXT DEFAULT ''"),
        ("`N-Hexan`",     "TEXT DEFAULT ''"),
        ("승인유량",       "TEXT DEFAULT ''"),
        ("기타특이사항",   "TEXT DEFAULT ''"),
    ];

    private static void EnsureAbbrevColumn(DbConnection conn)
    {
        // 관리번호 컬럼이 INTEGER면 TEXT로 변환 (문자열 관리번호 허용)
        try
        {
            using var typeCmd = conn.CreateCommand();
            typeCmd.CommandText = "ALTER TABLE `폐수배출업소` MODIFY COLUMN `관리번호` TEXT";
            typeCmd.ExecuteNonQuery();
        }
        catch { }

        foreach (var (col, def) in _extraCols)
        {
            var bare = col.Trim('`');
            if (!DbConnectionFactory.ColumnExists(conn, "폐수배출업소", bare))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE `폐수배출업소` ADD COLUMN {col} {def}";
                try { cmd.ExecuteNonQuery(); } catch { }
            }
        }
    }

    // ── 메모리 캐시 ──────────────────────────────────────────────────────────
    private static List<WasteCompany>? _itemsCache;
    public static void InvalidateCache() => _itemsCache = null;

    public static List<WasteCompany> GetAllItems()
    {
        if (_itemsCache != null) return _itemsCache;

        var items = new List<WasteCompany>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);

        // ── 율촌/세풍: 기존 폐수배출업소 테이블 (여수 제외) ──────────────────
        using var cmd1 = conn.CreateCommand();
        cmd1.CommandText = @"
            SELECT
                프로젝트, 프로젝트명, 관리번호, 업체명, 사업자번호,
                COALESCE(약칭, '') AS 약칭,
                COALESCE(비용부담금_업체명, '') AS 비용부담금_업체명,
                COALESCE(BOD, '') AS BOD, COALESCE(TOC, '') AS TOC,
                COALESCE(SS, '') AS SS, COALESCE(`T-N`, '') AS TN,
                COALESCE(`T-P`, '') AS TP, COALESCE(Phenols, '') AS Phenols,
                COALESCE(`N-Hexan`, '') AS NHexan,
                COALESCE(승인유량, '') AS 승인유량,
                COALESCE(기타특이사항, '') AS 기타특이사항
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
                    BOD               = GetStringOrEmpty(reader, "BOD"),
                    TOC               = GetStringOrEmpty(reader, "TOC"),
                    SS                = GetStringOrEmpty(reader, "SS"),
                    TN                = GetStringOrEmpty(reader, "TN"),
                    TP                = GetStringOrEmpty(reader, "TP"),
                    Phenols           = GetStringOrEmpty(reader, "Phenols"),
                    NHexan            = GetStringOrEmpty(reader, "NHexan"),
                    승인유량          = GetStringOrEmpty(reader, "승인유량"),
                    기타특이사항      = GetStringOrEmpty(reader, "기타특이사항"),
                    Original업체명    = GetStringOrEmpty(reader, "업체명"),
                });
            }
        }

        // ── 여수: 여수_폐수배출업소 테이블 ───────────────────────────────────
        if (DbConnectionFactory.TableExists(conn, "여수_폐수배출업소"))
        {
            EnsureYeosuExtraColumns(conn);
            using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = @"
                SELECT id, 업체명, COALESCE(약칭, '') AS 약칭,
                    COALESCE(BOD, '') AS BOD, COALESCE(TOC, '') AS TOC,
                    COALESCE(SS, '') AS SS, COALESCE(`T-N`, '') AS TN,
                    COALESCE(`T-P`, '') AS TP, COALESCE(Phenols, '') AS Phenols,
                    COALESCE(`N-Hexan`, '') AS NHexan,
                    COALESCE(승인유량, '') AS 승인유량,
                    COALESCE(기타특이사항, '') AS 기타특이사항
                FROM `여수_폐수배출업소`
                ORDER BY id ASC";
            using var r2 = cmd2.ExecuteReader();
            while (r2.Read())
            {
                var name = r2.IsDBNull(1) ? "" : r2.GetString(1).Trim();
                items.Add(new WasteCompany
                {
                    프로젝트          = "여수",
                    프로젝트명        = "여수",
                    관리번호          = r2.GetValue(0).ToString() ?? "",
                    업체명            = name,
                    사업자번호        = "",
                    약칭              = GetStringOrEmpty(r2, "약칭"),
                    비용부담금_업체명 = "",
                    BOD               = GetStringOrEmpty(r2, "BOD"),
                    TOC               = GetStringOrEmpty(r2, "TOC"),
                    SS                = GetStringOrEmpty(r2, "SS"),
                    TN                = GetStringOrEmpty(r2, "TN"),
                    TP                = GetStringOrEmpty(r2, "TP"),
                    Phenols           = GetStringOrEmpty(r2, "Phenols"),
                    NHexan            = GetStringOrEmpty(r2, "NHexan"),
                    승인유량          = GetStringOrEmpty(r2, "승인유량"),
                    기타특이사항      = GetStringOrEmpty(r2, "기타특이사항"),
                    Original업체명    = name,
                });
            }
        }

        _itemsCache = items;
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
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `폐수배출업소` SET `비용부담금_업체명`=@n WHERE `관리번호`=@id";
        cmd.Parameters.AddWithValue("@n",  costName);
        cmd.Parameters.AddWithValue("@id", 관리번호);
        cmd.ExecuteNonQuery();
        InvalidateCache();
    }

    // ── 비용부담금_업체명 전체 조회 (Show4 파란색 표시용) ─────────────────────
    public static HashSet<string> GetAllCostNames()
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
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
            return false;
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // ── 여수: 여수_폐수배출업소 테이블 ───────────────────────────────────
        if (wasteCompany.프로젝트명 == "여수")
        {
            if (!DbConnectionFactory.TableExists(conn, "여수_폐수배출업소"))
            {
                return false;
            }
            EnsureYeosuAbbrevColumn(conn);
            EnsureYeosuExtraColumns(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                UPDATE `여수_폐수배출업소`
                SET 업체명=@업체명, 약칭=@약칭,
                    BOD=@BOD, TOC=@TOC, SS=@SS, `T-N`=@TN, `T-P`=@TP,
                    Phenols=@Phenols, `N-Hexan`=@NHexan,
                    승인유량=@승인유량, 기타특이사항=@기타특이사항
                WHERE 업체명 = @Original업체명";
            cmd.Parameters.AddWithValue("@업체명", wasteCompany.업체명 ?? "");
            BindExtraParams(cmd, wasteCompany);
            cmd.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);
            int rows = cmd.ExecuteNonQuery();
            if (rows > 0)
                wasteCompany.Original업체명 = wasteCompany.업체명 ?? "";
            return rows > 0;
        }

        // ── 율촌/세풍: 폐수배출업소 테이블 ───────────────────────────────────
        EnsureAbbrevColumn(conn);
        using var cmd2 = conn.CreateCommand();

        cmd2.CommandText = @"
            UPDATE `폐수배출업소`
            SET 업체명=@업체명, 프로젝트=@프로젝트, 프로젝트명=@프로젝트명,
                사업자번호=@사업자번호, 약칭=@약칭,
                BOD=@BOD, TOC=@TOC, SS=@SS, `T-N`=@TN, `T-P`=@TP,
                Phenols=@Phenols, `N-Hexan`=@NHexan,
                승인유량=@승인유량, 기타특이사항=@기타특이사항
            WHERE 업체명 = @Original업체명";

        cmd2.Parameters.AddWithValue("@업체명",       wasteCompany.업체명 ?? "");
        cmd2.Parameters.AddWithValue("@프로젝트",     wasteCompany.프로젝트 ?? "");
        cmd2.Parameters.AddWithValue("@프로젝트명",   wasteCompany.프로젝트명 ?? "");
        var 사업자번호Val = string.IsNullOrEmpty(wasteCompany.사업자번호)
            ? (object)DBNull.Value : wasteCompany.사업자번호;
        cmd2.Parameters.AddWithValue("@사업자번호", 사업자번호Val);
        BindExtraParams(cmd2, wasteCompany);
        cmd2.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);

        int rows2 = cmd2.ExecuteNonQuery();

        if (rows2 > 0)
        {
            wasteCompany.Original업체명 = wasteCompany.업체명 ?? "";
            InvalidateCache();
        }
        return rows2 > 0;
    }

    // ── 신규 업소 등록 ────────────────────────────────────────────────────────
    public static bool Insert(WasteCompany c)
    {
        if (string.IsNullOrWhiteSpace(c.업체명)) return false;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 여수
        if (c.프로젝트명 == "여수")
        {
            if (!DbConnectionFactory.TableExists(conn, "여수_폐수배출업소")) return false;
            EnsureYeosuAbbrevColumn(conn);
            EnsureYeosuExtraColumns(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO `여수_폐수배출업소`
                    (업체명, 약칭, BOD, TOC, SS, `T-N`, `T-P`, Phenols, `N-Hexan`, 승인유량, 기타특이사항)
                VALUES (@업체명, @약칭, @BOD, @TOC, @SS, @TN, @TP, @Phenols, @NHexan, @승인유량, @기타특이사항)";
            cmd.Parameters.AddWithValue("@업체명", c.업체명 ?? "");
            BindExtraParams(cmd, c);
            int r = cmd.ExecuteNonQuery();
            return r > 0;
        }

        // 율촌/세풍
        EnsureAbbrevColumn(conn);
        using var cmd2 = conn.CreateCommand();
        cmd2.CommandText = @"
            INSERT INTO `폐수배출업소`
                (관리번호, 업체명, 프로젝트, 프로젝트명, 사업자번호, 약칭,
                 BOD, TOC, SS, `T-N`, `T-P`, Phenols, `N-Hexan`, 승인유량, 기타특이사항)
            VALUES (@관리번호, @업체명, @프로젝트, @프로젝트명, @사업자번호, @약칭,
                    @BOD, @TOC, @SS, @TN, @TP, @Phenols, @NHexan, @승인유량, @기타특이사항)";
        cmd2.Parameters.AddWithValue("@관리번호",   c.관리번호 ?? "");
        cmd2.Parameters.AddWithValue("@업체명",     c.업체명);
        cmd2.Parameters.AddWithValue("@프로젝트",   c.프로젝트 ?? "");
        cmd2.Parameters.AddWithValue("@프로젝트명", c.프로젝트명 ?? "");
        var 사번 = string.IsNullOrEmpty(c.사업자번호) ? (object)DBNull.Value : c.사업자번호;
        cmd2.Parameters.AddWithValue("@사업자번호", 사번);
        BindExtraParams(cmd2, c);
        int r2 = cmd2.ExecuteNonQuery();
        return r2 > 0;
    }

    private static void BindExtraParams(DbCommand cmd, WasteCompany c)
    {
        cmd.Parameters.AddWithValue("@약칭",       c.약칭 ?? "");
        cmd.Parameters.AddWithValue("@BOD",        c.BOD ?? "");
        cmd.Parameters.AddWithValue("@TOC",        c.TOC ?? "");
        cmd.Parameters.AddWithValue("@SS",         c.SS ?? "");
        cmd.Parameters.AddWithValue("@TN",         c.TN ?? "");
        cmd.Parameters.AddWithValue("@TP",         c.TP ?? "");
        cmd.Parameters.AddWithValue("@Phenols",    c.Phenols ?? "");
        cmd.Parameters.AddWithValue("@NHexan",     c.NHexan ?? "");
        cmd.Parameters.AddWithValue("@승인유량",   c.승인유량 ?? "");
        cmd.Parameters.AddWithValue("@기타특이사항", c.기타특이사항 ?? "");
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

    private static readonly (string col, string def)[] _yeosuExtraCols =
    [
        ("BOD",         "TEXT DEFAULT ''"),
        ("TOC",         "TEXT DEFAULT ''"),
        ("SS",          "TEXT DEFAULT ''"),
        ("`T-N`",       "TEXT DEFAULT ''"),
        ("`T-P`",       "TEXT DEFAULT ''"),
        ("Phenols",     "TEXT DEFAULT ''"),
        ("`N-Hexan`",   "TEXT DEFAULT ''"),
        ("승인유량",     "TEXT DEFAULT ''"),
        ("기타특이사항", "TEXT DEFAULT ''"),
    ];

    private static void EnsureYeosuExtraColumns(DbConnection conn)
    {
        foreach (var (col, def) in _yeosuExtraCols)
        {
            var bare = col.Trim('`');
            if (!DbConnectionFactory.ColumnExists(conn, "여수_폐수배출업소", bare))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE `여수_폐수배출업소` ADD COLUMN {col} {def}";
                try { cmd.ExecuteNonQuery(); } catch { }
            }
        }
    }
}