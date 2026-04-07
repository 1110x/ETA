using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class FacilityResultService
{
    // ── 메모리 캐시 ──────────────────────────────────────────────────────────
    private static List<string>? _facilityNamesCache;
    private static List<(string 시설명, string 시료명, int 마스터Id)>? _masterSamplesCache;
    public static void InvalidateCache() { _facilityNamesCache = null; _masterSamplesCache = null; }

    // ── 시설명 목록 (처리시설_마스터에서 DISTINCT) ─────────────────────────
    public static List<string> GetFacilityNames()
    {
        if (_facilityNamesCache != null) return _facilityNamesCache;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 시설명 FROM `처리시설_마스터` ORDER BY id";
        var list = new List<string>();
        using var reader = cmd.ExecuteReader();
        while (reader.Read()) list.Add(reader.GetString(0));
        _facilityNamesCache = list;
        return list;
    }

    // ── 마스터 + 측정결과 JOIN 조회 ───────────────────────────────────────
    public static List<FacilityResultRow> GetRows(string facility, string date)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT
                m.id        AS 마스터id,
                m.시료명,
                m.비고      AS 비고마스터,
                m.BOD       AS BOD활성,
                m.TOC       AS TOC활성,
                m.SS        AS SS활성,
                m.`T-N`     AS TN활성,
                m.`T-P`     AS TP활성,
                m.총대장균군 AS 총대장균군활성,
                m.COD       AS COD활성,
                m.염소이온   AS 염소이온활성,
                m.영양염류   AS 영양염류활성,
                m.함수율    AS 함수율활성,
                m.중금속    AS 중금속활성,
                r.id,
                r.BOD       AS BOD값,
                r.TOC       AS TOC값,
                r.SS        AS SS값,
                r.`T-N`     AS TN값,
                r.`T-P`     AS TP값,
                r.총대장균군 AS 총대장균군값,
                r.COD       AS COD값,
                r.염소이온   AS 염소이온값,
                r.영양염류   AS 영양염류값,
                r.함수율    AS 함수율값,
                r.중금속    AS 중금속값,
                r.비고      AS 비고값
            FROM `처리시설_마스터` m
            LEFT JOIN `처리시설_측정결과` r
                ON r.마스터_id = m.id AND r.채취일자 = @date
            WHERE m.시설명 = @facility
            ORDER BY m.id";
        cmd.Parameters.AddWithValue("@facility", facility);
        cmd.Parameters.AddWithValue("@date", date);

        var rows = new List<FacilityResultRow>();
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            static bool Active(object v) =>
                v is not DBNull && v?.ToString()?.StartsWith("O", StringComparison.OrdinalIgnoreCase) == true;
            static string Val(object v) =>
                v is DBNull ? "" : v?.ToString() ?? "";

            rows.Add(new FacilityResultRow
            {
                Id           = reader["id"] is DBNull ? 0 : Convert.ToInt32(reader["id"]),
                마스터Id      = Convert.ToInt32(reader["마스터id"]),
                시료명        = Val(reader["시료명"]),
                비고마스터     = Val(reader["비고마스터"]),
                BOD활성       = Active(reader["BOD활성"]),
                TOC활성       = Active(reader["TOC활성"]),
                SS활성        = Active(reader["SS활성"]),
                TN활성        = Active(reader["TN활성"]),
                TP활성        = Active(reader["TP활성"]),
                총대장균군활성  = Active(reader["총대장균군활성"]),
                COD활성       = Active(reader["COD활성"]),
                염소이온활성   = Active(reader["염소이온활성"]),
                영양염류활성   = Active(reader["영양염류활성"]),
                함수율활성    = Active(reader["함수율활성"]),
                중금속활성    = Active(reader["중금속활성"]),
                BOD          = Val(reader["BOD값"]),
                TOC          = Val(reader["TOC값"]),
                SS           = Val(reader["SS값"]),
                TN           = Val(reader["TN값"]),
                TP           = Val(reader["TP값"]),
                총대장균군    = Val(reader["총대장균군값"]),
                COD          = Val(reader["COD값"]),
                염소이온      = Val(reader["염소이온값"]),
                영양염류      = Val(reader["영양염류값"]),
                함수율       = Val(reader["함수율값"]),
                중금속       = Val(reader["중금속값"]),
                비고         = Val(reader["비고값"]),
            });
        }
        return rows;
    }

    // ── 저장 (INSERT or UPDATE) ───────────────────────────────────────────
    public static void SaveRows(string facility, string date, List<FacilityResultRow> rows, string inputUser)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        foreach (var r in rows)
        {
            if (r.Id == 0)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    INSERT INTO `처리시설_측정결과`
                        (마스터_id, 시설명, 시료명, 채취일자,
                         BOD, TOC, SS, `T-N`, `T-P`,
                         총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속,
                         비고, 입력자, 입력일시)
                    VALUES
                        (@mid, @시설명, @시료명, @date,
                         @BOD, @TOC, @SS, @TN, @TP,
                         @총대장균군, @COD, @염소이온, @영양염류, @함수율, @중금속,
                         @비고, @user, @now)";
                cmd.Parameters.AddWithValue("@mid",    r.마스터Id);
                cmd.Parameters.AddWithValue("@시설명",  facility);
                cmd.Parameters.AddWithValue("@시료명",  r.시료명);
                cmd.Parameters.AddWithValue("@date",   date);
                cmd.Parameters.AddWithValue("@BOD",    r.BOD);
                cmd.Parameters.AddWithValue("@TOC",    r.TOC);
                cmd.Parameters.AddWithValue("@SS",     r.SS);
                cmd.Parameters.AddWithValue("@TN",     r.TN);
                cmd.Parameters.AddWithValue("@TP",     r.TP);
                cmd.Parameters.AddWithValue("@총대장균군", r.총대장균군);
                cmd.Parameters.AddWithValue("@COD",    r.COD);
                cmd.Parameters.AddWithValue("@염소이온", r.염소이온);
                cmd.Parameters.AddWithValue("@영양염류", r.영양염류);
                cmd.Parameters.AddWithValue("@함수율",  r.함수율);
                cmd.Parameters.AddWithValue("@중금속",  r.중금속);
                cmd.Parameters.AddWithValue("@비고",   r.비고);
                cmd.Parameters.AddWithValue("@user",   inputUser);
                cmd.Parameters.AddWithValue("@now",    now);
                cmd.ExecuteNonQuery();
            }
            else
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    UPDATE `처리시설_측정결과` SET
                        BOD=@BOD, TOC=@TOC, SS=@SS, `T-N`=@TN, `T-P`=@TP,
                        총대장균군=@총대장균군, COD=@COD, 염소이온=@염소이온,
                        영양염류=@영양염류, 함수율=@함수율, 중금속=@중금속,
                        비고=@비고, 입력일시=@now
                    WHERE id=@id";
                cmd.Parameters.AddWithValue("@id",     r.Id);
                cmd.Parameters.AddWithValue("@BOD",    r.BOD);
                cmd.Parameters.AddWithValue("@TOC",    r.TOC);
                cmd.Parameters.AddWithValue("@SS",     r.SS);
                cmd.Parameters.AddWithValue("@TN",     r.TN);
                cmd.Parameters.AddWithValue("@TP",     r.TP);
                cmd.Parameters.AddWithValue("@총대장균군", r.총대장균군);
                cmd.Parameters.AddWithValue("@COD",    r.COD);
                cmd.Parameters.AddWithValue("@염소이온", r.염소이온);
                cmd.Parameters.AddWithValue("@영양염류", r.영양염류);
                cmd.Parameters.AddWithValue("@함수율",  r.함수율);
                cmd.Parameters.AddWithValue("@중금속",  r.중금속);
                cmd.Parameters.AddWithValue("@비고",   r.비고);
                cmd.Parameters.AddWithValue("@now",    now);
                cmd.ExecuteNonQuery();
            }
        }
    }

    // ── 해당 일자에 처리시설 측정결과가 하나라도 있는지 ─────────────────────
    public static bool HasResultsForDate(string date)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM `처리시설_측정결과` WHERE 채취일자=@d";
        cmd.Parameters.AddWithValue("@d", date);
        var val = cmd.ExecuteScalar();
        return Convert.ToInt64(val ?? 0L) > 0;
    }

    // ── 해당 일자 처리시설 항목별 입력 현황 (모든 시설 합산) ─────────────────
    public static (bool BOD, bool TOC, bool SS, bool TN, bool TP, bool 대장균)
        GetFillStatusForDate(string date)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        // 활성화된 항목 중 값이 존재하는 행이 하나라도 있으면 true
        cmd.CommandText = @"
            SELECT
                SUM(CASE WHEN m.BOD   LIKE 'O%' AND COALESCE(r.BOD,  '') <> '' THEN 1 ELSE 0 END),
                SUM(CASE WHEN m.TOC   LIKE 'O%' AND COALESCE(r.TOC,  '') <> '' THEN 1 ELSE 0 END),
                SUM(CASE WHEN m.SS    LIKE 'O%' AND COALESCE(r.SS,   '') <> '' THEN 1 ELSE 0 END),
                SUM(CASE WHEN m.`T-N` LIKE 'O%' AND COALESCE(r.`T-N`,'') <> '' THEN 1 ELSE 0 END),
                SUM(CASE WHEN m.`T-P` LIKE 'O%' AND COALESCE(r.`T-P`,'') <> '' THEN 1 ELSE 0 END),
                SUM(CASE WHEN m.총대장균군 LIKE 'O%' AND COALESCE(r.총대장균군,'') <> '' THEN 1 ELSE 0 END)
            FROM `처리시설_마스터` m
            LEFT JOIN `처리시설_측정결과` r ON r.마스터_id=m.id AND r.채취일자=@d";
        cmd.Parameters.AddWithValue("@d", date);
        using var reader = cmd.ExecuteReader();
        if (!reader.Read()) return default;
        static bool Pos(object v) => v is not DBNull && Convert.ToInt64(v) > 0;
        return (Pos(reader[0]), Pos(reader[1]), Pos(reader[2]),
                Pos(reader[3]), Pos(reader[4]), Pos(reader[5]));
    }

    // ── 전체 시설의 시료명 목록 (분류용) ─────────────────────────────────
    public static List<(string 시설명, string 시료명, int 마스터Id)> GetAllMasterSamples()
    {
        if (_masterSamplesCache != null) return _masterSamplesCache;
        var list = new List<(string, string, int)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT id, 시설명, 시료명 FROM `처리시설_마스터` ORDER BY id";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(1), r.GetString(2), Convert.ToInt32(r.GetValue(0))));
        _masterSamplesCache = list;
        return list;
    }

    // ── 마스터 비고 딕셔너리 (masterId → 비고) ───────────────────────────
    public static Dictionary<int, string> GetMasterNotes()
    {
        var dict = new Dictionary<int, string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT id, COALESCE(비고,'') FROM `처리시설_마스터` ORDER BY id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dict[Convert.ToInt32(r.GetValue(0))] = r.GetString(1);
        }
        catch { }
        return dict;
    }

    // ── 시료명으로 시설 검색 (부분일치) ──────────────────────────────────
    public static (string 시설명, int 마스터Id)? FindBySampleName(
        List<(string 시설명, string 시료명, int 마스터Id)> masters, string sampleName)
    {
        var exact = masters.FirstOrDefault(m => m.시료명 == sampleName);
        if (exact != default) return (exact.시설명, exact.마스터Id);
        var partial = masters.FirstOrDefault(m =>
            sampleName.Contains(m.시료명) || m.시료명.Contains(sampleName));
        if (partial != default) return (partial.시설명, partial.마스터Id);
        return null;
    }
}
