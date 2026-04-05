using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteSampleService
{
    // ── 날짜 목록 (채수일 역순) ───────────────────────────────────────────────
    public static List<string> GetDates()
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 채수일 FROM `폐수의뢰및결과` ORDER BY 채수일 DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 날짜별 전체 행 (순서 오름차순) ──────────────────────────────────────
    public static List<WasteSample> GetByDate(string 채수일)
    {
        var list = new List<WasteSample>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT Id, 채수일, 구분, 순서, SN, 업체명, 관리번호,
                   BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols,
                   비고, 확인자
            FROM `폐수의뢰및결과`
            WHERE 채수일 = @d
            ORDER BY CASE 구분 WHEN '여수' THEN 0 WHEN '율촌' THEN 1 WHEN '세풍' THEN 2 ELSE 3 END,
                     순서 ASC";
        cmd.Parameters.AddWithValue("@d", 채수일);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(Map(r));
        return list;
    }

    // ── 날짜별 업체명 Set (의뢰 이미 등록된 업체 확인용) ────────────────────
    public static HashSet<string> GetCompanyNamesForDate(string 채수일)
    {
        var set = new HashSet<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 업체명 FROM `폐수의뢰및결과` WHERE 채수일 = @d";
        cmd.Parameters.AddWithValue("@d", 채수일);
        using var r = cmd.ExecuteReader();
        while (r.Read()) set.Add(r.GetString(0));
        return set;
    }

    // ── 행 추가 ──────────────────────────────────────────────────────────────
    public static WasteSample Insert(string 채수일, string 구분, string 업체명, string 관리번호 = "", string 확인자 = "")
    {
        // 해당 날짜+구분의 마지막 순서 +1
        int nextSeq = NextSeq(채수일, 구분);
        string sn   = WasteSample.BuildSN(채수일, 구분, nextSeq);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO `폐수의뢰및결과` (채수일, 구분, 순서, SN, 업체명, 관리번호,
                   BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols, 확인자)
            VALUES (@d, @g, @s, @sn, @name, @no,
                   '', '', '', '', '', '', '', @chk);
            SELECT {DbConnectionFactory.LastInsertId};";
        cmd.Parameters.AddWithValue("@d",    채수일);
        cmd.Parameters.AddWithValue("@g",    구분);
        cmd.Parameters.AddWithValue("@s",    nextSeq);
        cmd.Parameters.AddWithValue("@sn",   sn);
        cmd.Parameters.AddWithValue("@name", 업체명);
        cmd.Parameters.AddWithValue("@no",   관리번호);
        cmd.Parameters.AddWithValue("@chk",  확인자);
        long id = Convert.ToInt64(cmd.ExecuteScalar()!);

        return new WasteSample
        {
            Id = (int)id, 채수일 = 채수일, 구분 = 구분,
            순서 = nextSeq, SN = sn, 업체명 = 업체명, 관리번호 = 관리번호, 확인자 = 확인자
        };
    }

    // ── 행 삭제 후 해당 날짜+구분 순서 재정렬 ────────────────────────────────
    public static void Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 삭제 전 날짜/구분 기억
        string 채수일 = "", 구분 = "";
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT 채수일, 구분 FROM `폐수의뢰및결과` WHERE Id=@id";
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            if (r.Read()) { 채수일 = r.GetString(0); 구분 = r.GetString(1); }
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "DELETE FROM `폐수의뢰및결과` WHERE Id=@id";
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
        }

        if (!string.IsNullOrEmpty(채수일))
            Renumber(conn, 채수일, 구분);
    }

    // ── 순서 변경 (위/아래 이동) ──────────────────────────────────────────────
    public static List<WasteSample> MoveUp(int id, string 채수일, string 구분)
        => Move(id, 채수일, 구분, -1);

    public static List<WasteSample> MoveDown(int id, string 채수일, string 구분)
        => Move(id, 채수일, 구분, +1);

    private static List<WasteSample> Move(int id, string 채수일, string 구분, int dir)
    {
        var rows = GetByDate(채수일);
        var group = rows.FindAll(r => r.구분 == 구분);
        int idx = group.FindIndex(r => r.Id == id);
        int swapIdx = idx + dir;
        if (swapIdx < 0 || swapIdx >= group.Count) return rows;

        // 순서 교환
        (group[idx].순서, group[swapIdx].순서) = (group[swapIdx].순서, group[idx].순서);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        UpdateSeq(conn, group[idx]);
        UpdateSeq(conn, group[swapIdx]);
        Renumber(conn, 채수일, 구분);

        return GetByDate(채수일);
    }

    // ── 드래그앤드롭 순서 변경 (임의 위치로 이동) ────────────────────────────
    public static void ReorderTo(int id, string 채수일, string 구분, int newIndex)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 현재 그룹의 Id 목록 (순서대로)
        var ids = new List<int>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Id FROM `폐수의뢰및결과` WHERE 채수일=@d AND 구분=@g ORDER BY 순서 ASC";
            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@g", 구분);
            using var r = cmd.ExecuteReader();
            while (r.Read()) ids.Add(r.GetInt32(0));
        }

        int oldIndex = ids.IndexOf(id);
        if (oldIndex < 0 || oldIndex == newIndex) return;
        if (newIndex < 0) newIndex = 0;
        if (newIndex >= ids.Count) newIndex = ids.Count - 1;

        // 리스트에서 제거 후 새 위치에 삽입
        ids.RemoveAt(oldIndex);
        ids.Insert(newIndex, id);

        // 전체 재번호
        for (int i = 0; i < ids.Count; i++)
        {
            int seq = i + 1;
            string sn = WasteSample.BuildSN(채수일, 구분, seq);
            using var upd = conn.CreateCommand();
            upd.CommandText = "UPDATE `폐수의뢰및결과` SET 순서=@s, SN=@sn WHERE Id=@id";
            upd.Parameters.AddWithValue("@s",  seq);
            upd.Parameters.AddWithValue("@sn", sn);
            upd.Parameters.AddWithValue("@id", ids[i]);
            upd.ExecuteNonQuery();
        }
    }

    // ── 내부 헬퍼 ─────────────────────────────────────────────────────────────
    public static void UpdateValues(int id, string bod, string toc, string ss,
                                     string tn, string tp, string nHexan, string phenols)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `폐수의뢰및결과`
            SET BOD=@bod, `TOC`=@toc, SS=@ss,
                `T-N`=@tn, `T-P`=@tp, `N-Hexan`=@nh, Phenols=@ph
            WHERE Id=@id";
        cmd.Parameters.AddWithValue("@bod", bod);
        cmd.Parameters.AddWithValue("@toc", toc);
        cmd.Parameters.AddWithValue("@ss",  ss);
        cmd.Parameters.AddWithValue("@tn",  tn);
        cmd.Parameters.AddWithValue("@tp",  tp);
        cmd.Parameters.AddWithValue("@nh",  nHexan);
        cmd.Parameters.AddWithValue("@ph",  phenols);
        cmd.Parameters.AddWithValue("@id",  id);
        cmd.ExecuteNonQuery();
    }

    private static int NextSeq(string 채수일, string 구분)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(MAX(순서),0)+1 FROM `폐수의뢰및결과` WHERE 채수일=@d AND 구분=@g";
        cmd.Parameters.AddWithValue("@d", 채수일);
        cmd.Parameters.AddWithValue("@g", 구분);
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    private static void UpdateSeq(DbConnection conn, WasteSample s)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `폐수의뢰및결과` SET 순서=@seq WHERE Id=@id";
        cmd.Parameters.AddWithValue("@seq", s.순서);
        cmd.Parameters.AddWithValue("@id",  s.Id);
        cmd.ExecuteNonQuery();
    }

    // 순서 재정렬 + SN 재생성
    private static void Renumber(DbConnection conn, string 채수일, string 구분)
    {
        using var sel = conn.CreateCommand();
        sel.CommandText = "SELECT Id FROM `폐수의뢰및결과` WHERE 채수일=@d AND 구분=@g ORDER BY 순서 ASC";
        sel.Parameters.AddWithValue("@d", 채수일);
        sel.Parameters.AddWithValue("@g", 구분);
        var ids = new List<long>();
        using (var r = sel.ExecuteReader())
            while (r.Read()) ids.Add(r.GetInt64(0));

        for (int i = 0; i < ids.Count; i++)
        {
            int seq = i + 1;
            string sn = WasteSample.BuildSN(채수일, 구분, seq);
            using var upd = conn.CreateCommand();
            upd.CommandText = "UPDATE `폐수의뢰및결과` SET 순서=@s, SN=@sn WHERE Id=@id";
            upd.Parameters.AddWithValue("@s",   seq);
            upd.Parameters.AddWithValue("@sn",  sn);
            upd.Parameters.AddWithValue("@id",  ids[i]);
            upd.ExecuteNonQuery();
        }
    }

    private static WasteSample Map(DbDataReader r) => new()
    {
        Id       = r.GetInt32(0),
        채수일   = r.IsDBNull(1) ? "" : r.GetString(1),
        구분     = r.IsDBNull(2) ? "" : r.GetString(2),
        순서     = r.GetInt32(3),
        SN       = r.IsDBNull(4) ? "" : r.GetString(4),
        업체명   = r.IsDBNull(5) ? "" : r.GetString(5),
        관리번호 = r.IsDBNull(6) ? "" : r.GetString(6),
        BOD      = r.IsDBNull(7) ? "" : r.GetString(7),
        TOC= r.IsDBNull(8) ? "" : r.GetString(8),
        SS       = r.IsDBNull(9) ? "" : r.GetString(9),
        TN       = r.IsDBNull(10) ? "" : r.GetString(10),
        TP       = r.IsDBNull(11) ? "" : r.GetString(11),
        NHexan   = r.IsDBNull(12) ? "" : r.GetString(12),
        Phenols  = r.IsDBNull(13) ? "" : r.GetString(13),
        비고     = r.IsDBNull(14) ? "" : r.GetString(14),
        확인자   = r.IsDBNull(15) ? "" : r.GetString(15),
    };

    // ── *_DATA 원시 측정값 조회 ──────────────────────────────────────────────

    /// <summary>*_DATA 테이블에서 채수일+SN 기준 전체 컬럼을 Dictionary로 반환</summary>
    public static Dictionary<string, string>? GetRawData(string tableName, string 채수일, string sn)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 테이블 존재 확인
            if (!DbConnectionFactory.TableExists(conn, tableName))
            {
                System.Diagnostics.Debug.WriteLine($"[GetRawData] 테이블 없음: {tableName}");
                return null;
            }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `{tableName}` WHERE 채수일=@d AND SN=@sn LIMIT 1";
            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@sn", sn);
            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
            {
                System.Diagnostics.Debug.WriteLine($"[GetRawData] 데이터 없음: {tableName} 채수일={채수일} SN={sn}");
                return null;
            }
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var name = reader.GetName(i);
                var val = reader.IsDBNull(i) ? "" : reader.GetValue(i)?.ToString() ?? "";
                dict[name] = val;
            }
            System.Diagnostics.Debug.WriteLine($"[GetRawData] {tableName} SN={sn}: {dict.Count}개 컬럼 로드");
            return dict;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GetRawData] 오류: {tableName} - {ex.Message}");
            return null;
        }
    }

    // ── *_DATA 원시 측정값 UPSERT ─────────────────────────────────────────────

    /// <summary>BOD_DATA에 원시 측정값 UPSERT (채수일+SN 기준)</summary>
    public static void UpsertBodData(
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string d1, string d2, string 희석배수, string 결과,
        string 식종시료량 = "", string 식종D1 = "", string 식종D2 = "",
        string 식종BOD = "", string 식종함유량 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM `BOD_DATA` WHERE 채수일=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d",  채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `BOD_DATA`
                    SET 시료량=@vol, D1=@d1, D2=@d2, 희석배수=@dil, 결과=@r,
                        식종시료량=@sv, 식종D1=@sd1, 식종D2=@sd2, 식종BOD=@sbod, 식종함유량=@spct,
                        등록일시={DbConnectionFactory.NowExpr}
                    WHERE 채수일=@d AND SN=@sn";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `BOD_DATA`
                    (채수일, SN, 업체명, 구분, 시료량, D1, D2, 희석배수, 결과,
                     식종시료량, 식종D1, 식종D2, 식종BOD, 식종함유량, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @vol, @d1, @d2, @dil, @r,
                            @sv, @sd1, @sd2, @sbod, @spct, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }
            cmd.Parameters.AddWithValue("@d",    채수일);
            cmd.Parameters.AddWithValue("@sn",   sn);
            cmd.Parameters.AddWithValue("@vol",  시료량);
            cmd.Parameters.AddWithValue("@d1",   d1);
            cmd.Parameters.AddWithValue("@d2",   d2);
            cmd.Parameters.AddWithValue("@dil",  희석배수);
            cmd.Parameters.AddWithValue("@r",    결과);
            cmd.Parameters.AddWithValue("@sv",   식종시료량);
            cmd.Parameters.AddWithValue("@sd1",  식종D1);
            cmd.Parameters.AddWithValue("@sd2",  식종D2);
            cmd.Parameters.AddWithValue("@sbod", 식종BOD);
            cmd.Parameters.AddWithValue("@spct", 식종함유량);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[UpsertBodData] 오류: {ex.Message}");
        }
    }

    /// <summary>SS_DATA / NHexan_DATA 등 단순 결과값 테이블 UPSERT</summary>
    public static void UpsertSimpleData(string tableName, string resultCol,
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 결과)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE 채수일=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d",  채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 시료량=@vol, `{resultCol}`=@r,
                        등록일시={DbConnectionFactory.NowExpr}
                    WHERE 채수일=@d AND SN=@sn";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (채수일, SN, 업체명, 구분, 시료량, `{resultCol}`, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @vol, @r, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }
            cmd.Parameters.AddWithValue("@d",   채수일);
            cmd.Parameters.AddWithValue("@sn",  sn);
            cmd.Parameters.AddWithValue("@vol", 시료량);
            cmd.Parameters.AddWithValue("@r",   결과);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[UpsertSimpleData:{tableName}] 오류: {ex.Message}");
        }
    }

    /// <summary>SS_DATA 전용 UPSERT (전무게/후무게/무게차/희석배수 포함)</summary>
    public static void UpsertSsData(
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 전무게, string 후무게, string 무게차, string 희석배수, string 결과)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 컬럼 존재 확인 및 추가
            foreach (var col in new[] { "전무게", "후무게", "무게차", "희석배수" })
            {
                if (!DbConnectionFactory.ColumnExists(conn, "SS_DATA", col))
                {
                    using var alt = conn.CreateCommand();
                    alt.CommandText = $"ALTER TABLE `SS_DATA` ADD COLUMN `{col}` TEXT DEFAULT ''";
                    alt.ExecuteNonQuery();
                }
            }

            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM `SS_DATA` WHERE 채수일=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d", 채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `SS_DATA`
                    SET 시료량=@vol, `전무게`=@pre, `후무게`=@post, `무게차`=@diff, `희석배수`=@dil, `결과`=@r,
                        등록일시={DbConnectionFactory.NowExpr}
                    WHERE 채수일=@d AND SN=@sn";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `SS_DATA`
                    (채수일, SN, 업체명, 구분, 시료량, `전무게`, `후무게`, `무게차`, `희석배수`, `결과`, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @vol, @pre, @post, @diff, @dil, @r, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }
            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@sn", sn);
            cmd.Parameters.AddWithValue("@vol", 시료량);
            cmd.Parameters.AddWithValue("@pre", 전무게);
            cmd.Parameters.AddWithValue("@post", 후무게);
            cmd.Parameters.AddWithValue("@diff", 무게차);
            cmd.Parameters.AddWithValue("@dil", 희석배수);
            cmd.Parameters.AddWithValue("@r", 결과);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[UpsertSsData] 오류: {ex.Message}");
        }
    }

    /// <summary>TN_DATA / TP_DATA / Phenols_DATA 등 UV VIS 기반 테이블 UPSERT</summary>
    public static void UpsertUvvisData(string tableName,
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 흡광도, string 희석배수, string 검량선a, string 농도)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE 채수일=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d",  채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 시료량=@vol, 흡광도=@abs, 희석배수=@dil, 검량선_a=@slope, 농도=@r,
                        등록일시={DbConnectionFactory.NowExpr}
                    WHERE 채수일=@d AND SN=@sn";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (채수일, SN, 업체명, 구분, 시료량, 흡광도, 희석배수, 검량선_a, 농도, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @vol, @abs, @dil, @slope, @r, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }
            cmd.Parameters.AddWithValue("@d",     채수일);
            cmd.Parameters.AddWithValue("@sn",    sn);
            cmd.Parameters.AddWithValue("@vol",   시료량);
            cmd.Parameters.AddWithValue("@abs",   흡광도);
            cmd.Parameters.AddWithValue("@dil",   희석배수);
            cmd.Parameters.AddWithValue("@slope", 검량선a);
            cmd.Parameters.AddWithValue("@r",     농도);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[UpsertUvvisData:{tableName}] 오류: {ex.Message}");
        }
    }
}
