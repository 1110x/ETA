using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE4;

namespace ETA.Services.SERVICE2;

public static class WasteSampleService
{
    // UV 5항목 컬럼이 비용부담금_결과 테이블에 존재하는지 여부 캐시
    private static bool _uvColumnsEnsured;
    private static readonly string[] _uvColumnNames = { "시안", "6가크롬", "색도", "ABS", "불소" };

    /// <summary>비용부담금_결과 테이블에 UV 5항목 컬럼이 존재하는지 확인하고, 없으면 추가</summary>
    private static void EnsureUvColumns(DbConnection conn)
    {
        if (_uvColumnsEnsured) return;
        foreach (var col in _uvColumnNames)
        {
            if (!DbConnectionFactory.ColumnExists(conn, "비용부담금_결과", col))
            {
                try
                {
                    using var alt = conn.CreateCommand();
                    alt.CommandText = $"ALTER TABLE `비용부담금_결과` ADD COLUMN `{col}` TEXT DEFAULT ''";
                    alt.ExecuteNonQuery();
                }
                catch { /* 이미 존재하거나 병렬 마이그레이션 */ }
            }
        }
        _uvColumnsEnsured = true;
    }

    // 비용부담금_결과 SELECT의 공통 컬럼 목록
    private const string SelectColumns = @"Id, 채수일, 구분, 순서, SN, 업체명, 관리번호,
                   BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols,
                   비고, 확인자,
                   `시안`, `6가크롬`, `색도`, `ABS`, `불소`";

    // ── 연월 목록 (비용부담금_결과 + 처리시설_작업 UNION, 역순) ────────────────
    public static List<string> GetMonths()
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT DISTINCT ym FROM (
                SELECT SUBSTR(채수일, 1, 7) AS ym FROM `비용부담금_결과`
                UNION
                SELECT SUBSTR(채취일자, 1, 7) AS ym FROM `처리시설_작업`
            ) t ORDER BY ym DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 월별 날짜 목록 (UNION, 역순) ──────────────────────────────────────────
    public static List<string> GetDatesByMonth(string yearMonth)
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT DISTINCT d FROM (
                SELECT 채수일 AS d FROM `비용부담금_결과` WHERE SUBSTR(채수일,1,7)=@ym
                UNION
                SELECT 채취일자 AS d FROM `처리시설_작업` WHERE SUBSTR(채취일자,1,7)=@ym
            ) t ORDER BY d DESC";
        cmd.Parameters.AddWithValue("@ym", yearMonth);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 날짜 목록 (채수일 역순) ───────────────────────────────────────────────
    public static List<string> GetDates()
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 채수일 FROM `비용부담금_결과` ORDER BY 채수일 DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 날짜 목록 (지정일 이하, 채수일 역순) ─────────────────────────────────
    public static List<string> GetDatesUpTo(string maxDateInclusive)
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT DISTINCT 채수일
            FROM `비용부담금_결과`
            WHERE 채수일 <= @maxDate
            ORDER BY 채수일 DESC";
        cmd.Parameters.AddWithValue("@maxDate", maxDateInclusive);
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
        EnsureUvColumns(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT {SelectColumns}
            FROM `비용부담금_결과`
            WHERE 채수일 = @d
            ORDER BY CASE 구분 WHEN '여수' THEN 0 WHEN '율촌' THEN 1 WHEN '세풍' THEN 2 ELSE 3 END,
                     순서 ASC";
        cmd.Parameters.AddWithValue("@d", 채수일);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(Map(r));
        return list;
    }

    // ── 최근 N개월 시료 목록 조회 (수동 매칭 팝업용) ────────────────────────
    public static List<WasteSample> GetRecentSamples(int months = 1)
    {
        var list = new List<WasteSample>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureUvColumns(conn);
            using var cmd = conn.CreateCommand();
            var cutoff = DateTime.Today.AddMonths(-months).ToString("yyyy-MM-dd");
            cmd.CommandText = $@"
                SELECT {SelectColumns}
                FROM `비용부담금_결과`
                WHERE 채수일 >= @cutoff
                ORDER BY 채수일 DESC, 순서 ASC";
            cmd.Parameters.AddWithValue("@cutoff", cutoff);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(Map(r));
        }
        catch (Exception ex)
        {
        }
        return list;
    }

    // ── 날짜별 업체명 Set (의뢰 이미 등록된 업체 확인용) ────────────────────
    public static HashSet<string> GetCompanyNamesForDate(string 채수일)
    {
        var set = new HashSet<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 업체명 FROM `비용부담금_결과` WHERE 채수일 = @d";
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
            INSERT INTO `비용부담금_결과` (채수일, 구분, 순서, SN, 업체명, 관리번호,
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
            cmd.CommandText = "SELECT 채수일, 구분 FROM `비용부담금_결과` WHERE Id=@id";
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            if (r.Read()) { 채수일 = r.GetString(0); 구분 = r.GetString(1); }
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "DELETE FROM `비용부담금_결과` WHERE Id=@id";
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
            cmd.CommandText = "SELECT Id FROM `비용부담금_결과` WHERE 채수일=@d AND 구분=@g ORDER BY 순서 ASC";
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
            upd.CommandText = "UPDATE `비용부담금_결과` SET 순서=@s, SN=@sn WHERE Id=@id";
            upd.Parameters.AddWithValue("@s",  seq);
            upd.Parameters.AddWithValue("@sn", sn);
            upd.Parameters.AddWithValue("@id", ids[i]);
            upd.ExecuteNonQuery();
        }
    }

    // ── 내부 헬퍼 ─────────────────────────────────────────────────────────────
    public static void UpdateValues(int id, string bod, string toc, string ss,
                                     string tn, string tp, string nHexan, string phenols,
                                     string cn = "", string cr6 = "", string color = "",
                                     string abs = "", string fluoride = "")
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureUvColumns(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `비용부담금_결과`
            SET BOD=@bod, `TOC`=@toc, SS=@ss,
                `T-N`=@tn, `T-P`=@tp, `N-Hexan`=@nh, Phenols=@ph,
                `시안`=@cn, `6가크롬`=@cr6, `색도`=@color, `ABS`=@abs, `불소`=@fl
            WHERE Id=@id";
        cmd.Parameters.AddWithValue("@bod", bod);
        cmd.Parameters.AddWithValue("@toc", toc);
        cmd.Parameters.AddWithValue("@ss",  ss);
        cmd.Parameters.AddWithValue("@tn",  tn);
        cmd.Parameters.AddWithValue("@tp",  tp);
        cmd.Parameters.AddWithValue("@nh",  nHexan);
        cmd.Parameters.AddWithValue("@ph",  phenols);
        cmd.Parameters.AddWithValue("@cn",  cn);
        cmd.Parameters.AddWithValue("@cr6", cr6);
        cmd.Parameters.AddWithValue("@color", color);
        cmd.Parameters.AddWithValue("@abs", abs);
        cmd.Parameters.AddWithValue("@fl",  fluoride);
        cmd.Parameters.AddWithValue("@id",  id);
        cmd.ExecuteNonQuery();
    }

    public static int NextSeq(string 채수일, string 구분)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(MAX(순서),0)+1 FROM `비용부담금_결과` WHERE 채수일=@d AND 구분=@g";
        cmd.Parameters.AddWithValue("@d", 채수일);
        cmd.Parameters.AddWithValue("@g", 구분);
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    private static void UpdateSeq(DbConnection conn, WasteSample s)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `비용부담금_결과` SET 순서=@seq WHERE Id=@id";
        cmd.Parameters.AddWithValue("@seq", s.순서);
        cmd.Parameters.AddWithValue("@id",  s.Id);
        cmd.ExecuteNonQuery();
    }

    // 순서 재정렬 + SN 재생성
    private static void Renumber(DbConnection conn, string 채수일, string 구분)
    {
        using var sel = conn.CreateCommand();
        sel.CommandText = "SELECT Id FROM `비용부담금_결과` WHERE 채수일=@d AND 구분=@g ORDER BY 순서 ASC";
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
            upd.CommandText = "UPDATE `비용부담금_결과` SET 순서=@s, SN=@sn WHERE Id=@id";
            upd.Parameters.AddWithValue("@s",   seq);
            upd.Parameters.AddWithValue("@sn",  sn);
            upd.Parameters.AddWithValue("@id",  ids[i]);
            upd.ExecuteNonQuery();
        }
    }

    /// <summary>비용부담금_결과 SELECT 결과를 WasteSample로 매핑 (UV 5항목은 SELECT에 포함된 경우만)</summary>
    private static WasteSample Map(DbDataReader r)
    {
        var s = new WasteSample
        {
            Id       = r.GetInt32(0),
            채수일   = r.IsDBNull(1) ? "" : r.GetString(1),
            구분     = r.IsDBNull(2) ? "" : r.GetString(2),
            순서     = r.GetInt32(3),
            SN       = r.IsDBNull(4) ? "" : r.GetString(4),
            업체명   = r.IsDBNull(5) ? "" : r.GetString(5),
            관리번호 = r.IsDBNull(6) ? "" : r.GetString(6),
            BOD      = r.IsDBNull(7) ? "" : r.GetString(7),
            TOC      = r.IsDBNull(8) ? "" : r.GetString(8),
            SS       = r.IsDBNull(9) ? "" : r.GetString(9),
            TN       = r.IsDBNull(10) ? "" : r.GetString(10),
            TP       = r.IsDBNull(11) ? "" : r.GetString(11),
            NHexan   = r.IsDBNull(12) ? "" : r.GetString(12),
            Phenols  = r.IsDBNull(13) ? "" : r.GetString(13),
            비고     = r.IsDBNull(14) ? "" : r.GetString(14),
            확인자   = r.IsDBNull(15) ? "" : r.GetString(15),
        };
        // UV 5항목 (r.FieldCount > 16 이면 SELECT에 포함됨)
        if (r.FieldCount > 16)
        {
            s.CN       = r.IsDBNull(16) ? "" : r.GetString(16);
            s.CR6      = r.IsDBNull(17) ? "" : r.GetString(17);
            s.COLOR    = r.IsDBNull(18) ? "" : r.GetString(18);
            s.ABS      = r.IsDBNull(19) ? "" : r.GetString(19);
            s.FLUORIDE = r.IsDBNull(20) ? "" : r.GetString(20);
        }
        return s;
    }

    // ── *_DATA 원시 측정값 조회 ──────────────────────────────────────────────

    /// <summary>*_DATA 테이블에서 분석일+SN 기준 전체 컬럼을 Dictionary로 반환</summary>
    public static Dictionary<string, string>? GetRawData(string tableName, string 채수일, string sn)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 테이블 존재 확인
            if (!DbConnectionFactory.TableExists(conn, tableName))
            {
                return null;
            }

            using var cmd = conn.CreateCommand();
            // SN만으로 조회 (분석일 ≠ 채수일 대비), 최신 등록 우선
            cmd.CommandText = $"SELECT * FROM `{tableName}` WHERE SN=@sn ORDER BY Id DESC LIMIT 1";
            cmd.Parameters.AddWithValue("@sn", sn);
            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
            {
                return null;
            }
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var name = reader.GetName(i);
                var val = reader.IsDBNull(i) ? "" : reader.GetValue(i)?.ToString() ?? "";
                dict[name] = val;
            }
            return dict;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    // ── *_DATA 원시 측정값 UPSERT ─────────────────────────────────────────────

    /// <summary>BOD_DATA에 원시 측정값 UPSERT (SN+업체명 기준)</summary>
    public static void UpsertBodData(
        string tableName,
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string d1, string d2, string 희석배수, string 결과,
        string 소스구분 = "폐수배출업소", string 식종시료량 = "", string 식종D1 = "", string 식종D2 = "",
        string 식종BOD = "", string 식종함유량 = "", string 비고 = "", string 시료명 = "",
        string 분석자1 = "", string 분석자2 = "", string minDO = "", string dayDO = "", string 희석수시료량 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE SN=@sn AND 업체명=@nm";
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm", 업체명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 분석일=@d, 시료량=@vol, D1=@d1, D2=@d2, 희석배수=@dil, 결과=@r,
                        식종시료량=@sv, 식종D1=@sd1, 식종D2=@sd2, 식종BOD=@sbod, 식종함유량=@spct,
                        분석자1=@a1, 분석자2=@a2, `15min_DO`=@mdo, `5Day_DO`=@ddo, 희석수시료량=@hsv,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE SN=@sn AND 업체명=@nm";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (분석일, SN, 업체명, 구분, 소스구분, 시료명, 시료량, D1, D2, 희석배수, 결과,
                     식종시료량, 식종D1, 식종D2, 식종BOD, 식종함유량,
                     분석자1, 분석자2, `15min_DO`, `5Day_DO`, 희석수시료량, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @src, @nm2, @vol, @d1, @d2, @dil, @r,
                            @sv, @sd1, @sd2, @sbod, @spct,
                            @a1, @a2, @mdo, @ddo, @hsv, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@gu",  구분);
                cmd.Parameters.AddWithValue("@src", 소스구분);
                cmd.Parameters.AddWithValue("@nm2", 시료명);
            }
            cmd.Parameters.AddWithValue("@d",     채수일);
            cmd.Parameters.AddWithValue("@sn",    sn);
            cmd.Parameters.AddWithValue("@nm",    업체명);
            cmd.Parameters.AddWithValue("@vol",   시료량);
            cmd.Parameters.AddWithValue("@d1",    d1);
            cmd.Parameters.AddWithValue("@d2",    d2);
            cmd.Parameters.AddWithValue("@dil",   희석배수);
            cmd.Parameters.AddWithValue("@r",     결과);
            cmd.Parameters.AddWithValue("@sv",    식종시료량);
            cmd.Parameters.AddWithValue("@sd1",   식종D1);
            cmd.Parameters.AddWithValue("@sd2",   식종D2);
            cmd.Parameters.AddWithValue("@sbod",  식종BOD);
            cmd.Parameters.AddWithValue("@spct",  식종함유량);
            cmd.Parameters.AddWithValue("@a1",    분석자1);
            cmd.Parameters.AddWithValue("@a2",    분석자2);
            cmd.Parameters.AddWithValue("@mdo",   minDO);
            cmd.Parameters.AddWithValue("@ddo",   dayDO);
            cmd.Parameters.AddWithValue("@hsv",   희석수시료량);
            cmd.Parameters.AddWithValue("@remark", 비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[UpsertBodData] {ex.Message}");
        }
    }

    /// <summary>비용부담금_결과 테이블에 결과값 UPDATE (채수일+SN 기준)</summary>
    /// <summary>
    /// 분석일(BOD-DATA의 실제 분석 날짜) + SN으로 비용부담금_결과를 찾아 결과값 UPDATE.
    /// SN 형식 MM-DD-NN 에서 채수월을 추출하고,
    /// 채수일은 분석일 기준 -40일 ~ 분석일 범위 내에서 SN 매칭.
    /// 12월 채수 시료는 다음 해 1월 분석일도 포함.
    /// </summary>
    public static void UpdateWasteResult(string 분석일, string SN, string 시료명,
        string BOD = "", string TOC = "", string SS = "", string TN = "", string TP = "",
        string NHexan = "", string Phenols = "")
    {
        try
        {
            // 분석일에서 연도와 날짜 파싱
            if (!DateTime.TryParse(분석일, out var 분석dt)) return;

            // SN에서 채수월 추출 (MM-DD-NN → parts[0] = MM)
            var snParts = SN.Split('-');
            int 채수월 = (snParts.Length >= 1 && int.TryParse(snParts[0], out var m)) ? m : 0;

            // 채수일 검색 범위: 분석일 기준 -40일 ~ 분석일
            var rangeEnd   = 분석dt.ToString("yyyy-MM-dd");
            var rangeStart = 분석dt.AddDays(-40).ToString("yyyy-MM-dd");

            // 12월 채수의 경우 분석일이 다음 해 1~2월일 수 있으므로 연도 후보 추가
            // → rangeStart/rangeEnd로 이미 커버됨 (날짜 범위로 조회)

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = @"SELECT Id FROM `비용부담금_결과`
                WHERE SN=@sn
                  AND 채수일 BETWEEN @start AND @end
                LIMIT 1";
            chk.Parameters.AddWithValue("@sn",    SN);
            chk.Parameters.AddWithValue("@start", rangeStart);
            chk.Parameters.AddWithValue("@end",   rangeEnd);
            var id = chk.ExecuteScalar();
            if (id == null) return;

            var setCols = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(BOD))     setCols.Add("BOD=@bod");
            if (!string.IsNullOrEmpty(TOC))     setCols.Add("`TOC`=@toc");
            if (!string.IsNullOrEmpty(SS))      setCols.Add("SS=@ss");
            if (!string.IsNullOrEmpty(TN))      setCols.Add("`T-N`=@tn");
            if (!string.IsNullOrEmpty(TP))      setCols.Add("`T-P`=@tp");
            if (!string.IsNullOrEmpty(NHexan))  setCols.Add("`N-Hexan`=@nh");
            if (!string.IsNullOrEmpty(Phenols)) setCols.Add("Phenols=@ph");
            if (setCols.Count == 0) return;

            using var upd = conn.CreateCommand();
            upd.CommandText = $"UPDATE `비용부담금_결과` SET {string.Join(", ", setCols)} WHERE Id=@id";
            upd.Parameters.AddWithValue("@id", id);
            if (!string.IsNullOrEmpty(BOD))     upd.Parameters.AddWithValue("@bod", BOD);
            if (!string.IsNullOrEmpty(TOC))     upd.Parameters.AddWithValue("@toc", TOC);
            if (!string.IsNullOrEmpty(SS))      upd.Parameters.AddWithValue("@ss",  SS);
            if (!string.IsNullOrEmpty(TN))      upd.Parameters.AddWithValue("@tn",  TN);
            if (!string.IsNullOrEmpty(TP))      upd.Parameters.AddWithValue("@tp",  TP);
            if (!string.IsNullOrEmpty(NHexan))  upd.Parameters.AddWithValue("@nh",  NHexan);
            if (!string.IsNullOrEmpty(Phenols)) upd.Parameters.AddWithValue("@ph",  Phenols);
            upd.ExecuteNonQuery();
        }
        catch { }
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
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE LEFT(분석일,10)=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d",  채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 시료량=@vol, `{resultCol}`=@r,
                        등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (분석일, SN, 업체명, 구분, 시료량, `{resultCol}`, 등록일시)
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
        }
    }

    /// <summary>SS_DATA 전용 UPSERT (전무게/후무게/무게차/희석배수 포함)</summary>
    public static void UpsertSsData(
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 전무게, string 후무게, string 무게차, string 희석배수, string 결과,
        string 소스구분 = "폐수배출업소", string 비고 = "", string 시료명 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 컬럼 존재 확인 및 추가
            foreach (var col in new[] { "전무게", "후무게", "무게차", "희석배수", "시료명" })
            {
                if (!DbConnectionFactory.ColumnExists(conn, "SS_시험기록부", col))
                {
                    using var alt = conn.CreateCommand();
                    alt.CommandText = $"ALTER TABLE `SS_시험기록부` ADD COLUMN `{col}` TEXT DEFAULT ''";
                    alt.ExecuteNonQuery();
                }
            }

            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM `SS_시험기록부` WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            chk.Parameters.AddWithValue("@d", 채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm2", 시료명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `SS_시험기록부`
                    SET 시료량=@vol, `전무게`=@pre, `후무게`=@post, `무게차`=@diff, `희석배수`=@dil, `결과`=@r,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `SS_시험기록부`
                    (분석일, SN, 업체명, 구분, 소스구분, 시료명, 시료량, `전무게`, `후무게`, `무게차`, `희석배수`, `결과`, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @src, @nm2, @vol, @pre, @post, @diff, @dil, @r, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
                cmd.Parameters.AddWithValue("@src", 소스구분);
            }
            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@sn", sn);
            cmd.Parameters.AddWithValue("@nm2", 시료명);
            cmd.Parameters.AddWithValue("@vol", 시료량);
            cmd.Parameters.AddWithValue("@pre", 전무게);
            cmd.Parameters.AddWithValue("@post", 후무게);
            cmd.Parameters.AddWithValue("@diff", 무게차);
            cmd.Parameters.AddWithValue("@dil", 희석배수);
            cmd.Parameters.AddWithValue("@r", 결과);
            cmd.Parameters.AddWithValue("@remark", 비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>NHexan_DATA 전용 UPSERT (전무게/후무게/무게차/희석배수 포함)</summary>
    public static void UpsertNHexanData(
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 전무게, string 후무게, string 무게차, string 희석배수, string 결과,
        string 소스구분 = "폐수배출업소", string 비고 = "", string 시료명 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            foreach (var col in new[] { "전무게", "후무게", "무게차", "희석배수", "시료명" })
                if (!DbConnectionFactory.ColumnExists(conn, "NHexan_시험기록부", col))
                {
                    using var alt = conn.CreateCommand();
                    alt.CommandText = $"ALTER TABLE `NHexan_시험기록부` ADD COLUMN `{col}` TEXT DEFAULT ''";
                    alt.ExecuteNonQuery();
                }

            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM `NHexan_시험기록부` WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            chk.Parameters.AddWithValue("@d", 채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm2", 시료명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `NHexan_시험기록부`
                    SET 시료량=@vol, 전무게=@pre, 후무게=@post, 무게차=@diff, 희석배수=@dil, 결과=@r,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `NHexan_시험기록부`
                    (분석일, SN, 업체명, 구분, 소스구분, 시료명, 시료량, 전무게, 후무게, 무게차, 희석배수, 결과, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @src, @nm2, @vol, @pre, @post, @diff, @dil, @r, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
                cmd.Parameters.AddWithValue("@src", 소스구분);
            }
            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@sn", sn);
            cmd.Parameters.AddWithValue("@nm2", 시료명);
            cmd.Parameters.AddWithValue("@vol", 시료량);
            cmd.Parameters.AddWithValue("@pre", 전무게);
            cmd.Parameters.AddWithValue("@post", 후무게);
            cmd.Parameters.AddWithValue("@diff", 무게차);
            cmd.Parameters.AddWithValue("@dil", 희석배수);
            cmd.Parameters.AddWithValue("@r", 결과);
            cmd.Parameters.AddWithValue("@remark", 비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>TN_DATA / TP_DATA / Phenols_DATA 등 UV VIS 기반 테이블 UPSERT</summary>
    public static void UpsertUvvisData(string tableName,
        string 채수일, string sn, string 업체명, string 구분,
        string 시료량, string 흡광도, string 희석배수, string 검량선a, string 농도,
        string st01mgl = "", string st02mgl = "", string st03mgl = "", string st04mgl = "", string st05mgl = "",
        string st01abs = "", string st02abs = "", string st03abs = "", string st04abs = "", string st05abs = "",
        string 기울기 = "", string 절편 = "", string R2 = "",
        string 비고 = "", string 시료명 = "", string 소스구분 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            chk.Parameters.AddWithValue("@d",  채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm2", 시료명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 소스구분=@src, 시료량=@vol, 흡광도=@abs, 희석배수=@dil, 검량선_a=@slope, 농도=@r,
                        ST01_mgL=@s1c, ST02_mgL=@s2c, ST03_mgL=@s3c, ST04_mgL=@s4c, ST05_mgL=@s5c,
                        ST01_abs=@s1a, ST02_abs=@s2a, ST03_abs=@s3a, ST04_abs=@s4a, ST05_abs=@s5a,
                        기울기=@slope2, 절편=@intercept, R2=@r2,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (분석일, SN, 업체명, 구분, 시료명, 소스구분, 시료량, 흡광도, 희석배수, 검량선_a, 농도,
                     ST01_mgL, ST02_mgL, ST03_mgL, ST04_mgL, ST05_mgL,
                     ST01_abs, ST02_abs, ST03_abs, ST04_abs, ST05_abs,
                     기울기, 절편, R2, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @nm2, @src, @vol, @abs, @dil, @slope, @r,
                            @s1c, @s2c, @s3c, @s4c, @s5c,
                            @s1a, @s2a, @s3a, @s4a, @s5a,
                            @slope2, @intercept, @r2, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }
            cmd.Parameters.AddWithValue("@d",         채수일);
            cmd.Parameters.AddWithValue("@sn",        sn);
            cmd.Parameters.AddWithValue("@nm2",       시료명);
            cmd.Parameters.AddWithValue("@src",       소스구분);
            cmd.Parameters.AddWithValue("@vol",       시료량);
            cmd.Parameters.AddWithValue("@abs",       흡광도);
            cmd.Parameters.AddWithValue("@dil",       희석배수);
            cmd.Parameters.AddWithValue("@slope",     검량선a);
            cmd.Parameters.AddWithValue("@r",         농도);
            cmd.Parameters.AddWithValue("@s1c", st01mgl); cmd.Parameters.AddWithValue("@s2c", st02mgl);
            cmd.Parameters.AddWithValue("@s3c", st03mgl); cmd.Parameters.AddWithValue("@s4c", st04mgl);
            cmd.Parameters.AddWithValue("@s5c", st05mgl);
            cmd.Parameters.AddWithValue("@s1a", st01abs); cmd.Parameters.AddWithValue("@s2a", st02abs);
            cmd.Parameters.AddWithValue("@s3a", st03abs); cmd.Parameters.AddWithValue("@s4a", st04abs);
            cmd.Parameters.AddWithValue("@s5a", st05abs);
            cmd.Parameters.AddWithValue("@slope2",    기울기);
            cmd.Parameters.AddWithValue("@intercept", 절편);
            cmd.Parameters.AddWithValue("@r2",        R2);
            cmd.Parameters.AddWithValue("@remark",    비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>TOC UPSERT — method: "NPOC"|"TCIC", 엑셀/기기파일 공용</summary>
    public static void UpsertTocData(
        string method, string 분석일, string sn, string 업체명, string 구분,
        string 흡광도, string 희석배수, string 검량선a, string 측정농도, string 결과,
        string 소스구분 = "폐수배출업소", string 비고 = "", string 시료명 = "")
    {
        string tableName = method.Equals("TCIC", StringComparison.OrdinalIgnoreCase)
            ? "TOC_TCIC_시험기록부" : "TOC_NPOC_시험기록부";
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, tableName, "시료명"))
            {
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `{tableName}` ADD COLUMN `시료명` TEXT DEFAULT ''";
                alt.ExecuteNonQuery();
            }
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            chk.Parameters.AddWithValue("@d",  분석일);
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm2", 시료명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 흡광도=@abs, 희석배수=@dil, 검량선_a=@slope, 농도=@r,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (분석일, SN, 업체명, 구분, 소스구분, 시료명, 흡광도, 희석배수, 검량선_a, 농도, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @src, @nm2, @abs, @dil, @slope, @r, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
                cmd.Parameters.AddWithValue("@src", 소스구분);
            }
            cmd.Parameters.AddWithValue("@d",     분석일);
            cmd.Parameters.AddWithValue("@sn",    sn);
            cmd.Parameters.AddWithValue("@nm2",   시료명);
            cmd.Parameters.AddWithValue("@abs",   흡광도);
            cmd.Parameters.AddWithValue("@dil",   string.IsNullOrEmpty(희석배수) ? "1" : 희석배수);
            cmd.Parameters.AddWithValue("@slope", 검량선a);
            cmd.Parameters.AddWithValue("@r",     string.IsNullOrEmpty(결과) ? 측정농도 : 결과);
            cmd.Parameters.AddWithValue("@remark", 비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>GC/MS 데이터 UPSERT — 수질분석센터_*_DATA 테이블 (VOC 스키마)</summary>
    public static void UpsertGcData(
        string tableName, string 분석일, string sn, string 업체명, string 구분,
        string 농도, string ISTD, ExcelDocInfo? 검량선정보,
        GcCompoundCalInfo? compoundCal = null,
        string 비고 = "", string 시료명 = "", string 흡광도 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, tableName, "시료명"))
            {
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `{tableName}` ADD COLUMN `시료명` TEXT DEFAULT ''";
                alt.ExecuteNonQuery();
            }
            if (!DbConnectionFactory.ColumnExists(conn, tableName, "흡광도"))
            {
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `{tableName}` ADD COLUMN `흡광도` TEXT DEFAULT ''";
                alt.ExecuteNonQuery();
            }
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            chk.Parameters.AddWithValue("@d",  분석일);
            chk.Parameters.AddWithValue("@sn", sn);
            chk.Parameters.AddWithValue("@nm2", 시료명);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            // 검량선 데이터 추출 (지정된 성분 우선, 없으면 첫 번째 화합물)
            var firstCompound = compoundCal ?? 검량선정보?.GcCompoundCals?.FirstOrDefault();
            string[] st_농도s = new string[7], st_값s = new string[7], st_istds = new string[7];

            if (firstCompound != null)
            {
                for (int i = 0; i < 7; i++)
                {
                    st_농도s[i] = i < firstCompound.StdConcs.Length ? firstCompound.StdConcs[i] : "";
                    st_값s[i] = i < firstCompound.StdResps.Length ? firstCompound.StdResps[i] : "";
                    st_istds[i] = i < firstCompound.StdIstdResps.Length ? firstCompound.StdIstdResps[i] : "";
                }
            }

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                cmd.CommandText = $@"UPDATE `{tableName}`
                    SET 농도=@conc, ISTD=@istd, 흡광도=@abs,
                        ST1_농도=@st1c, ST1_값=@st1v, ST1_ISTD=@st1i,
                        ST2_농도=@st2c, ST2_값=@st2v, ST2_ISTD=@st2i,
                        ST3_농도=@st3c, ST3_값=@st3v, ST3_ISTD=@st3i,
                        ST4_농도=@st4c, ST4_값=@st4v, ST4_ISTD=@st4i,
                        ST5_농도=@st5c, ST5_값=@st5v, ST5_ISTD=@st5i,
                        ST6_농도=@st6c, ST6_값=@st6v, ST6_ISTD=@st6i,
                        ST7_농도=@st7c, ST7_값=@st7v, ST7_ISTD=@st7i,
                        기울기=@slope, 절편=@intercept, R값=@r,
                        비고=@remark, 등록일시={DbConnectionFactory.NowExpr}
                    WHERE LEFT(분석일,10)=@d AND SN=@sn AND 시료명=@nm2";
            }
            else
            {
                cmd.CommandText = $@"INSERT INTO `{tableName}`
                    (분석일, SN, 업체명, 구분, 시료명, 농도, ISTD, 흡광도,
                     ST1_농도, ST1_값, ST1_ISTD, ST2_농도, ST2_값, ST2_ISTD,
                     ST3_농도, ST3_값, ST3_ISTD, ST4_농도, ST4_값, ST4_ISTD,
                     ST5_농도, ST5_값, ST5_ISTD, ST6_농도, ST6_값, ST6_ISTD,
                     ST7_농도, ST7_값, ST7_ISTD, 기울기, 절편, R값, 비고, 등록일시)
                    VALUES (@d, @sn, @nm, @gu, @nm2, @conc, @istd, @abs,
                            @st1c, @st1v, @st1i, @st2c, @st2v, @st2i,
                            @st3c, @st3v, @st3i, @st4c, @st4v, @st4i,
                            @st5c, @st5v, @st5i, @st6c, @st6v, @st6i,
                            @st7c, @st7v, @st7i, @slope, @intercept, @r, @remark, {DbConnectionFactory.NowExpr})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }

            cmd.Parameters.AddWithValue("@d", 분석일);
            cmd.Parameters.AddWithValue("@sn", sn);
            cmd.Parameters.AddWithValue("@nm2", 시료명);
            cmd.Parameters.AddWithValue("@conc", 농도);
            cmd.Parameters.AddWithValue("@istd", ISTD);
            cmd.Parameters.AddWithValue("@abs", 흡광도);

            // ST 데이터 파라미터 추가
            for (int i = 0; i < 7; i++)
            {
                cmd.Parameters.AddWithValue($"@st{i+1}c", st_농도s[i]);
                cmd.Parameters.AddWithValue($"@st{i+1}v", st_값s[i]);
                cmd.Parameters.AddWithValue($"@st{i+1}i", st_istds[i]);
            }

            cmd.Parameters.AddWithValue("@slope", firstCompound?.Slope ?? "");
            cmd.Parameters.AddWithValue("@intercept", firstCompound?.Intercept ?? "");
            cmd.Parameters.AddWithValue("@r", firstCompound?.R ?? "");
            cmd.Parameters.AddWithValue("@remark", 비고);

            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>의뢰및결과 테이블의 동적 컬럼 갱신 (다성분 GC/ICP/PFAS용)</summary>
    public static void UpdateDynamicValue(string tableName, int id, string columnName, string value)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, tableName, columnName))
            {
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `{tableName}` ADD COLUMN `{columnName}` TEXT DEFAULT ''";
                alt.ExecuteNonQuery();
            }
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `{tableName}` SET `{columnName}` = @v WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@v", value);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    // ── 생태독성 UPSERT ──────────────────────────────────────────────────────
    public static void UpsertEcotoxData(
        string 채수일, string sn, string 업체명, string 구분, string 시료명, string 소스구분,
        string 시험종, string 시험시간, string 시험시간단위,
        int 대조군_생물수, int 대조군_사망수,
        double[] 농도, int[] 생물수, int[] 사망수,
        EcotoxicityService.EcotoxResult? result = null,
        string 비고 = "",
        string 시험번호 = "", string endpoint = "", string 농도단위 = "",
        EcotoxicityService.EcotoxResult? probitResult = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM `생태독성_시험기록부` WHERE LEFT(분석일,10)=@d AND SN=@sn";
            chk.Parameters.AddWithValue("@d", 채수일);
            chk.Parameters.AddWithValue("@sn", sn);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            int cnt = Math.Min(농도.Length, 8);

            using var cmd = conn.CreateCommand();
            if (exists)
            {
                var setCols = new System.Text.StringBuilder();
                setCols.Append("시험번호=@tno, endpoint=@ep, 농도단위=@cu, ");
                setCols.Append("시료명=@nm2, 소스구분=@src, 시험종=@sp, 시험시간=@dur, 시험시간단위=@duru, ");
                setCols.Append("대조군_생물수=@co, 대조군_사망수=@cm, ");
                for (int i = 1; i <= 8; i++)
                    setCols.Append($"농도_{i}=@c{i}, 생물수_{i}=@o{i}, 사망수_{i}=@m{i}, ");
                setCols.Append("LC50=@lc, LC50_하한=@lcl, LC50_상한=@lcu, TU=@tu, 분석방법=@method, trim_percent=@trim, ");
                setCols.Append("probit_EC50=@plc, probit_하한=@plcl, probit_상한=@plcu, probit_TU=@ptu, probit_method=@pm, ");
                setCols.Append($"결과=@result, 비고=@remark, 등록일시={DbConnectionFactory.NowExpr}");
                cmd.CommandText = $"UPDATE `생태독성_시험기록부` SET {setCols} WHERE LEFT(분석일,10)=@d AND SN=@sn";
            }
            else
            {
                var cols = new System.Text.StringBuilder();
                var vals = new System.Text.StringBuilder();
                cols.Append("분석일, SN, 업체명, 구분, 시험번호, endpoint, 농도단위, 시료명, 소스구분, 시험종, 시험시간, 시험시간단위, 대조군_생물수, 대조군_사망수, ");
                vals.Append("@d, @sn, @nm, @gu, @tno, @ep, @cu, @nm2, @src, @sp, @dur, @duru, @co, @cm, ");
                for (int i = 1; i <= 8; i++)
                {
                    cols.Append($"농도_{i}, 생물수_{i}, 사망수_{i}, ");
                    vals.Append($"@c{i}, @o{i}, @m{i}, ");
                }
                cols.Append("LC50, LC50_하한, LC50_상한, TU, 분석방법, trim_percent, probit_EC50, probit_하한, probit_상한, probit_TU, probit_method, 결과, 비고, 등록일시");
                vals.Append($"@lc, @lcl, @lcu, @tu, @method, @trim, @plc, @plcl, @plcu, @ptu, @pm, @result, @remark, {DbConnectionFactory.NowExpr}");
                cmd.CommandText = $"INSERT INTO `생태독성_시험기록부` ({cols}) VALUES ({vals})";
                cmd.Parameters.AddWithValue("@nm", 업체명);
                cmd.Parameters.AddWithValue("@gu", 구분);
            }

            cmd.Parameters.AddWithValue("@d", 채수일);
            cmd.Parameters.AddWithValue("@sn", sn);
            cmd.Parameters.AddWithValue("@tno", 시험번호);
            cmd.Parameters.AddWithValue("@ep",  endpoint);
            cmd.Parameters.AddWithValue("@cu",  농도단위);
            cmd.Parameters.AddWithValue("@nm2", 시료명);
            cmd.Parameters.AddWithValue("@src", 소스구분);
            cmd.Parameters.AddWithValue("@sp", 시험종);
            cmd.Parameters.AddWithValue("@dur", 시험시간);
            cmd.Parameters.AddWithValue("@duru", 시험시간단위);
            cmd.Parameters.AddWithValue("@co", 대조군_생물수.ToString());
            cmd.Parameters.AddWithValue("@cm", 대조군_사망수.ToString());

            for (int i = 1; i <= 8; i++)
            {
                int idx = i - 1;
                cmd.Parameters.AddWithValue($"@c{i}", idx < cnt ? 농도[idx].ToString("G") : "");
                cmd.Parameters.AddWithValue($"@o{i}", idx < cnt ? 생물수[idx].ToString() : "");
                cmd.Parameters.AddWithValue($"@m{i}", idx < cnt ? 사망수[idx].ToString() : "");
            }

            cmd.Parameters.AddWithValue("@lc",     result?.EC50.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@lcl",    result?.LowerCI.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@lcu",    result?.UpperCI.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@tu",     result?.TU.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@method", result?.Method ?? "");
            cmd.Parameters.AddWithValue("@trim",   result?.TrimPercent.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@plc",    probitResult?.EC50.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@plcl",   probitResult?.LowerCI.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@plcu",   probitResult?.UpperCI.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@ptu",    probitResult?.TU.ToString("G") ?? "");
            cmd.Parameters.AddWithValue("@pm",     probitResult?.Method ?? "");
            cmd.Parameters.AddWithValue("@result", result?.TU.ToString("F1") ?? "");
            cmd.Parameters.AddWithValue("@remark", 비고);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    // ── 헬퍼 메서드 ─────────────────────────────────────────────────────────
    private static string GetString(System.Data.Common.DbDataReader r, int index)
    {
        try { return r.IsDBNull(index) ? "" : r.GetString(index) ?? ""; }
        catch { return ""; }
    }

    /// <summary>SN으로 비용부담금_결과에서 업체명/구분 조회</summary>
    public static WasteSample? FindBySN(string sn)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT SN, 업체명, 구분 FROM `비용부담금_결과` WHERE SN=@sn LIMIT 1";
            cmd.Parameters.AddWithValue("@sn", sn);
            using var r = cmd.ExecuteReader();
            if (r.Read())
                return new WasteSample { SN = GetString(r, 0), 업체명 = GetString(r, 1), 구분 = GetString(r, 2) };
        }
        catch { }
        return null;
    }
}
