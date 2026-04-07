using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteRequestService
{
    // =========================================================================
    // 의뢰번호 생성
    //   여수: MM-dd-순번(D2)          예) 03-21-01
    //   율촌: [율촌]MM-dd-관리번호    예) [율촌]03-21-10
    //   세풍: [세풍]MM-dd-순번(D2)
    // =========================================================================
    public static string GenerateRequestNumber(DateTime date, string 구분, string 관리번호)
    {
        string mm = date.Month.ToString("D2");
        string dd = date.Day.ToString("D2");

        if (구분 == "율촌")
        {
            int.TryParse(관리번호, out int num);
            return $"[율촌]{mm}-{dd}-{num:D2}";
        }

        string prefix = 구분 == "세풍" ? "[세풍]" : "";
        string dateStr = date.ToString("yyyy-MM-dd");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        int seq = GetNextSequence(conn, 구분, dateStr);
        return $"{prefix}{mm}-{dd}-{seq:D2}";
    }

    private static int GetNextSequence(DbConnection conn, string 구분, string 채취일자)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) + 1 FROM `폐수_의뢰` WHERE 채취일자 = @d AND 구분 = @g";
        cmd.Parameters.AddWithValue("@d", 채취일자);
        cmd.Parameters.AddWithValue("@g", 구분);
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    // =========================================================================
    // 의뢰 생성
    // =========================================================================
    public static int CreateRequest(WasteRequest req, IEnumerable<string> 항목들)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `폐수_의뢰`
                (의뢰번호, 구분, 채취일자, 업체명, 관리번호, 상태, 등록자, 등록일시)
            VALUES
                (@no, @g, @d, @c, @cn, '대기', @u, @dt)";
        cmd.Parameters.AddWithValue("@no", req.의뢰번호);
        cmd.Parameters.AddWithValue("@g",  req.구분);
        cmd.Parameters.AddWithValue("@d",  req.채취일자);
        cmd.Parameters.AddWithValue("@c",  req.업체명);
        cmd.Parameters.AddWithValue("@cn", req.관리번호);
        cmd.Parameters.AddWithValue("@u",  req.등록자);
        cmd.Parameters.AddWithValue("@dt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();

        long reqId;
        using (var idCmd = conn.CreateCommand())
        {
            idCmd.CommandText = "SELECT LAST_INSERT_ID()";
            reqId = Convert.ToInt64(idCmd.ExecuteScalar());
        }

        foreach (var 항목 in 항목들)
        {
            using var iCmd = conn.CreateCommand();
            iCmd.CommandText = @"
                INSERT INTO `폐수_의뢰_항목` (의뢰_id, 항목, 상태)
                VALUES (@id, @h, '미담')";
            iCmd.Parameters.AddWithValue("@id", reqId);
            iCmd.Parameters.AddWithValue("@h",  항목);
            iCmd.ExecuteNonQuery();
        }

        return (int)reqId;
    }

    // =========================================================================
    // 날짜별 의뢰 항목 조회 (JOIN)
    // =========================================================================
    public static List<WasteRequestItem> GetItemsByDate(string date)
    {
        var result = new List<WasteRequestItem>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT i.id, i.의뢰_id, i.항목, i.상태, i.배정자, i.배정일시, i.완료일시,
                   r.의뢰번호, r.구분, r.업체명, r.채취일자
            FROM `폐수_의뢰_항목` i
            JOIN `폐수_의뢰` r ON r.id = i.의뢰_id
            WHERE r.채취일자 = @d
            ORDER BY r.구분, r.의뢰번호, i.항목";
        cmd.Parameters.AddWithValue("@d", date);

        using var rdr = cmd.ExecuteReader();
        while (rdr.Read())
        {
            result.Add(new WasteRequestItem
            {
                Id       = rdr.GetInt32(0),
                의뢰Id   = rdr.GetInt32(1),
                항목     = S(rdr, 2),
                상태     = S(rdr, 3),
                배정자   = S(rdr, 4),
                배정일시 = S(rdr, 5),
                완료일시 = S(rdr, 6),
                의뢰번호 = S(rdr, 7),
                구분     = S(rdr, 8),
                업체명   = S(rdr, 9),
                채취일자 = S(rdr, 10),
            });
        }
        return result;
    }

    // =========================================================================
    // 상태 변경
    // =========================================================================
    public static void AssignItem(int itemId, string 배정자)
        => UpdateItemStatus(itemId, "담음", 배정자, nameof(배정자));

    public static void CompleteItem(int itemId, string 배정자)
        => UpdateItemStatus(itemId, "완료", 배정자, "완료");

    public static void UnassignItem(int itemId)
        => UpdateItemStatus(itemId, "미담", "", "");

    private static void UpdateItemStatus(int itemId, string 상태, string user, string field)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        string timeCol = 상태 == "담음" ? "배정일시" : (상태 == "완료" ? "완료일시" : "배정일시");
        cmd.CommandText = $@"
            UPDATE `폐수_의뢰_항목`
            SET 상태 = @s,
                배정자 = @u,
                {timeCol} = @t
            WHERE id = @id";
        cmd.Parameters.AddWithValue("@s",  상태);
        cmd.Parameters.AddWithValue("@u",  user);
        cmd.Parameters.AddWithValue("@t",  상태 == "미담" ? "" : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("@id", itemId);
        cmd.ExecuteNonQuery();
    }

    // =========================================================================
    // 처리시설 작업 항목 (요일별 스케줄 매핑 + INSERT OR IGNORE)
    // =========================================================================
    public static List<FacilityWorkItem> GetFacilityItems(string date)
    {
        var result = new List<FacilityWorkItem>();
        DateTime dt = DateTime.Parse(date);
        DayOfWeek dow = dt.DayOfWeek;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 마스터 전체 조회
        var masters = new List<(int id, string 시설명, string 시료명, string 비고,
            string BOD, string TOC, string SS, string TN, string TP,
            string 총대장균군, string COD, string 염소이온, string 영양염류,
            string 함수율, string 중금속)>();

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                SELECT id, 시설명, 시료명, 비고,
                       BOD, TOC, SS, `T-N`, `T-P`,
                       총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속
                FROM `처리시설_마스터`
                ORDER BY id";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
                masters.Add((rdr.GetInt32(0), S(rdr,1), S(rdr,2), S(rdr,3),
                    S(rdr,4), S(rdr,5), S(rdr,6), S(rdr,7), S(rdr,8),
                    S(rdr,9), S(rdr,10), S(rdr,11), S(rdr,12), S(rdr,13), S(rdr,14)));
        }

        // 2. 요일에 해당하는 행만 INSERT OR IGNORE → 작업 목록 생성
        foreach (var m in masters)
        {
            if (!IsScheduledFor(m.비고, dow)) continue;

            var 항목목록 = BuildActiveList(m.BOD, m.TOC, m.SS, m.TN, m.TP,
                m.총대장균군, m.COD, m.염소이온, m.영양염류, m.함수율, m.중금속);
            if (string.IsNullOrEmpty(항목목록)) continue;

            // INSERT OR IGNORE (UNIQUE on 마스터_id, 채취일자)
            using var ins = conn.CreateCommand();
            ins.CommandText = @"INSERT IGNORE INTO `처리시설_작업`
                       (마스터_id, 채취일자, 시설명, 시료명, 항목목록, 상태)
                    VALUES (@mid, @d, @f, @s, @h, '미담')";
            ins.Parameters.AddWithValue("@mid", m.id);
            ins.Parameters.AddWithValue("@d",   date);
            ins.Parameters.AddWithValue("@f",   m.시설명);
            ins.Parameters.AddWithValue("@s",   m.시료명);
            ins.Parameters.AddWithValue("@h",   항목목록);
            ins.ExecuteNonQuery();
        }

        // 3. 현재 상태 조회
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                SELECT id, 마스터_id, 채취일자, 시설명, 시료명,
                       항목목록, 상태, 배정자, 배정일시, 완료일시
                FROM `처리시설_작업`
                WHERE 채취일자 = @d
                ORDER BY 마스터_id";
            cmd.Parameters.AddWithValue("@d", date);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var item = new FacilityWorkItem
                {
                    Id        = rdr.GetInt32(0),
                    마스터Id  = rdr.GetInt32(1),
                    채취일자  = S(rdr, 2),
                    시설명    = S(rdr, 3),
                    시료명    = S(rdr, 4),
                    항목목록  = S(rdr, 5),
                    상태      = S(rdr, 6),
                    배정자    = S(rdr, 7),
                    배정일시  = S(rdr, 8),
                    완료일시  = S(rdr, 9),
                };
                // 비고 마스터 붙이기
                var master = masters.Find(x => x.id == item.마스터Id);
                item.비고마스터 = master.비고;
                result.Add(item);
            }
        }

        return result;
    }

    // 처리시설 작업 상태 변경
    public static void AssignFacilityItem(int itemId, string 배정자)
        => UpdateFacilityStatus(itemId, "담음", 배정자);

    public static void CompleteFacilityItem(int itemId, string 배정자)
        => UpdateFacilityStatus(itemId, "완료", 배정자);

    public static void UnassignFacilityItem(int itemId)
        => UpdateFacilityStatus(itemId, "미담", "");

    private static void UpdateFacilityStatus(int itemId, string 상태, string user)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        string timeCol = 상태 == "담음" ? "배정일시" : (상태 == "완료" ? "완료일시" : "배정일시");
        cmd.CommandText = $@"
            UPDATE `처리시설_작업`
            SET 상태 = @s, 배정자 = @u, {timeCol} = @t
            WHERE id = @id";
        cmd.Parameters.AddWithValue("@s",  상태);
        cmd.Parameters.AddWithValue("@u",  user);
        cmd.Parameters.AddWithValue("@t",  상태 == "미담" ? "" : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("@id", itemId);
        cmd.ExecuteNonQuery();
    }

    // ── 업체 목록 (구분별) ───────────────────────────────────────────────
    public static List<(string 업체명, string 관리번호)> GetCompanies(string 구분)
    {
        var list = new List<(string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        // 프로젝트명으로 구분 (여수산단/율촌산단/세풍 등) - 실제 값에 따라 조정
        cmd.CommandText = @"
            SELECT 업체명, COALESCE(관리번호,'')
            FROM `폐수배출업소`
            WHERE 프로젝트명 LIKE @g
            ORDER BY 관리번호, 업체명";
        cmd.Parameters.AddWithValue("@g", $"%{구분}%");
        using var rdr = cmd.ExecuteReader();
        while (rdr.Read())
            list.Add((S(rdr, 0), S(rdr, 1)));
        return list;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static string S(DbDataReader r, int i)
    {
        try { return r.IsDBNull(i) ? "" : r.GetString(i) ?? ""; }
        catch { return ""; }
    }

    private static bool IsScheduledFor(string 비고, DayOfWeek dow)
    {
        if (string.IsNullOrWhiteSpace(비고)) return true; // 매일

        var map = new (string key, DayOfWeek day)[]
        {
            ("월", DayOfWeek.Monday),
            ("화", DayOfWeek.Tuesday),
            ("수", DayOfWeek.Wednesday),
            ("목", DayOfWeek.Thursday),
            ("금", DayOfWeek.Friday),
        };

        bool hasDay = false;
        foreach (var (key, day) in map)
        {
            if (비고.Contains(key))
            {
                hasDay = true;
                if (day == dow) return true;
            }
        }
        // "주1회" 같이 요일 미지정 → 항상 표시
        return !hasDay;
    }

    private static string BuildActiveList(
        string bod, string toc, string ss, string tn, string tp,
        string col, string cod, string cl, string nu, string mc, string hm)
    {
        var parts = new List<string>();
        if (!string.IsNullOrEmpty(bod)) parts.Add("BOD");
        if (!string.IsNullOrEmpty(toc)) parts.Add("TOC");
        if (!string.IsNullOrEmpty(ss))  parts.Add("SS");
        if (!string.IsNullOrEmpty(tn))  parts.Add("T-N");
        if (!string.IsNullOrEmpty(tp))  parts.Add("T-P");
        if (!string.IsNullOrEmpty(col)) parts.Add("총대장균군");
        if (!string.IsNullOrEmpty(cod)) parts.Add("COD");
        if (!string.IsNullOrEmpty(cl))  parts.Add("염소이온");
        if (!string.IsNullOrEmpty(nu))  parts.Add("영양염류");
        if (!string.IsNullOrEmpty(mc))  parts.Add("함수율");
        if (!string.IsNullOrEmpty(hm))  parts.Add("중금속");
        return string.Join(", ", parts);
    }
}
