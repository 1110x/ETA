using System;
using System.Collections.Generic;
using System.Data.Common;
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
    // 날짜 트리뷰 소스 (비용부담금 모드)
    //   - 폐수의뢰및결과.채수일 기준
    // =========================================================================
    public static List<string> GetMonths()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "폐수의뢰및결과")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT SUBSTR(채수일, 1, 7) AS ym
                FROM `폐수의뢰및결과`
                WHERE 채수일 IS NOT NULL AND 채수일 <> ''
                ORDER BY ym DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                var ym = r.GetString(0);
                if (!string.IsNullOrWhiteSpace(ym)) list.Add(ym);
            }
        }
        catch (Exception ex) { }
        return list;
    }

    public static List<string> GetDatesByMonth(string yearMonth)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "폐수의뢰및결과")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT 채수일
                FROM `폐수의뢰및결과`
                WHERE SUBSTR(채수일, 1, 7) = @ym
                ORDER BY 채수일 DESC";
            cmd.Parameters.AddWithValue("@ym", yearMonth);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                var d = r.GetString(0);
                if (!string.IsNullOrWhiteSpace(d)) list.Add(d);
            }
        }
        catch (Exception ex) { }
        return list;
    }

    // =========================================================================
    // 날짜별 의뢰 항목 키 집합 (배지 필터링용)
    //   - 폐수의뢰및결과 에서 해당 날짜에 값이 있는 항목 컬럼 반환
    //   - 비용부담금 모드 트리뷰 배지가 결과가 있는 항목만 표시하도록 사용
    // =========================================================================
    public static HashSet<string> GetRequestedItemSetByDate(string date)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            // 해당 날짜에 각 항목 컬럼에 값이 있으면 해당 키 추가
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT
                    MAX(CASE WHEN BOD       IS NOT NULL AND BOD       <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN `TOC`     IS NOT NULL AND `TOC`     <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN SS        IS NOT NULL AND SS        <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN `T-N`     IS NOT NULL AND `T-N`     <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN `T-P`     IS NOT NULL AND `T-P`     <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN Phenols   IS NOT NULL AND Phenols   <> '' THEN 1 ELSE 0 END),
                    MAX(CASE WHEN `N-Hexan` IS NOT NULL AND `N-Hexan` <> '' THEN 1 ELSE 0 END)
                FROM `폐수의뢰및결과`
                WHERE 채수일 = @d";
            cmd.Parameters.AddWithValue("@d", date);
            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                var keys = new[] { "BOD", "TOC", "SS", "TN", "TP", "PHENOLS", "NHEX" };
                for (int i = 0; i < keys.Length; i++)
                    if (!rdr.IsDBNull(i) && rdr.GetInt32(i) == 1)
                        set.Add(keys[i]);
            }
        }
        catch (Exception ex) { }
        return set;
    }

    private static string NormalizeAnalyteKey(string raw)
    {
        var s = (raw ?? "").Trim().ToUpperInvariant().Replace(" ", "").Replace("-", "");
        return s switch
        {
            "BOD"               => "BOD",
            "TOC"               => "TOC",
            "SS"                => "SS",
            "TN" or "TOTALN"    => "TN",
            "TP" or "TOTALP"    => "TP",
            "PHENOLS" or "PHENOL" or "페놀" or "페놀류" => "PHENOLS",
            "NHEXANE" or "NHEXAN" or "NHEX" or "노르말헥산추출물질" => "NHEX",
            "COD"               => "COD",
            _                   => "",
        };
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
    // 처리시설 작업 항목 (처리시설_측정결과 기반 — 해당 날짜에 행이 있으면 작업 표시)
    // 항목목록은 처리시설_분석계획에서 읽음 (마스터 컬럼 플래그 아님)
    // =========================================================================
    public static List<FacilityWorkItem> GetFacilityItems(string date)
    {
        var result = new List<FacilityWorkItem>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 마스터 조회 (id → 시설명/시료명/비고 매핑)
        var masterById = new Dictionary<int, (string 시설명, string 시료명, string 비고)>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT id, 시설명, 시료명, 비고 FROM `처리시설_마스터` ORDER BY id";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
                masterById[rdr.GetInt32(0)] = (S(rdr,1), S(rdr,2), S(rdr,3));
        }


        // 2. 분석계획에서 오늘 요일에 해당하는 (시설명,시료명) → 항목목록만 수집
        DateTime dt = DateTime.Parse(date);
        int planDay = ((int)dt.DayOfWeek + 6) % 7; // 월=0..일=6

        var planItems = new Dictionary<(string, string), string>();
        if (DbConnectionFactory.TableExists(conn, "처리시설_분석계획"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT 시설명, 시료명,
                       BOD, TOC, SS, `T-N`, `T-P`,
                       총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속
                FROM `처리시설_분석계획`
                WHERE 요일 = @day
                ORDER BY id";
            cmd.Parameters.AddWithValue("@day", planDay);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string 시설명 = S(rdr, 0), 시료명 = S(rdr, 1);
                string V(int i) { try { return rdr.IsDBNull(i) ? "" : rdr.GetString(i); } catch { return ""; } }
                var parts = new System.Collections.Generic.List<string>();
                if (V(2).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("BOD");
                if (V(3).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("TOC");
                if (V(4).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("SS");
                if (V(5).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("T-N");
                if (V(6).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("T-P");
                if (V(7).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("총대장균군");
                if (V(8).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("COD");
                if (V(9).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("염소이온");
                if (V(10).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("영양염류");
                if (V(11).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("함수율");
                if (V(12).StartsWith("O", StringComparison.OrdinalIgnoreCase)) parts.Add("중금속");
                if (parts.Count == 0) continue;
                planItems[(시설명, 시료명)] = string.Join(", ", parts);
            }
        }


        // 오늘 미담 행 전체 삭제 후 재삽입 (정합성 보장)
        using (var del = conn.CreateCommand())
        {
            del.CommandText = "DELETE FROM `처리시설_작업` WHERE 채취일자 = @d AND 상태 = '미담'";
            del.Parameters.AddWithValue("@d", date);
            del.ExecuteNonQuery();
        }

        // 3. 처리시설_측정결과에 오늘 행이 있고 + 오늘 분석계획에도 있는 시설만 작업 등록
        if (planItems.Count > 0 && DbConnectionFactory.TableExists(conn, "처리시설_측정결과"))
        {
            var measuredRows = new List<(int 마스터Id, string 시설명, string 시료명)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT DISTINCT 마스터_id, 시설명, 시료명 FROM `처리시설_측정결과` WHERE LEFT(채취일자,10) = @d";
                cmd.Parameters.AddWithValue("@d", date);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                    measuredRows.Add((Convert.ToInt32(rdr.GetValue(0)), S(rdr,1), S(rdr,2)));
            }

            foreach (var (mid, 시설명, 시료명) in measuredRows)
            {
                // 오늘 요일 분석계획에 없는 시설/시료는 제외
                if (!planItems.TryGetValue((시설명, 시료명), out var 항목목록)) continue;

                using var ins = conn.CreateCommand();
                ins.CommandText = @"INSERT IGNORE INTO `처리시설_작업`
                           (마스터_id, 채취일자, 시설명, 시료명, 항목목록, 상태)
                        VALUES (@mid, @d, @f, @s, @h, '미담')";
                ins.Parameters.AddWithValue("@mid", mid);
                ins.Parameters.AddWithValue("@d",   date);
                ins.Parameters.AddWithValue("@f",   시설명);
                ins.Parameters.AddWithValue("@s",   시료명);
                ins.Parameters.AddWithValue("@h",   항목목록);
                ins.ExecuteNonQuery();
            }
        }

        // 마스터 목록도 유지 (비고마스터용)
        var masters = masterById;

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
                masters.TryGetValue(item.마스터Id, out var master);
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
