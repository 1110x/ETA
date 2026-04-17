using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;

namespace ETA.Services.Common;

/// <summary>
/// 직원별 메뉴 접근 권한을 관리합니다.
/// DB 테이블: AgentAccess (사번, 메뉴키, 허용여부)
/// </summary>
public static class AccessService
{
    // ── 시스템에 존재하는 전체 메뉴 정의 ────────────────────────────────────
    public static readonly List<(string Key, string Label, string Category)> AllMenus = new()
    {
        // 수질분석센터
        ("Agent",               "직원정보",           "수질분석센터"),
        ("Contract",            "계약업체",            "수질분석센터"),
        ("Quotation",           "견적/의뢰서",         "수질분석센터"),
        ("QuotationIssue",      "거래명세발행",        "수질분석센터"),
        ("TestReport",          "시험성적서",          "수질분석센터"),
        ("DbMigration",         "DB Migration",        "수질분석센터"),
        // 비용부담금
        ("WasteCompany",        "폐수배출업소",        "비용부담금"),
        ("WasteSampleList",     "분석의뢰리스트",      "비용부담금"),
        ("WasteTestReport",     "시험성적서(비용)",    "비용부담금"),
        // 처리시설
        ("ProcessingFacility",  "분석계획",            "처리시설"),
        // 분석결과입력
        ("WasteAnalysisInputWater",   "분석결과입력(수질)",    "분석결과입력"),
        ("WasteAnalysisInputWaste",   "분석결과입력(비용)",    "분석결과입력"),
        // 결과 제출
        ("ResultSubmitMeasure", "결과제출(측정인)",    "결과 제출"),
        ("ResultSubmitErp",     "결과제출(ERP)",       "결과 제출"),
        // 일정업무관리
        ("Repair",              "보수요청",            "일정업무관리"),
        ("Purchase",            "물품구매",            "일정업무관리"),
        ("Schedule",            "출장/일정관리",       "일정업무관리"),
        // 리스크관리
        ("Reagent",             "시약관리",            "리스크관리"),
        // 권한관리
        ("WaterQualityNameReconcile", "센터명칭",            "권한관리"),
        ("WasteNameReconcile",        "폐수배출업소명칭",    "권한관리"),
        ("Access",                    "접근관리",            "권한관리"),
        ("ServerManagement",          "서버관리",            "권한관리"),
    };

    // ── DB 마이그레이션 ──────────────────────────────────────────────────────
    public static void EnsureTable()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"CREATE TABLE IF NOT EXISTS `AgentAccess` (
                        `사번`    VARCHAR(50)  NOT NULL,
                        `메뉴키`  VARCHAR(100) NOT NULL,
                        `허용`    TINYINT      NOT NULL DEFAULT 1,
                        PRIMARY KEY (`사번`, `메뉴키`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    // ── 메모리 캐시 (사번별) ─────────────────────────────────────────────────
    private static readonly Dictionary<string, HashSet<string>> _allowedCache = [];
    private static readonly Dictionary<string, Dictionary<string, bool>> _mapCache = [];
    public static void InvalidateCache(string 사번) { _allowedCache.Remove(사번); _mapCache.Remove(사번); }
    public static void InvalidateAllCache() { _allowedCache.Clear(); _mapCache.Clear(); }

    // ── 특정 직원의 접근 허용 메뉴 키 목록 로드 ────────────────────────────
    /// <summary>
    /// 허용된 메뉴키 목록을 반환합니다.
    /// AgentAccess 행이 없는 메뉴 = 허용 (기본 전체 허용)
    /// 행이 있고 허용=0 인 경우만 차단
    /// </summary>
    public static HashSet<string> GetAllowedMenus(string 사번)
    {
        if (_allowedCache.TryGetValue(사번, out var cached)) return cached;

        var allowed = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);

            // AgentAccess 행이 전혀 없는 사번 = 관리자가 아직 권한 설정 전 → 전체 차단
            using var existCmd = conn.CreateCommand();
            existCmd.CommandText = "SELECT COUNT(*) FROM `AgentAccess` WHERE `사번`=@id";
            existCmd.Parameters.AddWithValue("@id", 사번);
            var rowCount = Convert.ToInt32(existCmd.ExecuteScalar());
            if (rowCount == 0) return allowed; // 전체 차단

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `메뉴키` FROM `AgentAccess` WHERE `사번`=@id AND `허용`=1";
            cmd.Parameters.AddWithValue("@id", 사번);
            using var r = cmd.ExecuteReader();
            while (r.Read()) allowed.Add(r.GetString(0));
        }
        catch (Exception ex) { }
        _allowedCache[사번] = allowed;
        return allowed;
    }

    // ── 특정 직원의 전체 허용 맵 로드 (키 → true/false) ────────────────────
    public static Dictionary<string, bool> GetAccessMap(string 사번)
    {
        if (_mapCache.TryGetValue(사번, out var cachedMap)) return cachedMap;

        // 기본 전체 차단 (화이트리스트 방식)
        var map = AllMenus.ToDictionary(m => m.Key, _ => false);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `메뉴키`, `허용` FROM `AgentAccess` WHERE `사번`=@id";
            cmd.Parameters.AddWithValue("@id", 사번);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var key   = r.GetString(0);
                var allow = r.GetInt32(1) != 0;
                if (map.ContainsKey(key)) map[key] = allow;
            }
        }
        catch (Exception ex) { }
        _mapCache[사번] = map;
        return map;
    }

    // ── 접근 권한 저장 (허용=true/false → upsert) ──────────────────────────
    public static void SetAccess(string 사번, string 메뉴키, bool 허용)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO `AgentAccess` (`사번`,`메뉴키`,`허용`) VALUES (@id,@key,@val) ON DUPLICATE KEY UPDATE `허용`=@val";
            cmd.Parameters.AddWithValue("@id",  사번);
            cmd.Parameters.AddWithValue("@key", 메뉴키);
            cmd.Parameters.AddWithValue("@val", 허용 ? 1 : 0);
            cmd.ExecuteNonQuery();
            InvalidateCache(사번);
        }
        catch (Exception ex) { }
    }

    // ── 신규 직원 기본 권한 초기화 (전체 차단) ────────────────────────────────
    public static void InitializeAccessBlocked(string 사번)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);
            foreach (var (key, _, _) in AllMenus)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO `AgentAccess` (`사번`,`메뉴키`,`허용`) VALUES (@id,@key,0) ON DUPLICATE KEY UPDATE `허용`=0";
                cmd.Parameters.AddWithValue("@id",  사번);
                cmd.Parameters.AddWithValue("@key", key);
                cmd.ExecuteNonQuery();
            }
            InvalidateCache(사번);
        }
        catch (Exception ex) { }
    }

    // ── 헬퍼 ────────────────────────────────────────────────────────────────
    private static void EnsureTableOnConn(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"CREATE TABLE IF NOT EXISTS `AgentAccess` (
                    `사번`    VARCHAR(50)  NOT NULL,
                    `메뉴키`  VARCHAR(100) NOT NULL,
                    `허용`    TINYINT      NOT NULL DEFAULT 1,
                    PRIMARY KEY (`사번`, `메뉴키`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
        try { cmd.ExecuteNonQuery(); } catch { }
    }
}
