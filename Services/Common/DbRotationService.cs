using System;
using System.Collections.Generic;
using System.Diagnostics;
using MySqlConnector;

namespace ETA.Services.Common;

/// <summary>
/// 6개월 단위 DB 생성 + 3년 수명 자동 삭제 관리.
///
/// ■ 규칙
///   - 매 6개월(4/1, 10/1) 새 DB를 빈 상태로 생성
///   - DB 이름: ETAS{시작yyyyMMdd}E{시작+3년-1일yyyyMMdd}
///     예) ETAS20260401E20290331
///   - 종료일이 오늘보다 이전인 DB는 자동 삭제 (3년 수명 만료)
///   - 로그인 화면에서 살아있는 모든 DB 선택 가능, 기본값 = 최신
///
/// ■ 예시
///   DB1: 20210101 ~ 20231231  (2021/1/1 생성, 2023/12/31 만료 후 삭제)
///   DB2: 20210601 ~ 20240530  (6개월 후 새 DB, 3년 수명)
///   DB3: 20220101 ~ 20241231
/// </summary>
public static class DbRotationService
{
    private static void Log(string msg)
    {
        string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [DbRotation] {msg}";
        Debug.WriteLine(line);
    }

    // ── 기간 계산 ─────────────────────────────────────────────────────────────

    /// <summary>
    /// 날짜 기준 현재 기간 시작/종료 반환.
    /// 시작 = 가장 최근 4/1 또는 10/1
    /// 종료 = 시작 + 3년 - 1일  (DB 수명)
    /// </summary>
    public static (DateTime Start, DateTime End) GetPeriod(DateTime date)
    {
        DateTime start;
        if (date.Month >= 10)
            start = new DateTime(date.Year, 10, 1);
        else if (date.Month >= 4)
            start = new DateTime(date.Year, 4, 1);
        else
            start = new DateTime(date.Year - 1, 10, 1);

        return (start, start.AddYears(3).AddDays(-1));
    }

    /// <summary>날짜 기준 DB 이름.  예) ETAS20260401E20290331</summary>
    public static string GetDbName(DateTime date)
    {
        var (s, e) = GetPeriod(date);
        return $"ETAS{s:yyyyMMdd}E{e:yyyyMMdd}";
    }

    /// <summary>현재 운영 DB 이름</summary>
    public static string CurrentDbName => GetDbName(DateTime.Today);

    // ── 시작 시 호출 ──────────────────────────────────────────────────────────

    /// <summary>
    /// 현재 기간 DB 생성(없으면) + 사용 가능한 ETAS* DB 목록 반환.
    /// 연결 1회로 두 작업을 처리 (NAS 연결 오버헤드 최소화).
    /// Login.OnLoaded에서 호출.
    /// </summary>
    public static List<string> EnsureAndGetDbs()
    {
        string currentDb = CurrentDbName;
        Log($"현재 운영 DB: {currentDb}");

        var dbs = new List<string>();
        using var conn = new MySqlConnection(DbPathHelper.BaseConnectionString);
        conn.Open();

        // ── 현재 기간 DB 생성 + 사용자 권한 부여 ────────────────────────
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText =
                $"CREATE DATABASE IF NOT EXISTS `{currentDb}` " +
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci";
            cmd.ExecuteNonQuery();
        }
        // eta_user는 eta_db.*에만 권한이 있어 신규 DB에 접근 불가 → 명시적 GRANT
        if (!string.IsNullOrEmpty(DbPathHelper.AppUser))
        {
            try
            {
                using var grantCmd = conn.CreateCommand();
                grantCmd.CommandText =
                    $"GRANT ALL PRIVILEGES ON `{currentDb}`.* TO '{DbPathHelper.AppUser}'@'%'";
                grantCmd.ExecuteNonQuery();
                using var flushCmd = conn.CreateCommand();
                flushCmd.CommandText = "FLUSH PRIVILEGES";
                flushCmd.ExecuteNonQuery();
                Log($"GRANT 완료: {DbPathHelper.AppUser}@% → {currentDb}");
            }
            catch (Exception ex)
            {
                Log($"GRANT 실패 (무시): {ex.Message}");
            }
        }

        // ── ETAS* DB 목록 조회 (같은 연결 재사용) ────────────────────────
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SHOW DATABASES LIKE 'ETAS%'";
            using var r = cmd.ExecuteReader();
            while (r.Read()) dbs.Add(r.GetString(0));
        }
        dbs.Sort(StringComparer.Ordinal);
        dbs.Reverse(); // 최신 먼저

        Log($"DB 확인 완료. 목록({dbs.Count}개): {string.Join(", ", dbs)}");
        return dbs;
    }

    // ── 유틸리티 ─────────────────────────────────────────────────────────────

/// <summary>DB명 → 표시 라벨.  ETAS20260401E20290331 → "2026-2029  (4/1 ~ 3/31)"</summary>
    public static string GetLabelFromDbName(string dbName)
    {
        if (dbName.Length < 21 || !dbName.StartsWith("ETAS", StringComparison.Ordinal))
            return dbName;
        try
        {
            if (!DateTime.TryParseExact(dbName.Substring(4, 8), "yyyyMMdd",
                    null, System.Globalization.DateTimeStyles.None, out var start)) return dbName;
            if (!TryParseEndDate(dbName, out var end)) return dbName;
            string cur = dbName == CurrentDbName ? "  ★ 현재" : "";
            return $"{start.Year}-{end.Year}  ({start.Month}/{start.Day} ~ {end.Month}/{end.Day}){cur}";
        }
        catch { return dbName; }
    }

    /// <summary>현재 기간 설명.  예) "2026-2029  (4/1/2026 ~ 3/31/2029)"</summary>
    public static string CurrentPeriodLabel
    {
        get
        {
            var (s, e) = GetPeriod(DateTime.Today);
            return $"{s.Year}-{e.Year}  ({s:M/d/yyyy} ~ {e:M/d/yyyy})";
        }
    }

    // ── 내부 헬퍼 ────────────────────────────────────────────────────────────

    private static bool TryParseEndDate(string dbName, out DateTime endDate)
    {
        endDate = default;
        // 형식: ETAS + 8자(시작) + E + 8자(종료) = 최소 21자
        if (dbName.Length < 21) return false;
        int eIdx = dbName.IndexOf('E', 4);
        if (eIdx < 0 || eIdx + 8 >= dbName.Length + 1) return false;
        return DateTime.TryParseExact(
            dbName.Substring(eIdx + 1, 8), "yyyyMMdd",
            null, System.Globalization.DateTimeStyles.None, out endDate);
    }
}
