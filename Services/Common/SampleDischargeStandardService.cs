using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;

namespace ETA.Services.Common;

/// <summary>
/// (약칭, 시료명) → 방류기준(방류기준표.구분) 매핑 서비스.
/// 의뢰서 작성 시 시료명 선택으로 방류기준이 자동 채워지도록 사용.
/// </summary>
public static class SampleDischargeStandardService
{
    private static Dictionary<(string abbr, string sample), string>? _cache;
    private static readonly object _lock = new();

    // ── Public API ──────────────────────────────────────────────────

    /// <summary>매핑된 방류기준 반환. 없으면 빈 문자열.</summary>
    public static string Resolve(string 약칭, string 시료명)
    {
        if (string.IsNullOrWhiteSpace(시료명)) return "";
        EnsureLoaded();
        var key = (Normalize(약칭), Normalize(시료명));
        return _cache!.TryGetValue(key, out var v) ? v : "";
    }

    /// <summary>특정 약칭의 모든 매핑 반환.</summary>
    public static List<(string 시료명, string 방류기준)> GetByAbbr(string 약칭)
    {
        EnsureLoaded();
        var norm = Normalize(약칭);
        lock (_lock)
        {
            return _cache!
                .Where(kv => kv.Key.abbr == norm)
                .Select(kv => (kv.Key.sample, kv.Value))
                .ToList();
        }
    }

    /// <summary>신규/갱신. UPSERT.</summary>
    public static bool AddOrUpdate(string 약칭, string 시료명, string 방류기준)
    {
        if (string.IsNullOrWhiteSpace(시료명) || string.IsNullOrWhiteSpace(방류기준)) return false;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            EnsureTable(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO `시료_방류기준_매핑`
                (`약칭`, `시료명`, `방류기준`, `등록일시`)
                VALUES (@abbr, @sample, @std, @t)
                ON DUPLICATE KEY UPDATE `방류기준`=VALUES(`방류기준`), `등록일시`=VALUES(`등록일시`)";
            cmd.Parameters.AddWithValue("@abbr",   Normalize(약칭));
            cmd.Parameters.AddWithValue("@sample", Normalize(시료명));
            cmd.Parameters.AddWithValue("@std",    방류기준.Trim());
            cmd.Parameters.AddWithValue("@t",      DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.ExecuteNonQuery();

            EnsureLoaded();
            lock (_lock)
            {
                _cache![(Normalize(약칭), Normalize(시료명))] = 방류기준.Trim();
            }
            return true;
        }
        catch { return false; }
    }

    /// <summary>매핑 삭제.</summary>
    public static bool Remove(string 약칭, string 시료명)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "시료_방류기준_매핑")) return false;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `시료_방류기준_매핑` WHERE `약칭`=@abbr AND `시료명`=@sample";
            cmd.Parameters.AddWithValue("@abbr",   Normalize(약칭));
            cmd.Parameters.AddWithValue("@sample", Normalize(시료명));
            cmd.ExecuteNonQuery();

            EnsureLoaded();
            lock (_lock)
            {
                _cache!.Remove((Normalize(약칭), Normalize(시료명)));
            }
            return true;
        }
        catch { return false; }
    }

    public static void InvalidateCache()
    {
        lock (_lock) { _cache = null; }
    }

    // ── DB 스키마 ───────────────────────────────────────────────────

    public static void EnsureTable()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            EnsureTable(conn);
        }
        catch { }
    }

    public static void EnsureTable(DbConnection conn)
    {
        try
        {
            if (DbConnectionFactory.TableExists(conn, "시료_방류기준_매핑")) return;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"CREATE TABLE `시료_방류기준_매핑` (
                `_id`      INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                `약칭`     VARCHAR(100) NOT NULL DEFAULT '',
                `시료명`   VARCHAR(200) NOT NULL,
                `방류기준` VARCHAR(100) NOT NULL,
                `등록일시` VARCHAR(30)  DEFAULT '',
                UNIQUE(`약칭`, `시료명`)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
            cmd.ExecuteNonQuery();
        }
        catch { }
    }

    // ── 내부 ────────────────────────────────────────────────────────

    private static string Normalize(string? s) => (s ?? "").Trim();

    private static void EnsureLoaded()
    {
        if (_cache != null) return;
        lock (_lock)
        {
            if (_cache != null) return;
            _cache = new Dictionary<(string, string), string>();
            try
            {
                using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
                EnsureTable(conn);
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT `약칭`, `시료명`, `방류기준` FROM `시료_방류기준_매핑`";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    var abbr   = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                    var sample = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                    var std    = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                    if (sample.Length > 0 && std.Length > 0)
                        _cache[(abbr, sample)] = std;
                }
            }
            catch { }
        }
    }
}
