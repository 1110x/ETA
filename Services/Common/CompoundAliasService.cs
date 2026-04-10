using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;

namespace ETA.Services.Common;

/// <summary>
/// 화합물 별칭(Alias) 서비스.
/// GC/ICP/PFAS 등 다성분 분석에서 CSV 성분명의 변형을 정규화한다.
/// 예: "Dichloromethane", "DCM", "Methylene Chloride" → 표준코드 "DCM", 분석항목 "다이클로로메탄"
/// </summary>
public static class CompoundAliasService
{
    /// <summary>정규화 결과: 표준코드(테이블 라우팅용) + 분석항목(의뢰및결과 컬럼용)</summary>
    public record struct CompoundInfo(string 표준코드, string 분석항목);

    private static Dictionary<string, CompoundInfo>? _cache;
    private static readonly object _lock = new();

    // ── Public API ──────────────────────────────────────────────────

    /// <summary>별칭 → (표준코드, 분석항목) 반환. 미등록이면 null.</summary>
    public static CompoundInfo? Resolve(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return null;
        EnsureLoaded();
        return _cache!.TryGetValue(rawName.Trim(), out var info) ? info : null;
    }

    /// <summary>미등록 시 rawName을 그대로 반환 (데이터 유실 방지) + 로그 경고.</summary>
    public static CompoundInfo ResolveOrFallback(string rawName)
    {
        var result = Resolve(rawName);
        if (result != null) return result.Value;

        Debug.WriteLine($"[CompoundAlias] 미등록 성분명: '{rawName}' — 원본명으로 처리됨. 화합물별명 테이블에 등록 필요.");
        return new CompoundInfo(rawName, rawName);
    }

    /// <summary>캐시 무효화 (별칭 추가/삭제 후 호출).</summary>
    public static void InvalidateCache()
    {
        lock (_lock) { _cache = null; }
    }

    // ── 별칭 CRUD ───────────────────────────────────────────────────

    /// <summary>새 별칭 등록 또는 기존 별칭 업데이트 + 캐시 갱신.</summary>
    public static bool AddOrUpdateAlias(string alias, string standardCode, string analyteName)
    {
        if (string.IsNullOrWhiteSpace(alias)) return false;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO `화합물별명`
                (`별명`, `표준코드`, `분석항목`, `등록일시`)
                VALUES (@a, @c, @n, @t)
                ON DUPLICATE KEY UPDATE `표준코드` = @c, `분석항목` = @n, `등록일시` = @t";
            cmd.Parameters.AddWithValue("@a", alias.Trim());
            cmd.Parameters.AddWithValue("@c", standardCode.Trim());
            cmd.Parameters.AddWithValue("@n", analyteName.Trim());
            cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.ExecuteNonQuery();

            // 캐시에 즉시 반영
            EnsureLoaded();
            lock (_lock)
            {
                _cache![alias.Trim()] = new CompoundInfo(standardCode.Trim(), analyteName.Trim());
            }
            Debug.WriteLine($"[CompoundAlias] 별칭 등록/갱신: '{alias}' → 표준코드='{standardCode}', 분석항목='{analyteName}'");
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 별칭 등록 실패: {ex.Message}");
            return false;
        }
    }

    /// <summary>새 별칭 등록 (INSERT IGNORE — seed용).</summary>
    public static bool AddAlias(string alias, string standardCode, string analyteName)
    {
        if (string.IsNullOrWhiteSpace(alias)) return false;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT IGNORE INTO `화합물별명`
                (`별명`, `표준코드`, `분석항목`, `등록일시`)
                VALUES (@a, @c, @n, @t)";
            cmd.Parameters.AddWithValue("@a", alias.Trim());
            cmd.Parameters.AddWithValue("@c", standardCode.Trim());
            cmd.Parameters.AddWithValue("@n", analyteName.Trim());
            cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            var affected = cmd.ExecuteNonQuery();

            if (affected > 0)
            {
                EnsureLoaded();
                lock (_lock)
                {
                    _cache![alias.Trim()] = new CompoundInfo(standardCode.Trim(), analyteName.Trim());
                }
            }
            return affected > 0;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 별칭 등록 실패: {ex.Message}");
            return false;
        }
    }

    /// <summary>별칭 삭제.</summary>
    public static void RemoveAlias(string alias)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `화합물별명` WHERE `별명` = @a";
            cmd.Parameters.AddWithValue("@a", alias.Trim());
            cmd.ExecuteNonQuery();
            InvalidateCache();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 별칭 삭제 실패: {ex.Message}");
        }
    }

    /// <summary>전체 별칭 목록 조회.</summary>
    public static List<(string 별명, string 표준코드, string 분석항목)> GetAll()
    {
        var list = new List<(string, string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `별명`, `표준코드`, `분석항목` FROM `화합물별명` ORDER BY `표준코드`, `별명`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add((r.GetString(0), r.GetString(1), r.GetString(2)));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 목록 조회 실패: {ex.Message}");
        }
        return list;
    }

    /// <summary>DISTINCT 표준코드 + 분석항목 목록 (시험기록부 테이블 생성용).</summary>
    public static List<(string 표준코드, string 분석항목)> GetDistinctStandardCodes()
    {
        var list = new List<(string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT `표준코드`, `분석항목` FROM `화합물별명`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add((r.GetString(0), r.GetString(1)));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 표준코드 목록 조회 실패: {ex.Message}");
        }
        return list;
    }

    /// <summary>특정 분석항목의 기존 표준코드를 찾는다 (드래그앤드랍 매칭 시 사용).</summary>
    public static string? FindStandardCodeByAnalyte(string analyteName)
    {
        EnsureLoaded();
        var match = _cache!.Values.FirstOrDefault(v =>
            v.분석항목.Equals(analyteName, StringComparison.OrdinalIgnoreCase));
        return match.표준코드 is { Length: > 0 } ? match.표준코드 : null;
    }

    // ── DB 스키마 ───────────────────────────────────────────────────

    /// <summary>화합물별명 테이블 생성 (없으면).</summary>
    public static void EnsureTable(DbConnection conn)
    {
        try
        {
            if (DbConnectionFactory.TableExists(conn, "화합물별명")) return;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"CREATE TABLE `화합물별명` (
                `id`       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                `별명`     VARCHAR(191) NOT NULL,
                `표준코드` VARCHAR(100) NOT NULL,
                `분석항목` VARCHAR(100) NOT NULL,
                `비고`     TEXT DEFAULT '',
                `등록일시` VARCHAR(30) DEFAULT '',
                UNIQUE(`별명`)
            )";
            cmd.ExecuteNonQuery();
            Debug.WriteLine("[CompoundAlias] 화합물별명 테이블 생성 완료");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompoundAlias] 테이블 생성 실패: {ex.Message}");
        }
    }

    /// <summary>초기 Seed 데이터 삽입 (분석정보 기반 한글명 검증 포함).</summary>
    public static void SeedIfNeeded(DbConnection conn)
    {
        // 분석정보에서 권위 있는 한글명 수집
        var analyteSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var a = r.GetString(0).Trim();
                if (!string.IsNullOrEmpty(a)) analyteSet.Add(a);
            }
        }
        catch { }

        // 별칭, 표준코드, 분석항목(분석정보.Analyte와 일치해야 함)
        var seeds = new List<(string alias, string code, string analyte)>
        {
            // ── VOC10 ──
            ("1,1-DCE", "1,1-DCE", "1,1-다이클로로에틸렌"),
            ("1,1-Dichloroethylene", "1,1-DCE", "1,1-다이클로로에틸렌"),
            ("1,1-다이클로로에틸렌", "1,1-DCE", "1,1-다이클로로에틸렌"),
            ("DCM", "DCM", "다이클로로메탄"),
            ("Dichloromethane", "DCM", "다이클로로메탄"),
            ("Methylene Chloride", "DCM", "다이클로로메탄"),
            ("다이클로로메탄", "DCM", "다이클로로메탄"),
            ("Chloroform", "Chloroform", "클로로폼"),
            ("Trichloromethane", "Chloroform", "클로로폼"),
            ("클로로폼", "Chloroform", "클로로폼"),
            ("Carbon Chloride", "Carbon_Chloride", "사염화탄소"),
            ("Carbon tetrachloride", "Carbon_Chloride", "사염화탄소"),
            ("사염화탄소", "Carbon_Chloride", "사염화탄소"),
            ("1,2-DCE", "1,2-DCE", "1,2-다이클로로에탄"),
            ("1,2-Dichloroethane", "1,2-DCE", "1,2-다이클로로에탄"),
            ("1,2-다이클로로에탄", "1,2-DCE", "1,2-다이클로로에탄"),
            ("Benzene", "Benzene", "벤젠"),
            ("벤젠", "Benzene", "벤젠"),
            ("TCE", "TCE", "트리클로로에틸렌"),
            ("Trichloroethylene", "TCE", "트리클로로에틸렌"),
            ("트리클로로에틸렌", "TCE", "트리클로로에틸렌"),
            ("Toluene", "Toluene", "톨루엔"),
            ("톨루엔", "Toluene", "톨루엔"),
            ("PCE", "PCE", "테트라클로로에틸렌"),
            ("Tetrachloroethylene", "PCE", "테트라클로로에틸렌"),
            ("테트라클로로에틸렌", "PCE", "테트라클로로에틸렌"),
            ("Xylene", "Xylene", "자일렌"),
            ("자일렌", "Xylene", "자일렌"),

            // ── VOC3 ──
            ("Vinylchloride", "Vinylchloride", "염화비닐"),
            ("Vinyl chloride", "Vinylchloride", "염화비닐"),
            ("염화비닐", "Vinylchloride", "염화비닐"),
            ("Acrylonitrile", "Acrylonitrile", "아크릴로니트릴"),
            ("아크릴로니트릴", "Acrylonitrile", "아크릴로니트릴"),
            ("Bromoform", "Bromoform", "브로모포름"),
            ("브로모포름", "Bromoform", "브로모포름"),

            // ── 단일 GC ──
            ("1,4-Dioxane", "1,4-Dioxane", "1,4-다이옥산"),
            ("1,4-다이옥산", "1,4-Dioxane", "1,4-다이옥산"),
            ("Naphthalene", "Naphthalene", "나프탈렌"),
            ("나프탈렌", "Naphthalene", "나프탈렌"),
            ("Styrene", "Styrene", "스타이렌"),
            ("스타이렌", "Styrene", "스타이렌"),
            ("Acrylamide", "Acrylamide", "아크릴아미드"),
            ("아크릴아미드", "Acrylamide", "아크릴아미드"),
            ("Epichlorohydrin", "Epichlorohydrin", "에피클로로하이드린"),
            ("에피클로로하이드린", "Epichlorohydrin", "에피클로로하이드린"),
            ("Formaldehyde", "Formaldehyde", "폼알데하이드"),
            ("폼알데하이드", "Formaldehyde", "폼알데하이드"),
            ("Pentachlorophenol", "Pentachlorophenol", "펜타클로로페놀"),
            ("펜타클로로페놀", "Pentachlorophenol", "펜타클로로페놀"),
            ("Phenol", "Phenol", "페놀"),
            ("페놀", "Phenol", "페놀"),
            ("Org-P", "Org-P", "유기인"),
            ("유기인", "Org-P", "유기인"),
            ("PCB", "PCB", "폴리클로리네이티드비페닐"),
            ("폴리클로리네이티드비페닐", "PCB", "폴리클로리네이티드비페닐"),
            ("DEHA", "DEHA", "다이에틸헥실아디페이트"),
            ("다이에틸헥실아디페이트", "DEHA", "다이에틸헥실아디페이트"),
            ("DEHP", "DEHP", "다이에틸헥실프탈레이트"),
            ("다이에틸헥실프탈레이트", "DEHP", "다이에틸헥실프탈레이트"),

            // ── ICP 금속류 ──
            ("Cu", "Cu", "구리"), ("구리", "Cu", "구리"),
            ("Zn", "Zn", "아연"), ("아연", "Zn", "아연"),
            ("Pb", "Pb", "납"), ("납", "Pb", "납"),
            ("Cd", "Cd", "카드뮴"), ("카드뮴", "Cd", "카드뮴"),
            ("Cr", "Cr", "크롬"), ("크롬", "Cr", "크롬"),
            ("Ni", "Ni", "니켈"), ("니켈", "Ni", "니켈"),
            ("Fe", "Fe", "철"), ("철", "Fe", "철"),
            ("Mn", "Mn", "망간"), ("망간", "Mn", "망간"),
            ("As", "As", "비소"), ("비소", "As", "비소"),
            ("Hg", "Hg", "수은"), ("수은", "Hg", "수은"),
            ("Se", "Se", "셀레늄"), ("셀레늄", "Se", "셀레늄"),

            // ── PFAS ──
            ("PFOA", "PFOA", "과불화옥탄산(PFOA)"),
            ("과불화옥탄산(PFOA)", "PFOA", "과불화옥탄산(PFOA)"),
            ("PFOS", "PFOS", "과불화옥탄술폰산(PFOS)"),
            ("과불화옥탄술폰산(PFOS)", "PFOS", "과불화옥탄술폰산(PFOS)"),
            ("PFBS", "PFBS", "PFBS"),
        };

        // 분석정보에 없는 분석항목명 경고
        foreach (var (alias, code, analyte) in seeds)
        {
            if (!string.IsNullOrEmpty(analyte) && analyteSet.Count > 0 && !analyteSet.Contains(analyte))
                Debug.WriteLine($"[CompoundAlias] Seed 경고: 분석항목 '{analyte}' 이 분석정보 테이블에 없음 (별명: '{alias}')");
        }

        // INSERT IGNORE로 중복 안전 삽입
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        foreach (var (alias, code, analyte) in seeds)
        {
            try
            {
                using var ins = conn.CreateCommand();
                ins.CommandText = @"INSERT IGNORE INTO `화합물별명`
                    (`별명`, `표준코드`, `분석항목`, `비고`, `등록일시`)
                    VALUES (@a, @c, @n, 'seed', @t)";
                ins.Parameters.AddWithValue("@a", alias);
                ins.Parameters.AddWithValue("@c", code);
                ins.Parameters.AddWithValue("@n", analyte);
                ins.Parameters.AddWithValue("@t", now);
                ins.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CompoundAlias] Seed 삽입 실패 ('{alias}'): {ex.Message}");
            }
        }

        Debug.WriteLine($"[CompoundAlias] Seed 완료 — {seeds.Count}개 항목 처리");
    }

    // ── 내부 ────────────────────────────────────────────────────────

    private static void EnsureLoaded()
    {
        if (_cache != null) return;
        lock (_lock)
        {
            if (_cache != null) return;
            _cache = new Dictionary<string, CompoundInfo>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
                EnsureTable(conn);
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT `별명`, `표준코드`, `분석항목` FROM `화합물별명`";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    var alias   = r.GetString(0).Trim();
                    var code    = r.GetString(1).Trim();
                    var analyte = r.GetString(2).Trim();
                    if (!string.IsNullOrEmpty(alias))
                        _cache[alias] = new CompoundInfo(code, analyte);
                }
                Debug.WriteLine($"[CompoundAlias] 캐시 로드 완료 — {_cache.Count}개 별칭");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CompoundAlias] 캐시 로드 실패: {ex.Message}");
            }
        }
    }
}
