using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class FacilityResultService
{
    // ── 메모리 캐시 ──────────────────────────────────────────────────────────
    private static List<string>? _facilityNamesCache;
    private static List<(string 시설명, string 시료명, int 마스터Id)>? _masterSamplesCache;
    private static Dictionary<string, int>? _facilityOrderCache;

    public static void InvalidateCache()
    {
        _facilityNamesCache = null;
        _masterSamplesCache = null;
        _facilityOrderCache = null;
    }

    /// <summary>DB 설정 테이블 기반 시설 정렬 순서 (설정 없으면 99)</summary>
    internal static int FacilityOrder(string name)
    {
        if (_facilityOrderCache == null)
        {
            _facilityOrderCache = new Dictionary<string, int>();
            try
            {
                var settings = GetFacilitySettings();
                foreach (var kv in settings)
                    _facilityOrderCache[kv.Key] = kv.Value.순서;
            }
            catch { }
        }
        return _facilityOrderCache.TryGetValue(name, out var order) ? order : 99;
    }

    internal static string? GetExcelPath()
    {
        var path = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..",
                         "Data", "Templates", "요일별 분석항목 정리.xlsx"));
        return File.Exists(path) ? path : null;
    }

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
        // 고정 순서: 중흥→월내→율촌→4단계→세풍→해룡(폐홀)
        list.Sort((a, b) => FacilityOrder(a).CompareTo(FacilityOrder(b)));
        _facilityNamesCache = list;
        return list;
    }

    // ── 마스터 + 측정결과 JOIN 조회 (동적 항목 지원) ───────────────────────
    public static List<FacilityResultRow> GetRows(string facility, string date)
    {
        EnsureFacilityColumnsSync(); // 분석항목과 테이블 컬럼 동기화

        var items = GetAnalysisItems(activeOnly: false);
        // 베이스 컬럼명 (백틱 처리)
        string Q(string c) => c.Contains('-') || c.Contains(' ') ? $"`{c}`" : c;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();

        var selectParts = new List<string>
        {
            "m.id AS 마스터id",
            "m.시료명",
            "m.비고 AS 비고마스터",
            "r.id AS r_id",
            "r.비고 AS 비고값",
        };
        foreach (var item in items)
        {
            var col = item.컬럼명.Trim('`');
            selectParts.Add($"m.{Q(col)} AS `{col}_활성`");
            selectParts.Add($"r.{Q(col)} AS `{col}_값`");
        }

        cmd.CommandText = $@"
            SELECT {string.Join(", ", selectParts)}
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
            static bool ActiveBool(object v) =>
                v is not DBNull && v?.ToString()?.StartsWith("O", StringComparison.OrdinalIgnoreCase) == true;
            static string Val(object v) =>
                v is DBNull ? "" : v?.ToString() ?? "";

            var row = new FacilityResultRow
            {
                Id        = reader["r_id"] is DBNull ? 0 : Convert.ToInt32(reader["r_id"]),
                마스터Id   = Convert.ToInt32(reader["마스터id"]),
                시료명     = Val(reader["시료명"]),
                비고마스터  = Val(reader["비고마스터"]),
                비고      = Val(reader["비고값"]),
            };
            foreach (var item in items)
            {
                var col = item.컬럼명.Trim('`');
                row.Active[col] = ActiveBool(reader[$"{col}_활성"]);
                row.Values[col] = Val(reader[$"{col}_값"]);
            }
            rows.Add(row);
        }
        return rows;
    }

    // ── 저장 (INSERT or UPDATE) — 동적 항목 지원 ──────────────────────────
    public static void SaveRows(string facility, string date, List<FacilityResultRow> rows, string inputUser)
    {
        EnsureFacilityColumnsSync();

        var items = GetAnalysisItems(activeOnly: false);
        // 데이터 컬럼: 분석항목 + 비고
        var dataCols = items.Select(i => i.컬럼명.Trim('`')).Concat(new[] { "비고" }).ToList();
        string Q(string c) => c.Contains('-') || c.Contains(' ') ? $"`{c}`" : c;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        foreach (var r in rows)
        {
            if (r.Id == 0)
            {
                // INSERT
                var colList = string.Join(",", dataCols.Select(Q));
                var paramList = string.Join(",", dataCols.Select((c, i) => $"@v{i}"));
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"
                    INSERT INTO `처리시설_측정결과`
                        (마스터_id, 시설명, 시료명, 채취일자, {colList}, 입력자, 입력일시)
                    VALUES
                        (@mid, @시설명, @시료명, @date, {paramList}, @user, @now)";
                cmd.Parameters.AddWithValue("@mid",    r.마스터Id);
                cmd.Parameters.AddWithValue("@시설명",  facility);
                cmd.Parameters.AddWithValue("@시료명",  r.시료명);
                cmd.Parameters.AddWithValue("@date",   date);
                for (int i = 0; i < dataCols.Count; i++)
                {
                    var col = dataCols[i];
                    var val = col == "비고" ? r.비고 : r[col];
                    cmd.Parameters.AddWithValue($"@v{i}", val ?? "");
                }
                cmd.Parameters.AddWithValue("@user", inputUser);
                cmd.Parameters.AddWithValue("@now",  now);
                cmd.ExecuteNonQuery();
            }
            else
            {
                // UPDATE
                var setList = string.Join(",", dataCols.Select((c, i) => $"{Q(c)}=@v{i}"));
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"
                    UPDATE `처리시설_측정결과` SET
                        {setList}, 입력일시=@now
                    WHERE id=@id";
                cmd.Parameters.AddWithValue("@id", r.Id);
                for (int i = 0; i < dataCols.Count; i++)
                {
                    var col = dataCols[i];
                    var val = col == "비고" ? r.비고 : r[col];
                    cmd.Parameters.AddWithValue($"@v{i}", val ?? "");
                }
                cmd.Parameters.AddWithValue("@now", now);
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

    // ── 해당 일자 처리시설 항목별 입력 현황 (동적 항목 지원) ──────────────────
    /// <summary>지정 날짜에 측정결과가 존재하는 항목 컬럼들. (컬럼명 → bool)</summary>
    public static Dictionary<string, bool> GetFillStatusForDate(string date)
    {
        EnsureFacilityColumnsSync();

        var items = GetAnalysisItems(activeOnly: true);
        var result = new Dictionary<string, bool>();
        if (items.Count == 0) return result;

        string Q(string c) => c.Contains('-') || c.Contains(' ') ? $"`{c}`" : c;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        var selectParts = items.Select((i, idx) =>
            $"SUM(CASE WHEN COALESCE({Q(i.컬럼명.Trim('`'))},'') <> '' THEN 1 ELSE 0 END) AS c{idx}").ToList();
        cmd.CommandText = $@"
            SELECT {string.Join(", ", selectParts)}
            FROM `처리시설_측정결과`
            WHERE 채취일자=@d";
        cmd.Parameters.AddWithValue("@d", date);
        using var reader = cmd.ExecuteReader();
        if (reader.Read())
        {
            for (int i = 0; i < items.Count; i++)
            {
                var v = reader[$"c{i}"];
                var key = items[i].컬럼명.Trim('`');
                result[key] = v is not DBNull && Convert.ToInt64(v) > 0;
            }
        }
        return result;
    }

    // ── 처리시설 원자료 UPSERT (배출업소 *_DATA 와 같은 스키마, 마스터_id 기준) ─
    /// <summary>처리시설 원자료 UPSERT. tableName: "처리시설_BOD_DATA" 등.
    /// extraCols: 해당 테이블 전용 추가 컬럼 (컬럼명 → 값).
    /// UNIQUE(마스터_id, 분석일) 기준으로 존재 확인 후 INSERT 또는 UPDATE.</summary>
    public static void UpsertFacilityRawData(
        string tableName, int 마스터id, string 분석일, string 시설명, string 시료명,
        Dictionary<string, string> extraCols, string 비고 = "")
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{tableName}` WHERE 마스터_id=@mid AND 분석일=@d";
            chk.Parameters.AddWithValue("@mid", 마스터id);
            chk.Parameters.AddWithValue("@d", 분석일);
            bool exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;

            using var cmd = conn.CreateCommand();
            // 공통 파라미터
            cmd.Parameters.AddWithValue("@mid",    마스터id);
            cmd.Parameters.AddWithValue("@d",      분석일);
            cmd.Parameters.AddWithValue("@시설명", 시설명 ?? "");
            cmd.Parameters.AddWithValue("@시료명", 시료명 ?? "");
            cmd.Parameters.AddWithValue("@remark", 비고 ?? "");
            cmd.Parameters.AddWithValue("@now",    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 추가 컬럼 파라미터 (@c_{i})
            var extraList = extraCols?.ToList() ?? new List<KeyValuePair<string, string>>();
            for (int i = 0; i < extraList.Count; i++)
                cmd.Parameters.AddWithValue($"@c{i}", extraList[i].Value ?? "");

            if (exists)
            {
                var setParts = new List<string> { "시설명=@시설명", "시료명=@시료명", "비고=@remark", "등록일시=@now" };
                for (int i = 0; i < extraList.Count; i++)
                    setParts.Add($"`{extraList[i].Key}`=@c{i}");
                cmd.CommandText = $"UPDATE `{tableName}` SET {string.Join(", ", setParts)} WHERE 마스터_id=@mid AND 분석일=@d";
            }
            else
            {
                var cols = new List<string> { "마스터_id", "분석일", "시설명", "시료명" };
                var vals = new List<string> { "@mid", "@d", "@시설명", "@시료명" };
                for (int i = 0; i < extraList.Count; i++)
                {
                    cols.Add($"`{extraList[i].Key}`");
                    vals.Add($"@c{i}");
                }
                cols.Add("비고"); vals.Add("@remark");
                cols.Add("등록일시"); vals.Add("@now");
                cmd.CommandText = $"INSERT INTO `{tableName}` ({string.Join(",", cols)}) VALUES ({string.Join(",", vals)})";
            }
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
        }
    }

    // ── 분석항목 ↔ 테이블 컬럼 동기화 ──────────────────────────────────────
    private static bool _columnsSynced;
    /// <summary>처리시설_분석항목 의 모든 항목이 분석계획/마스터/측정결과 3개 테이블에
    /// 컬럼으로 존재하는지 확인하고 누락된 컬럼은 자동 추가. 1회만 실행.</summary>
    public static void EnsureFacilityColumnsSync()
    {
        if (_columnsSynced) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분석항목 목록 직접 조회 (캐시와 무관)
            var items = new List<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT 컬럼명 FROM `처리시설_분석항목` ORDER BY 순서";
                using var r = cmd.ExecuteReader();
                while (r.Read()) items.Add(r.GetString(0).Trim('`'));
            }

            foreach (var table in new[] { "처리시설_분석계획", "처리시설_마스터", "처리시설_측정결과" })
            {
                // 현재 컬럼 목록 조회
                var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $@"
                        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                        WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME=@t";
                    cmd.Parameters.AddWithValue("@t", table);
                    using var r = cmd.ExecuteReader();
                    while (r.Read()) existing.Add(r.GetString(0));
                }

                foreach (var col in items)
                {
                    if (existing.Contains(col)) continue;
                    try
                    {
                        var qcol = col.Contains('-') || col.Contains(' ') ? $"`{col}`" : col;
                        using var alter = conn.CreateCommand();
                        alter.CommandText = $"ALTER TABLE `{table}` ADD COLUMN {qcol} TEXT DEFAULT ''";
                        alter.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            _columnsSynced = true;
        }
        catch (Exception ex)
        {
        }
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

    // ── 분석항목 (DB 기반 동적 관리) ────────────────────────────────────

    public record AnalysisItemInfo(int Id, string 항목명, string 컬럼명, int 순서, bool 활성, string AnalyteAlias = "");

    private static List<AnalysisItemInfo>? _analysisItemsCache;

    /// <summary>활성 분석항목 목록 (캐시됨)</summary>
    public static List<AnalysisItemInfo> GetAnalysisItems(bool activeOnly = true)
    {
        if (_analysisItemsCache == null)
        {
            _analysisItemsCache = new List<AnalysisItemInfo>();
            try
            {
                using var conn = DbConnectionFactory.CreateConnection();
                conn.Open();
                bool hasAlias = DbConnectionFactory.ColumnExists(conn, "처리시설_분석항목", "analyte_alias");
                using var cmd = conn.CreateCommand();
                cmd.CommandText = hasAlias
                    ? "SELECT id, 항목명, 컬럼명, 순서, 활성, COALESCE(analyte_alias,'') FROM `처리시설_분석항목` ORDER BY 순서"
                    : "SELECT id, 항목명, 컬럼명, 순서, 활성 FROM `처리시설_분석항목` ORDER BY 순서";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                    _analysisItemsCache.Add(new AnalysisItemInfo(
                        r.GetInt32(0), r.GetString(1), r.GetString(2),
                        Convert.ToInt32(r.GetValue(3)), Convert.ToInt32(r.GetValue(4)) != 0,
                        hasAlias ? r.GetString(5) : ""));
            }
            catch (Exception ex) { Console.Error.WriteLine($"[GetAnalysisItems] {ex.Message}: {ex}"); }
        }
        return activeOnly
            ? _analysisItemsCache.Where(i => i.활성).ToList()
            : _analysisItemsCache;
    }

    public static void InvalidateItemsCache() => _analysisItemsCache = null;

    /// <summary>활성 항목의 표시명 배열</summary>
    public static string[] AnalysisPlanItemNames
        => GetAnalysisItems().Select(i => i.항목명).ToArray();

    /// <summary>활성 항목의 DB 컬럼명 배열</summary>
    public static string[] PlanDbCols
        => GetAnalysisItems().Select(i => i.컬럼명).ToArray();

    /// <summary>분석항목 순서 저장</summary>
    public static void SaveItemOrder(List<int> orderedIds)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        for (int i = 0; i < orderedIds.Count; i++)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `처리시설_분석항목` SET 순서=@o WHERE id=@id";
            cmd.Parameters.AddWithValue("@o", i);
            cmd.Parameters.AddWithValue("@id", orderedIds[i]);
            cmd.ExecuteNonQuery();
        }
        InvalidateItemsCache();
    }

    /// <summary>분석항목 활성/비활성 토글</summary>
    public static void SetItemActive(int id, bool active)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `처리시설_분석항목` SET 활성=@a WHERE id=@id";
        cmd.Parameters.AddWithValue("@a", active ? 1 : 0);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
        InvalidateItemsCache();
    }

    /// <summary>분석항목 추가 (분석계획/마스터/측정결과 3개 테이블 컬럼 동기화)</summary>
    public static void AddAnalysisItem(string itemName)
    {
        var colName = itemName.Contains("-") ? $"`{itemName}`" : itemName;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 3개 테이블 모두에 컬럼 추가 (이미 존재하면 무시)
        foreach (var table in new[] { "처리시설_분석계획", "처리시설_마스터", "처리시설_측정결과" })
        {
            try
            {
                using var alter = conn.CreateCommand();
                alter.CommandText = $"ALTER TABLE `{table}` ADD COLUMN {colName} TEXT DEFAULT ''";
                alter.ExecuteNonQuery();
            }
            catch { /* 이미 존재하면 무시 */ }
        }

        // 항목 메타 등록
        int maxOrder = 0;
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT COALESCE(MAX(순서),0)+1 FROM `처리시설_분석항목`";
            maxOrder = Convert.ToInt32(cmd.ExecuteScalar());
        }
        using var ins = conn.CreateCommand();
        ins.CommandText = "INSERT IGNORE INTO `처리시설_분석항목` (항목명, 컬럼명, 순서) VALUES (@n, @c, @o)";
        ins.Parameters.AddWithValue("@n", itemName);
        ins.Parameters.AddWithValue("@c", colName);
        ins.Parameters.AddWithValue("@o", maxOrder);
        ins.ExecuteNonQuery();
        InvalidateItemsCache();
    }

    /// <summary>분석항목 삭제 (비활성 처리, 컬럼은 유지)</summary>
    public static void RemoveAnalysisItem(int id)
    {
        SetItemActive(id, false);
    }

    /// <summary>분석항목 이름 변경</summary>
    public static void RenameAnalysisItem(int id, string newName)
    {
        if (string.IsNullOrWhiteSpace(newName)) return;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `처리시설_분석항목` SET 항목명=@n WHERE id=@id";
        cmd.Parameters.AddWithValue("@n", newName.Trim());
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
        InvalidateItemsCache();
    }

    /// <summary>분석항목의 파싱 키(analyte_alias) 저장</summary>
    public static void SaveAnalyteAlias(int id, string alias)
    {
        if (string.IsNullOrWhiteSpace(alias)) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `처리시설_분석항목` SET analyte_alias=@a WHERE id=@id";
            cmd.Parameters.AddWithValue("@a", alias.Trim());
            cmd.Parameters.AddWithValue("@id", id);
            var rows = cmd.ExecuteNonQuery();
            InvalidateItemsCache();
        }
        catch (Exception ex) { Console.Error.WriteLine($"[SaveAnalyteAlias] ERROR: {ex.Message}"); }
    }

    /// <summary>analyte_alias -> 컬럼명(백틱 제거) 딕셔너리 반환 (파싱 저장용)</summary>
    public static Dictionary<string, string> GetAnalyteAliasMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in GetAnalysisItems(activeOnly: false))
        {
            if (!string.IsNullOrWhiteSpace(item.AnalyteAlias))
                map[item.AnalyteAlias] = item.컬럼명.Trim('`');
        }
        return map;
    }

    /// <summary>분석계획 구조 로딩 (시설 목록 + 시설별 시료 목록, 설정 순서 반영)</summary>
    public static (string[] facilities, Dictionary<string, string[]> samples)
        GetAnalysisPlanStructure()
    {
        var facilityOrder = new List<string>();
        var facilityMap = new Dictionary<string, List<string>>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT p.시설명, p.시료명
            FROM `처리시설_분석계획` p
            LEFT JOIN `처리시설_설정` s ON s.시설명 = p.시설명
            WHERE p.요일 = 0
            ORDER BY COALESCE(s.순서, 999), p.시료순서, p.id";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var fac = reader.GetString(0);
            var sample = reader.GetString(1);
            if (!facilityMap.ContainsKey(fac))
            {
                facilityOrder.Add(fac);
                facilityMap[fac] = new List<string>();
            }
            facilityMap[fac].Add(sample);
        }

        return (facilityOrder.ToArray(),
                facilityMap.ToDictionary(kv => kv.Key, kv => kv.Value.ToArray()));
    }

    /// <summary>분석계획 체크 상태 로딩 (dayIdx=-1이면 전체 OR 합산)</summary>
    /// <remarks>GetAnalysisPlanStructure 와 동일한 ORDER BY 로 시료 정렬해야
    /// 인덱스 기반 매칭이 어긋나지 않음 (시료순서 변경 시 순서 꼬임 방지)</remarks>
    public static Dictionary<string, List<bool[]>> GetAnalysisPlanState(int dayIdx = -1)
    {
        var state = new Dictionary<string, List<bool[]>>();

        var cols = PlanDbCols;
        if (cols.Length == 0) return state;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();

        if (dayIdx >= 0)
        {
            cmd.CommandText = $@"
                SELECT p.시설명, p.시료명, {string.Join(", ", cols.Select(c => "p." + c))}
                FROM `처리시설_분석계획` p
                LEFT JOIN `처리시설_설정` s ON s.시설명 = p.시설명
                WHERE p.요일 = @day
                ORDER BY COALESCE(s.순서, 999), p.시료순서, p.id";
            cmd.Parameters.AddWithValue("@day", dayIdx);
        }
        else
        {
            // 전체: 모든 요일 OR 합산 — MAX로 'O'가 하나라도 있으면 'O'
            cmd.CommandText = $@"
                SELECT p.시설명, p.시료명,
                    {string.Join(", ", cols.Select(c => $"MAX(p.{c})"))}
                FROM `처리시설_분석계획` p
                LEFT JOIN `처리시설_설정` s ON s.시설명 = p.시설명
                GROUP BY p.시설명, p.시료명, s.순서
                ORDER BY COALESCE(s.순서, 999), MIN(p.시료순서), MIN(p.id)";
        }

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var fac = reader.GetString(0);
            if (!state.ContainsKey(fac))
                state[fac] = new List<bool[]>();

            var checks = new bool[AnalysisPlanItemNames.Length];
            for (int i = 0; i < AnalysisPlanItemNames.Length; i++)
            {
                var val = reader.IsDBNull(2 + i) ? "" : reader.GetString(2 + i);
                checks[i] = val.StartsWith("O", StringComparison.OrdinalIgnoreCase);
            }
            state[fac].Add(checks);
        }
        return state;
    }

    /// <summary>분석계획 체크 상태 저장 (특정 시설의 전체 시료)</summary>
    public static void SaveAnalysisPlanState(string facility, string[] samples,
        List<bool[]> checkRows, int dayIdx)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        for (int si = 0; si < samples.Length && si < checkRows.Count; si++)
        {
            var checks = checkRows[si];
            using var cmd = conn.CreateCommand();
            var setClauses = new List<string>();
            for (int i = 0; i < PlanDbCols.Length && i < checks.Length; i++)
            {
                setClauses.Add($"{PlanDbCols[i]} = @v{i}");
                cmd.Parameters.AddWithValue($"@v{i}", checks[i] ? "O" : "");
            }
            cmd.CommandText = $@"
                UPDATE `처리시설_분석계획`
                SET {string.Join(", ", setClauses)}
                WHERE 시설명 = @fac AND 시료명 = @sample AND 요일 = @day";
            cmd.Parameters.AddWithValue("@fac", facility);
            cmd.Parameters.AddWithValue("@sample", samples[si]);
            cmd.Parameters.AddWithValue("@day", dayIdx);
            cmd.ExecuteNonQuery();
        }
    }

    /// <summary>BASE 모드: 체크 상태를 전체 요일(0~6)에 일괄 적용</summary>
    public static void ApplyBaseToAllDays(string facility, string[] samples, List<bool[]> checkRows)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        for (int si = 0; si < samples.Length && si < checkRows.Count; si++)
        {
            var checks = checkRows[si];
            for (int dayIdx = 0; dayIdx < 7; dayIdx++)
            {
                using var cmd = conn.CreateCommand();
                var setClauses = new List<string>();
                for (int i = 0; i < PlanDbCols.Length && i < checks.Length; i++)
                {
                    setClauses.Add($"{PlanDbCols[i]} = @v{i}");
                    cmd.Parameters.AddWithValue($"@v{i}", checks[i] ? "O" : "");
                }
                cmd.CommandText = $@"
                    UPDATE `처리시설_분석계획`
                    SET {string.Join(", ", setClauses)}
                    WHERE 시설명 = @fac AND 시료명 = @sample AND 요일 = @day";
                cmd.Parameters.AddWithValue("@fac", facility);
                cmd.Parameters.AddWithValue("@sample", samples[si]);
                cmd.Parameters.AddWithValue("@day", dayIdx);
                cmd.ExecuteNonQuery();
            }
        }
    }

    /// <summary>BASE 모드: 시료명 변경 → 전체 요일에 반영</summary>
    public static void RenameSampleAllDays(string facility, string oldName, string newName)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `처리시설_분석계획`
            SET 시료명 = @newName
            WHERE 시설명 = @fac AND 시료명 = @oldName";
        cmd.Parameters.AddWithValue("@fac", facility);
        cmd.Parameters.AddWithValue("@oldName", oldName);
        cmd.Parameters.AddWithValue("@newName", newName);
        cmd.ExecuteNonQuery();
    }

    /// <summary>시설 추가 (분석계획 7요일 + 마스터 동기화)</summary>
    public static void AddFacility(string facilityName, string firstSampleName = "유입수")
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        for (int day = 0; day < 7; day++)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT IGNORE INTO `처리시설_분석계획` (시설명, 시료명, 요일)
                VALUES (@fac, @sample, @day)";
            cmd.Parameters.AddWithValue("@fac", facilityName);
            cmd.Parameters.AddWithValue("@sample", firstSampleName);
            cmd.Parameters.AddWithValue("@day", day);
            cmd.ExecuteNonQuery();
        }
        EnsureMasterSample(conn, facilityName, firstSampleName);
        InvalidateCache();
    }

    /// <summary>시료 순서 저장 (전체 요일에 반영)</summary>
    public static void SaveSampleOrder(string facility, string[] orderedSamples)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        for (int i = 0; i < orderedSamples.Length; i++)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                UPDATE `처리시설_분석계획`
                SET 시료순서 = @o
                WHERE 시설명 = @f AND 시료명 = @s";
            cmd.Parameters.AddWithValue("@o", i);
            cmd.Parameters.AddWithValue("@f", facility);
            cmd.Parameters.AddWithValue("@s", orderedSamples[i]);
            cmd.ExecuteNonQuery();
        }
    }

    /// <summary>시설에 시료 추가 (분석계획 7요일 + 마스터 동기화)</summary>
    public static void AddSampleToFacility(string facility, string sampleName)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        for (int day = 0; day < 7; day++)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT IGNORE INTO `처리시설_분석계획` (시설명, 시료명, 요일)
                VALUES (@fac, @sample, @day)";
            cmd.Parameters.AddWithValue("@fac", facility);
            cmd.Parameters.AddWithValue("@sample", sampleName);
            cmd.Parameters.AddWithValue("@day", day);
            cmd.ExecuteNonQuery();
        }
        EnsureMasterSample(conn, facility, sampleName);
        InvalidateCache();
    }

    /// <summary>처리시설_마스터에 시설/시료가 없으면 추가</summary>
    private static void EnsureMasterSample(System.Data.Common.DbConnection conn, string facility, string sample)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"INSERT IGNORE INTO `처리시설_마스터` (시설명, 시료명) VALUES (@f, @s)";
        cmd.Parameters.AddWithValue("@f", facility);
        cmd.Parameters.AddWithValue("@s", sample);
        cmd.ExecuteNonQuery();
    }

    // ── 처리시설 설정 (약칭, 순서) ──────────────────────────────────────

    /// <summary>시설별 설정 조회 (약칭, 순서)</summary>
    public static Dictionary<string, (string 약칭, int 순서)> GetFacilitySettings()
    {
        var dict = new Dictionary<string, (string, int)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 시설명, COALESCE(약칭,''), 순서 FROM `처리시설_설정`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dict[r.GetString(0)] = (r.GetString(1), Convert.ToInt32(r.GetValue(2)));
        }
        catch { }
        return dict;
    }

    /// <summary>시설 약칭 저장</summary>
    public static void SaveFacilityAlias(string facility, string alias)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `처리시설_설정` (시설명, 약칭, 순서) VALUES (@f, @a, 0)
            ON DUPLICATE KEY UPDATE 약칭 = @a";
        cmd.Parameters.AddWithValue("@f", facility);
        cmd.Parameters.AddWithValue("@a", alias);
        cmd.ExecuteNonQuery();
    }

    /// <summary>시설 순서 일괄 저장</summary>
    public static void SaveFacilityOrder(string[] orderedFacilities)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        for (int i = 0; i < orderedFacilities.Length; i++)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO `처리시설_설정` (시설명, 순서) VALUES (@f, @o)
                ON DUPLICATE KEY UPDATE 순서 = @o";
            cmd.Parameters.AddWithValue("@f", orderedFacilities[i]);
            cmd.Parameters.AddWithValue("@o", i);
            cmd.ExecuteNonQuery();
        }
    }

    /// <summary>시설명 변경 (분석계획 + 설정 테이블)</summary>
    public static void RenameFacility(string oldName, string newName)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd1 = conn.CreateCommand();
        cmd1.CommandText = "UPDATE `처리시설_분석계획` SET 시설명=@n WHERE 시설명=@o";
        cmd1.Parameters.AddWithValue("@o", oldName);
        cmd1.Parameters.AddWithValue("@n", newName);
        cmd1.ExecuteNonQuery();

        using var cmd2 = conn.CreateCommand();
        cmd2.CommandText = "UPDATE `처리시설_설정` SET 시설명=@n WHERE 시설명=@o";
        cmd2.Parameters.AddWithValue("@o", oldName);
        cmd2.Parameters.AddWithValue("@n", newName);
        cmd2.ExecuteNonQuery();
    }

    // ── 월/날짜 목록 (처리시설 모드 트리뷰용) ─────────────────────────
    /// <summary>처리시설_작업에 의뢰가 존재하는 월 목록 (yyyy-MM)</summary>
    public static List<string> GetMonths()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "처리시설_작업")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT SUBSTR(채취일자, 1, 7) AS ym
                FROM `처리시설_작업`
                WHERE 채취일자 IS NOT NULL AND 채취일자 <> ''
                ORDER BY ym DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                var ym = r.GetString(0);
                if (!string.IsNullOrWhiteSpace(ym)) list.Add(ym);
            }
        }
        catch { }
        return list;
    }

    /// <summary>분석결과입력 트리뷰용 — 처리시설 모드에서 의뢰가 존재하는 일자 목록</summary>
    public static List<string> GetDatesByMonth(string yearMonth)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT 채취일자
                FROM `처리시설_작업`
                WHERE SUBSTR(채취일자, 1, 7) = @ym
                ORDER BY 채취일자 DESC";
            cmd.Parameters.AddWithValue("@ym", yearMonth);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                var d = r.GetString(0);
                if (!string.IsNullOrWhiteSpace(d)) list.Add(d);
            }
        }
        catch { }
        return list;
    }

    // ── 특정 일자에 의뢰된(예정) 처리시설 항목 집합 ───────────────────
    /// <summary>처리시설_작업 항목목록 파싱 → 해당 일자의 의뢰 항목 집합</summary>
    public static HashSet<string> GetScheduledItemsByDate(string date)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 항목목록 FROM `처리시설_작업` WHERE 채취일자=@d";
            cmd.Parameters.AddWithValue("@d", date);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                foreach (var it in r.GetString(0).Split(',', StringSplitOptions.RemoveEmptyEntries))
                    set.Add(it.Trim());
            }
        }
        catch { }
        return set;
    }

    // ── 분석계획 기반 처리시설_작업 일괄 생성 ───────────────────────────

    /// <summary>분석계획에 따라 지정 기간의 처리시설_작업을 자동 생성</summary>
    public static int GenerateFacilityWorkFromPlan(DateTime startDate, DateTime endDate)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 마스터 조회 (시설명+시료명 → 마스터id 매핑)
        var masterMap = new Dictionary<(string 시설명, string 시료명), int>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT id, 시설명, 시료명 FROM `처리시설_마스터` ORDER BY id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                masterMap[(r.GetString(1), r.GetString(2))] = Convert.ToInt32(r.GetValue(0));
        }

        // 2. 분석계획 전체 로딩 (요일별)
        var plans = new Dictionary<int, List<(string 시설명, string 시료명, string 항목목록)>>();
        for (int day = 0; day < 7; day++)
        {
            var dayPlan = new List<(string, string, string)>();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT 시설명, 시료명, BOD, TOC, SS, `T-N`, `T-P`,
                       총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속
                FROM `처리시설_분석계획`
                WHERE 요일 = @day ORDER BY id";
            cmd.Parameters.AddWithValue("@day", day);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var items = new List<string>();
                string V(int i) => r.IsDBNull(i) ? "" : r.GetString(i);
                if (V(2).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("BOD");
                if (V(3).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("TOC");
                if (V(4).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("SS");
                if (V(5).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("T-N");
                if (V(6).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("T-P");
                if (V(7).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("총대장균군");
                if (V(8).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("COD");
                if (V(9).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("염소이온");
                if (V(10).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("영양염류");
                if (V(11).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("함수율");
                if (V(12).StartsWith("O", StringComparison.OrdinalIgnoreCase)) items.Add("중금속");

                if (items.Count > 0)
                    dayPlan.Add((r.GetString(0), r.GetString(1), string.Join(",", items)));
            }
            plans[day] = dayPlan;
        }

        // 3. 날짜별 작업 생성
        int count = 0;
        for (var date = startDate; date <= endDate; date = date.AddDays(1))
        {
            // C# DayOfWeek: Sunday=0, Monday=1 → 분석계획: 월=0, 화=1, ..., 일=6
            int planDay = ((int)date.DayOfWeek + 6) % 7;
            if (!plans.TryGetValue(planDay, out var dayPlan)) continue;

            string dateStr = date.ToString("yyyy-MM-dd");
            foreach (var (시설명, 시료명, 항목목록) in dayPlan)
            {
                if (!masterMap.TryGetValue((시설명, 시료명), out var masterId))
                    continue;

                using var ins = conn.CreateCommand();
                ins.CommandText = @"INSERT IGNORE INTO `처리시설_작업`
                    (마스터_id, 채취일자, 시설명, 시료명, 항목목록, 상태)
                    VALUES (@mid, @d, @f, @s, @h, '미담')";
                ins.Parameters.AddWithValue("@mid", masterId);
                ins.Parameters.AddWithValue("@d", dateStr);
                ins.Parameters.AddWithValue("@f", 시설명);
                ins.Parameters.AddWithValue("@s", 시료명);
                ins.Parameters.AddWithValue("@h", 항목목록);
                ins.ExecuteNonQuery();
                count++;
            }
        }

        return count;
    }

    // ── 분석계획 기반 처리시설_측정결과 빈 행 자동 생성 (누락 날짜 소급 포함) ─────
    /// <summary>
    /// 처리시설_측정결과의 마지막 날짜 다음 날부터 오늘까지
    /// 분석계획에 맞는 빈 행을 생성합니다. 이미 있는 행은 스킵합니다.
    /// 프로그램을 며칠 만에 켜도 그동안 누락된 날짜가 모두 채워집니다.
    /// </summary>
    public static int EnsureTodayMeasurementResults()
    {
        var today = DateTime.Today;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 마스터 조회
        var masterMap = new Dictionary<(string 시설명, string 시료명), int>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT id, 시설명, 시료명 FROM `처리시설_마스터` ORDER BY id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                masterMap[(r.GetString(1), r.GetString(2))] = Convert.ToInt32(r.GetValue(0));
        }
        if (masterMap.Count == 0) return 0;

        // 2. 분석계획 전체 로딩 (요일 → 시료 목록)
        var planByDay = new Dictionary<int, List<(string 시설명, string 시료명)>>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT 요일, 시설명, 시료명 FROM `처리시설_분석계획` ORDER BY id";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                int day = Convert.ToInt32(r.GetValue(0));
                if (!planByDay.ContainsKey(day))
                    planByDay[day] = new List<(string, string)>();
                planByDay[day].Add((r.GetString(1), r.GetString(2)));
            }
        }
        if (planByDay.Count == 0) return 0;

        // 3. 마지막 생성 날짜 조회 → 그 다음 날부터 오늘까지 처리
        DateTime startDate;
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT MAX(채취일자) FROM `처리시설_측정결과`";
            var val = cmd.ExecuteScalar();
            if (val != null && val != DBNull.Value &&
                DateTime.TryParse(val.ToString(), out var lastDate))
                startDate = lastDate.AddDays(1); // 마지막 날 다음 날부터
            else
                startDate = today; // 데이터 없으면 오늘만
        }

        // 4. 날짜별로 없는 행만 삽입
        int totalCount = 0;
        for (var date = startDate; date <= today; date = date.AddDays(1))
        {
            string dateStr = date.ToString("yyyy-MM-dd");
            int planDay = ((int)date.DayOfWeek + 6) % 7; // 월=0..일=6

            if (!planByDay.TryGetValue(planDay, out var plan)) continue;

            var existing = new HashSet<int>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT 마스터_id FROM `처리시설_측정결과` WHERE 채취일자 = @d";
                cmd.Parameters.AddWithValue("@d", dateStr);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                    existing.Add(Convert.ToInt32(r.GetValue(0)));
            }

            foreach (var (시설명, 시료명) in plan)
            {
                if (!masterMap.TryGetValue((시설명, 시료명), out var masterId)) continue;
                if (existing.Contains(masterId)) continue;

                using var ins = conn.CreateCommand();
                ins.CommandText = @"INSERT INTO `처리시설_측정결과`
                    (마스터_id, 시설명, 시료명, 채취일자)
                    VALUES (@mid, @f, @s, @d)";
                ins.Parameters.AddWithValue("@mid", masterId);
                ins.Parameters.AddWithValue("@f", 시설명);
                ins.Parameters.AddWithValue("@s", 시료명);
                ins.Parameters.AddWithValue("@d", dateStr);
                ins.ExecuteNonQuery();
                existing.Add(masterId);
                totalCount++;
            }
        }

        return totalCount;
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
