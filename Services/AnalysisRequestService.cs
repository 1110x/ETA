using ETA.Views.Pages;
using System.Data;
using System.Data.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services;

/// <summary>
/// "분석의뢰및결과" 테이블 접근 서비스
/// TestReportService와 동일한 DbPathHelper.DbPath 사용
/// </summary>
public static class AnalysisRequestService
{

    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Treeview.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AnalysisRequestService] {msg}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    // =====================================================================
    //  전체 레코드 목록 (트리뷰용)
    // =====================================================================
    public static List<AnalysisRequestRecord> GetAllRecords()
    {
        var list = new List<AnalysisRequestRecord>();
        Log($"DB 경로: {DbPathHelper.DbPath}  존재={File.Exists(DbPathHelper.DbPath)}");
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Log("DB 파일 없음"); return list; }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        long tableExists = DbConnectionFactory.TableExists(conn, "분석의뢰및결과") ? 1L : 0L;
        Log($"테이블 존재: {tableExists}");
        if (tableExists == 0) return list;

        using var cnt = conn.CreateCommand();
        cnt.CommandText = "SELECT COUNT(*) FROM `분석의뢰및결과`";
        long rowCount = Convert.ToInt64(cnt.ExecuteScalar()!);
        Log($"전체 행 수: {rowCount}");
        if (rowCount == 0) return list;

        using var sample = conn.CreateCommand();
        sample.CommandText =
            "SELECT `채취일자`, `약칭`, `시료명`, `견적번호` FROM `분석의뢰및결과` LIMIT 3";
        using var srdr = sample.ExecuteReader();
        while (srdr.Read())
            Log($"  샘플: 채취일자={srdr.GetValue(0)} 약칭={srdr.GetValue(1)}" +
                $" 시료명={srdr.GetValue(2)} 견적번호={srdr.GetValue(3)}");

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT _id,
                   COALESCE(`약칭`,     ''),
                   COALESCE(`시료명`,   ''),
                   COALESCE(`견적번호`, ''),
                   COALESCE(`채취일자`, '')
            FROM   `분析의뢰및결과`
            ORDER  BY `채취일자` DESC";

        using var rdr = cmd.ExecuteReader();
        while (rdr.Read())
        {
            var date = rdr.GetString(4);
            string 연도 = "", 월 = "";
            if (date.Length >= 10 && (date[4] == '-' || date[4] == '/'))
            { 연도 = date[..4]; 월 = date[5..7]; }
            else if (date.Length >= 8 && int.TryParse(date[..4], out _))
            { 연도 = date[..4]; 월 = date[4..6]; }
            else if (date.Length >= 4)
            { 연도 = date[..4]; }

            list.Add(new AnalysisRequestRecord
            {
                Id       = (int)rdr.GetInt64(0),
                약칭     = rdr.GetString(1),
                시료명   = rdr.GetString(2),
                접수번호 = rdr.GetString(3),
                의뢰일   = date,
                연도     = 연도,
                월       = 월,
            });
        }
        Log($"완료: {list.Count}건");
        return list;
    }

    // =====================================================================
    //  단일 레코드 전체 컬럼 (DetailPanel 표시용)
    // =====================================================================
    public static Dictionary<string, string> GetRecordRow(int rowId)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return dict;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM `분析의뢰및결과` WHERE _id = @id";
        cmd.Parameters.AddWithValue("@id", rowId);

        using var rdr = cmd.ExecuteReader();
        if (rdr.Read())
            for (int i = 0; i < rdr.FieldCount; i++)
            {
                var col = rdr.GetName(i).Trim();
                var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString() ?? "";
                dict[col] = val;
            }
        return dict;
    }

    // =====================================================================
    //  분장표준처리 — 항목 정보 조회
    //
    //  실제 테이블 구조:
    //    항목명        | 부유물질 | 생물화학적 산소요구량 | 수소이온농도 | ...
    //    표준처리기한   |    7    |          9            |      2       | ...
    //    약칭          |   SS    |         BOD           |     pH       | ...
    //    2025-07-01   |  김지은  |        정준하         |    유경지    | ...
    //    2025-07-02   |  김지은  |        정준하         |    유경지    | ...
    //
    //  반환: 컬럼 전체명 → (처리일수, 약식명)
    //    예: "생물화학적 산소요구량" → (9, "BOD")
    // =====================================================================
    public static Dictionary<string, (int days, string shortName)> GetStandardDaysInfo()
    {
        var result = new Dictionary<string, (int, string)>(StringComparer.OrdinalIgnoreCase);
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Log("GetStandardDaysInfo: DB 없음"); return result; }

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            if (!DbConnectionFactory.TableExists(conn, "분장표준처리"))
            { Log("GetStandardDaysInfo: 분장표준처리 테이블 없음"); return result; }

            // 처음 3행: 표준처리기한, 약칭, (날짜행)
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` LIMIT 3";
            using var rdr = cmd.ExecuteReader();

            int      fc       = rdr.FieldCount;
            string[] headers  = new string[fc];
            int[]    daysArr  = new int[fc];
            string[] shortArr = new string[fc];
            bool     daysSet  = false;
            bool     shortSet = false;

            for (int i = 0; i < fc; i++)
                headers[i] = rdr.GetName(i).Trim();

            while (rdr.Read())
            {
                string label = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString()?.Trim() ?? "";

                if (!daysSet && label == "표준처리기한")
                {
                    for (int i = 1; i < fc; i++)
                    {
                        var v = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        int.TryParse(v, out daysArr[i]);
                    }
                    daysSet = true;
                    continue;
                }

                if (!shortSet && label == "약칭")
                {
                    for (int i = 1; i < fc; i++)
                        shortArr[i] = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    shortSet = true;
                    continue;
                }

                if (daysSet && shortSet) break;
            }

            for (int i = 1; i < fc; i++)
            {
                string fullName = headers[i];
                if (string.IsNullOrEmpty(fullName)) continue;
                string shortName = string.IsNullOrEmpty(shortArr[i]) ? fullName : shortArr[i];
                result[fullName] = (daysArr[i], shortName);
            }

            Log($"GetStandardDaysInfo: {result.Count}개 로드 (days={daysSet}, short={shortSet})");
        }
        catch (Exception ex) { Log($"GetStandardDaysInfo 오류: {ex.Message}"); }

        return result;
    }

    // =====================================================================
    //  분장표준처리 — 채취일자 기준 전체 담당자 한번에 조회
    //
    //  반환: 분석항목 전체명 → 담당자 이름
    //    예: "생물화학적 산소요구량" → "정준하"
    //
    //  동작: 항목명 컬럼이 채취일자와 일치하는 행을 찾아 각 컬럼 값(담당자) 반환
    // =====================================================================
    public static Dictionary<string, string> GetManagersByDate(string sampleDate)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 날짜 정규화: yyyy-MM-dd
            string dateKey = sampleDate.Trim();
            if (DateTime.TryParse(dateKey, out var dt))
                dateKey = dt.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` = @date LIMIT 1";
            cmd.Parameters.AddWithValue("@date", dateKey);

            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++) // col0 = 항목명 건너뜀
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(colName))
                        result[colName] = manager;
                }
            }
            Log($"GetManagersByDate({dateKey}): {result.Count}개 담당자 로드");
        }
        catch (Exception ex) { Log($"GetManagersByDate 오류: {ex.Message}"); }

        return result;
    }

    // ── 하위 호환용 ──────────────────────────────────────────────────────
    public static Dictionary<string, int> GetStandardDays()
    {
        var info   = GetStandardDaysInfo();
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in info)
            result[kv.Key] = kv.Value.days;
        return result;
    }

    // =====================================================================
    //  분장표준처리 — 특정 직원의 업무 분장 조회 (날짜 지정)
    //  queryDate: 조회할 날짜 (해당 월의 첫 날짜부터 마지막 날짜까지)
    //  반환: 분석항목 약식명 리스트
    // =====================================================================
    public static List<(string FullName, string ShortName)> GetAssignmentsForAgent(string employeeId, DateTime queryDate)
    {
        var assignments = new List<(string, string)>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignments;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 담당자 이름 조회
            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var result = cmd2.ExecuteScalar();
                if (result != null)
                    managerName = result.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return assignments;

            // 지정된 날짜 기준으로 할당된 항목 조회
            string queryDateStr = queryDate.ToString("yyyy-MM-dd");
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `분장표준처리` WHERE `항목명` = @date";
            cmd.Parameters.AddWithValue("@date", queryDateStr);

            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                // 분석항목 정보 가져오기
                var info = GetStandardDaysInfo();
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (manager == managerName)
                    {
                        // 약식명 찾기
                        foreach (var kv in info)
                        {
                            if (kv.Key == colName)
                            {
                                assignments.Add((kv.Key, kv.Value.shortName));
                                break;
                            }
                        }
                    }
                }
            }

            Log($"GetAssignmentsForAgent({employeeId}/{managerName}, {queryDateStr}): {assignments.Count}개 할당");
        }
        catch (Exception ex) { Log($"GetAssignmentsForAgent 오류: {ex.Message}"); }

        return assignments;
    }

    public static List<(string FullName, string ShortName)> GetAssignmentsForAgent(string employeeId)
        => GetAssignmentsForAgent(employeeId, DateTime.Today);

    // =====================================================================
    //  분장표준처리 — 기간 범위 조회 (중복 제거, 날짜 범위 내 할당된 항목 전체)
    // =====================================================================
    public static List<(string FullName, string ShortName)> GetAssignmentsForAgentRange(
        string employeeId, DateTime startDate, DateTime endDate)
    {
        var result = new List<(string, string)>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return result;

            var info = GetStandardDaysInfo();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", startDate.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", endDate.ToString("yyyy-MM-dd"));

            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (manager == managerName && seen.Add(colName))
                    {
                        if (info.TryGetValue(colName, out var meta))
                            result.Add((colName, meta.shortName));
                        else
                            result.Add((colName, colName));
                    }
                }
            }
            Log($"GetAssignmentsForAgentRange({employeeId}, {startDate:yyyy-MM-dd}~{endDate:yyyy-MM-dd}): {result.Count}개");
        }
        catch (Exception ex) { Log($"GetAssignmentsForAgentRange 오류: {ex.Message}"); }

        return result;
    }

    // =====================================================================
    //  분장표준처리 — 기간 내 특정 항목을 담당하는 직원 목록 (중복 제거)
    // =====================================================================
    /// <summary>
    /// 기간 내 특정 항목 담당자 목록 — 날짜별 중복 포함.
    /// 같은 사람이 여러 날 담당하면 여러 번 포함됩니다.
    /// </summary>
    public static List<string> GetAssigneesForAnalyteInRange(string analyteFullName, DateTime start, DateTime end)
    {
        var assignees = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignees;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", start.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", end.ToString("yyyy-MM-dd"));
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    if (!string.Equals(rdr.GetName(i).Trim(), analyteFullName,
                        StringComparison.OrdinalIgnoreCase)) continue;
                    var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(val)) assignees.Add(val);
                    break;
                }
            }
        }
        catch (Exception ex) { Log($"GetAssigneesForAnalyteInRange 오류: {ex.Message}"); }
        return assignees;
    }

    // =====================================================================
    //  분장표준처리 — 특정 날짜에 특정 항목을 담당하는 직원 목록
    // =====================================================================
    public static List<string> GetAssigneesForAnalyteOnDate(string analyteFullName, DateTime date)
    {
        var assignees = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignees;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = @date";
            cmd.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));

            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    if (!string.Equals(rdr.GetName(i).Trim(), analyteFullName,
                        StringComparison.OrdinalIgnoreCase)) continue;
                    var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(val)) assignees.Add(val);
                    break;
                }
            }
        }
        catch (Exception ex) { Log($"GetAssigneesForAnalyteOnDate 오류: {ex.Message}"); }

        return assignees;
    }

    // =====================================================================
    // =====================================================================
    //  분장표준처리 — 특정 직원의 기간 내 모든 분장을 NULL로 초기화
    // =====================================================================
    public static void ClearAssignmentsForAgent(string employeeId, DateTime startDate, DateTime endDate)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return;

            // 컬럼 목록 (항목명 제외)
            var columns = DbConnectionFactory.GetColumnNames(conn, "분장표준처리")
                .Where(c => c != "항목명").ToList();
            if (columns.Count == 0) return;

            // 해당 기간 내 이 직원이 담당인 셀을 NULL로
            var setClauses = columns.Select(
                c => $"`{c}` = CASE WHEN `{c}` = @mgr THEN NULL ELSE `{c}` END");
            var sql =
                $"UPDATE `분장표준처리` SET {string.Join(", ", setClauses)} " +
                $"WHERE `항목명` BETWEEN @start AND @end";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@mgr",   managerName);
            cmd.Parameters.AddWithValue("@start", startDate.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@end",   endDate.ToString("yyyy-MM-dd"));
            cmd.ExecuteNonQuery();

            Log($"ClearAssignmentsForAgent: {employeeId}({managerName}) " +
                $"{startDate:yyyy-MM-dd}~{endDate:yyyy-MM-dd}");
        }
        catch (Exception ex) { Log($"ClearAssignmentsForAgent 오류: {ex.Message}"); }
    }

    //  분장표준처리 — 업무 분장 추가 (기간별)
    //
    //  지정된 기간 동안 특정 분석항목을 특정 담당자로 설정
    // =====================================================================
    public static void AddAssignment(string employeeId, string analyte, DateTime startDate, DateTime endDate)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분석항목 컬럼명 확인 (전체명)
            var info = GetStandardDaysInfo();
            string? columnName = null;
            foreach (var kv in info)
            {
                if (kv.Value.shortName == analyte || kv.Key == analyte)
                {
                    columnName = kv.Key;
                    break;
                }
            }
            if (string.IsNullOrEmpty(columnName))
            {
                Log($"AddAssignment: 분석항목 '{analyte}' 컬럼을 찾을 수 없음");
                return;
            }

            // 담당자 이름 조회 (사번으로)
            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var result = cmd2.ExecuteScalar();
                if (result != null)
                    managerName = result.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName))
            {
                Log($"AddAssignment: 사번 '{employeeId}'의 이름을 찾을 수 없음");
                return;
            }

            // 기간 내 각 날짜에 대해 처리
            for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
            {
                string dateKey = date.ToString("yyyy-MM-dd");

                // 해당 날짜 행 존재 확인
                bool rowExists = false;
                using (var chk = conn.CreateCommand())
                {
                    chk.CommandText = "SELECT COUNT(*) FROM `분장표준처리` WHERE `항목명` = @date";
                    chk.Parameters.AddWithValue("@date", dateKey);
                    rowExists = Convert.ToInt64(chk.ExecuteScalar()!) > 0;
                }

                if (rowExists)
                {
                    // 업데이트
                    using var upd = conn.CreateCommand();
                    upd.CommandText = $"UPDATE `분장표준처리` SET `{columnName}` = @manager WHERE `항목명` = @date";
                    upd.Parameters.AddWithValue("@manager", managerName);
                    upd.Parameters.AddWithValue("@date", dateKey);
                    upd.ExecuteNonQuery();
                }
                else
                {
                    // 삽입 (새 행 생성)
                    // 먼저 컬럼 목록 가져오기
                    var columns = DbConnectionFactory.GetColumnNames(conn, "분장표준처리");

                    // INSERT 문 생성
                    var insertColumns = new List<string> { "`항목명`" };
                    var values = new List<string> { "@date" };
                    var parameters = new Dictionary<string, object> { { "@date", dateKey } };

                    foreach (var col in columns.Skip(1)) // 항목명 제외
                    {
                        insertColumns.Add($"`{col}`");
                        if (col == columnName)
                        {
                            values.Add("@manager");
                            parameters["@manager"] = managerName;
                        }
                        else
                        {
                            values.Add("NULL");
                        }
                    }

                    using var ins = conn.CreateCommand();
                    ins.CommandText = $"INSERT INTO `분장표준처리` ({string.Join(", ", insertColumns)}) VALUES ({string.Join(", ", values)})";
                    foreach (var param in parameters)
                        ins.Parameters.AddWithValue(param.Key, param.Value);
                    ins.ExecuteNonQuery();
                }
            }

            Log($"AddAssignment: {employeeId}({managerName})에게 {analyte}({columnName}) 분장 추가 ({startDate:yyyy-MM-dd} ~ {endDate:yyyy-MM-dd})");
        }
        catch (Exception ex) { Log($"AddAssignment 오류: {ex.Message}"); }
    }
}
