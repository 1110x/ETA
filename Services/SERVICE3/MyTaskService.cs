using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE3;

/// <summary>
/// 로그인한 직원에게 할당된 분석업무 조회 서비스
/// </summary>
public static class MyTaskService
{
    private static readonly string LogPath =
        Path.Combine(AppPaths.LogsDir, "MyTaskDebug.log");

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [MyTaskService] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // =========================================================================
    //  직원 이름 조회
    // =========================================================================
    public static string GetEmployeeName(string employeeId)
    {
        if (string.IsNullOrEmpty(employeeId)) return "";
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
            cmd.Parameters.AddWithValue("@id", employeeId);
            return cmd.ExecuteScalar()?.ToString()?.Trim() ?? "";
        }
        catch (Exception ex) { Log($"GetEmployeeName 오류: {ex.Message}"); return ""; }
    }

    // =========================================================================
    //  측대 — 분장표준처리에서 할당된 분석항목 목록
    // =========================================================================
    public static List<(string FullName, string ShortName)> GetMyAnalysisItems(string employeeId, DateTime date)
        => AnalysisRequestService.GetAssignmentsForAgent(employeeId, date);

    // =========================================================================
    //  측대 — 특정 날짜 + 분석항목에 해당하는 시료 목록
    // =========================================================================
    public record SampleTask(int Id, string Company, string Sample, string 접수번호);

    public static List<SampleTask> GetSamplesForAnalyte(string dateStr, string analyteFullName)
    {
        var result = new List<SampleTask>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 날짜 정규화
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT {DbConnectionFactory.RowId}, `약칭`, `시료명`, `견적번호` FROM `수질분석센터_결과` WHERE `채취일자` = @d ORDER BY `견적번호`";
            cmd.Parameters.AddWithValue("@d", dateStr);
            var baseRows = new List<(int id, string company, string sample, string 접수번호)>();
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                    baseRows.Add((Convert.ToInt32(rdr.GetValue(0)),
                        rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                        rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                        rdr.IsDBNull(3) ? "" : rdr.GetString(3)));
            }

            if (baseRows.Count == 0) return result;

            // 각 row에서 analyteFullName 컬럼이 'O'인 것만 필터
            using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = $"SELECT {DbConnectionFactory.RowId}, `{analyteFullName}` FROM `수질분석센터_결과` WHERE `채취일자` = @d";
            cmd2.Parameters.AddWithValue("@d", dateStr);

            var markedIds = new HashSet<int>();
            try
            {
                using var rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    string val = rdr2.IsDBNull(1) ? "" : rdr2.GetValue(1)?.ToString() ?? "";
                    if (val == "O")
                        markedIds.Add(Convert.ToInt32(rdr2.GetValue(0)));
                }
            }
            catch
            {
                // 컬럼이 없으면 빈 결과
                return result;
            }

            foreach (var row in baseRows)
                if (markedIds.Contains(row.id))
                    result.Add(new SampleTask(row.id, row.company, row.sample, row.접수번호));

            Log($"GetSamplesForAnalyte({dateStr}, {analyteFullName}): {result.Count}건");
        }
        catch (Exception ex) { Log($"GetSamplesForAnalyte 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  수질분석센터_결과 — 채취일자 목록 (중복 제거)
    // =========================================================================
    public static List<DateTime> GetAnalysisRequestDates()
    {
        var result = new List<DateTime>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT `채취일자` FROM `수질분석센터_결과` ORDER BY `채취일자` DESC";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var val = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString() ?? "";
                if (DateTime.TryParse(val, out var dt))
                    result.Add(dt.Date);
            }
            Log($"GetAnalysisRequestDates: {result.Count}건");
        }
        catch (Exception ex) { Log($"GetAnalysisRequestDates 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  비용부담금_결과 — 채수일 목록 (중복 제거)
    // =========================================================================
    public static List<DateTime> GetWasteRequestDates()
    {
        var result = new List<DateTime>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT `채수일` FROM `비용부담금_결과` ORDER BY `채수일` DESC";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var val = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString() ?? "";
                if (DateTime.TryParse(val, out var dt))
                    result.Add(dt.Date);
            }
            Log($"GetWasteRequestDates: {result.Count}건");
        }
        catch (Exception ex) { Log($"GetWasteRequestDates 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  여수/율촌/세풍 — 해당 구분의 폐수 의뢰 항목 (내 이름 배정 + 미배정 모두)
    // =========================================================================
    public static List<WasteRequestItem> GetWasteItems(string date, string 구분)
    {
        try
        {
            return WasteRequestService.GetItemsByDate(date)
                .Where(i => i.구분 == 구분)
                .OrderBy(i => i.업체명)
                .ThenBy(i => i.항목)
                .ToList();
        }
        catch (Exception ex) { Log($"GetWasteItems 오류: {ex.Message}"); return []; }
    }

    // =========================================================================
    //  처리시설 — 해당 날짜의 처리시설 작업 항목
    // =========================================================================
    public static List<FacilityWorkItem> GetFacilityItems(string date)
    {
        try
        {
            return WasteRequestService.GetFacilityItems(date)
                .Where(i => !string.IsNullOrEmpty(i.항목목록))
                .OrderBy(i => i.시설명)
                .ToList();
        }
        catch (Exception ex) { Log($"GetFacilityItems 오류: {ex.Message}"); return []; }
    }

    // =========================================================================
    //  의뢰목록 — 비용부담금_결과 (채수 일정)
    // =========================================================================
    public record RequestListItem(
        int    Id,
        string SN,
        string 업체명,
        string 구분,
        string 채수일,
        string 확인자,
        string 관리번호,
        string 비고,
        bool   HasResult);

    public static List<RequestListItem> GetRequestListItems(string dateStr)
    {
        var result = new List<RequestListItem>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT r.Id, r.SN, r.업체명, r.구분, r.채수일, r.확인자, r.관리번호, r.비고,
                       CASE WHEN res.id IS NOT NULL THEN 1 ELSE 0 END AS HasResult
                FROM `비용부담금_결과` r
                LEFT JOIN `폐수_결과` res ON res.관리번호 = r.관리번호 AND res.채취일자 = r.채수일
                WHERE r.채수일 = @d
                ORDER BY
                    CASE r.구분 WHEN '여수' THEN 0 WHEN '율촌' THEN 1 WHEN '세풍' THEN 2 ELSE 3 END,
                    r.순서 ASC";
            cmd.Parameters.AddWithValue("@d", dateStr);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                result.Add(new RequestListItem(
                    Convert.ToInt32(rdr.GetValue(0)),
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                    rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                    rdr.IsDBNull(4) ? "" : rdr.GetString(4),
                    rdr.IsDBNull(5) ? "" : rdr.GetString(5),
                    rdr.IsDBNull(6) ? "" : rdr.GetString(6),
                    rdr.IsDBNull(7) ? "" : rdr.GetString(7),
                    Convert.ToInt32(rdr.GetValue(8)) == 1));
            }
            Log($"GetRequestListItems({dateStr}): {result.Count}건");
        }
        catch (Exception ex) { Log($"GetRequestListItems 오류: {ex.Message}"); }
        return result;
    }

    public static List<RequestListItem> GetRequestListItemsRange(string fromDate, string toDate)
    {
        var result = new List<RequestListItem>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT r.Id, r.SN, r.업체명, r.구분, r.채수일, r.확인자, r.관리번호, r.비고,
                       CASE WHEN res.id IS NOT NULL THEN 1 ELSE 0 END AS HasResult
                FROM `비용부담금_결과` r
                LEFT JOIN `폐수_결과` res ON res.관리번호 = r.관리번호 AND res.채취일자 = r.채수일
                WHERE r.채수일 BETWEEN @from AND @to
                ORDER BY
                    r.채수일 DESC,
                    CASE r.구분 WHEN '여수' THEN 0 WHEN '율촌' THEN 1 WHEN '세풍' THEN 2 ELSE 3 END,
                    r.순서 ASC";
            cmd.Parameters.AddWithValue("@from", fromDate);
            cmd.Parameters.AddWithValue("@to",   toDate);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                result.Add(new RequestListItem(
                    Convert.ToInt32(rdr.GetValue(0)),
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                    rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                    rdr.IsDBNull(4) ? "" : rdr.GetString(4),
                    rdr.IsDBNull(5) ? "" : rdr.GetString(5),
                    rdr.IsDBNull(6) ? "" : rdr.GetString(6),
                    rdr.IsDBNull(7) ? "" : rdr.GetString(7),
                    Convert.ToInt32(rdr.GetValue(8)) == 1));
            }
            Log($"GetRequestListItemsRange({fromDate}~{toDate}): {result.Count}건");
        }
        catch (Exception ex) { Log($"GetRequestListItemsRange 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  수질분석센터 — 날짜별 의뢰 시료 전체 목록
    // =========================================================================
    public record SampleRow(int RowId, string 약칭, string 시료명, string 견적번호);

    public static List<SampleRow> GetSamplesForDate(string dateStr)
    {
        var result = new List<SampleRow>();
        try
        {
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT {DbConnectionFactory.RowId}, `약칭`, `시료명`, `견적번호` FROM `수질분석센터_결과` WHERE `채취일자` = @d ORDER BY `견적번호`, `시료명`";
            cmd.Parameters.AddWithValue("@d", dateStr);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
                result.Add(new SampleRow(
                    Convert.ToInt32(rdr.GetValue(0)),
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                    rdr.IsDBNull(3) ? "" : rdr.GetString(3)));
            Log($"GetSamplesForDate({dateStr}): {result.Count}건");
        }
        catch (Exception ex) { Log($"GetSamplesForDate 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  수질분석센터 — 특정 시료 행의 분석항목 + 담당자
    // =========================================================================
    public record AnalyteAssignment(string FullName, string ShortName, string AssignedAnalyst);

    // =========================================================================
    //  비용부담금 — 특정 비용부담금_결과 행의 분석항목 + 담당자
    // =========================================================================
    public static List<AnalyteAssignment> GetAnalytesForWasteRow(int id, string dateStr)
    {
        var result = new List<AnalyteAssignment>();
        try
        {
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            var info = AnalysisRequestService.GetStandardDaysInfo();

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분장표준처리에서 해당 날짜 담당자 매핑
            var analystMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = @d";
                cmd.Parameters.AddWithValue("@d", dateStr);
                using var rdr = cmd.ExecuteReader();
                if (rdr.Read())
                    for (int i = 1; i < rdr.FieldCount; i++)
                    {
                        string col = rdr.GetName(i).Trim();
                        string val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val))
                            analystMap[col] = val;
                    }
            }

            // 비용부담금_결과는 'O' 마커 없이 컬럼 자체가 분석항목 목록
            // 메타데이터 컬럼 제외 후 전체 컬럼을 분석항목으로 반환
            var metaCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Id", "SN", "업체명", "구분", "채수일", "확인자", "관리번호",
                "비고", "순서", "등록일시", "등록자", "수정일시", "수정자"
            };
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM `비용부담금_결과` LIMIT 0";
                using var rdr = cmd.ExecuteReader();
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string col = rdr.GetName(i).Trim();
                    if (metaCols.Contains(col)) continue;
                    string shortName = info.TryGetValue(col, out var meta) ? meta.shortName : col;
                    string analyst = analystMap.TryGetValue(col, out var a) ? a : "";
                    result.Add(new AnalyteAssignment(col, shortName, analyst));
                }
            }
            Log($"GetAnalytesForWasteRow({id}, {dateStr}): {result.Count}항목");
        }
        catch (Exception ex) { Log($"GetAnalytesForWasteRow 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  Show4 항상표시용 — 날짜 기준 전체 분석항목 + 담당자
    // =========================================================================
    public static List<AnalyteAssignment> GetAllAnalytesWithAssignments(string dateStr)
    {
        var result = new List<AnalyteAssignment>();
        try
        {
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분석정보 테이블에서 전체 항목 순서대로 로드
            var analyteList = new List<string>();
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"SELECT `Analyte` FROM `분석정보` ORDER BY `Category`, `{DbConnectionFactory.RowId}`";
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string analyte = rdr.IsDBNull(0) ? "" : rdr.GetString(0).Trim();
                    if (!string.IsNullOrEmpty(analyte))
                        analyteList.Add(analyte);
                }
            }
            catch (Exception ex2) { Log($"분석정보 쿼리 오류: {ex2.Message}"); }

            // 약식명 + 날짜별 담당자
            var info     = AnalysisRequestService.GetStandardDaysInfo();
            var managers = AnalysisRequestService.GetManagersByDate(dateStr);

            foreach (var fullName in analyteList)
            {
                string shortName = info.TryGetValue(fullName, out var meta) ? meta.shortName : "";
                string analyst   = managers.TryGetValue(fullName, out var a) ? a : "";
                result.Add(new AnalyteAssignment(fullName, shortName, analyst));
            }
            Log($"GetAllAnalytesWithAssignments({dateStr}): {result.Count}항목");
        }
        catch (Exception ex) { Log($"GetAllAnalytesWithAssignments 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  Show2 클릭용 — 특정 시료 행에서 'O' 표시된 항목 반환
    // =========================================================================
    public static HashSet<string> GetMarkedColumnsForRow(int rowId)
    {
        var marked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@id", rowId);
            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (val.Equals("O", StringComparison.OrdinalIgnoreCase))
                        marked.Add(rdr.GetName(i).Trim());
                }
            }
            Log($"GetMarkedColumnsForRow({rowId}): {marked.Count}개");
        }
        catch (Exception ex) { Log($"GetMarkedColumnsForRow 오류: {ex.Message}"); }
        return marked;
    }

    public static List<AnalyteAssignment> GetAnalytesForSampleRow(int rowId, string dateStr)
    {
        var result = new List<AnalyteAssignment>();
        try
        {
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            var info = AnalysisRequestService.GetStandardDaysInfo();

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분장표준처리에서 해당 날짜 담당자 매핑
            var analystMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = @d";
                cmd.Parameters.AddWithValue("@d", dateStr);
                using var rdr = cmd.ExecuteReader();
                if (rdr.Read())
                    for (int i = 1; i < rdr.FieldCount; i++)
                    {
                        string col = rdr.GetName(i).Trim();
                        string val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val))
                            analystMap[col] = val;
                    }
            }

            // 해당 시료 행에서 'O' 컬럼 추출
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id";
                cmd.Parameters.AddWithValue("@id", rowId);
                using var rdr = cmd.ExecuteReader();
                if (rdr.Read())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        string col = rdr.GetName(i).Trim();
                        string val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!val.Equals("O", StringComparison.OrdinalIgnoreCase)) continue;
                        string shortName = info.TryGetValue(col, out var meta) ? meta.shortName : col;
                        string analyst = analystMap.TryGetValue(col, out var a) ? a : "";
                        result.Add(new AnalyteAssignment(col, shortName, analyst));
                    }
                }
            }
            Log($"GetAnalytesForSampleRow({rowId}, {dateStr}): {result.Count}항목");
        }
        catch (Exception ex) { Log($"GetAnalytesForSampleRow 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  Show3 처리시설용 — 로그인 사용자의 해당 날짜 분장항목 (분석정보 형식)
    // =========================================================================
    public record UserAssignedItem(
        string FullName, string ShortName,
        string Category, string Unit,
        string ES, string Method, string Instrument);

    public static List<UserAssignedItem> GetUserAssignedItems(string dateStr, string employeeName)
    {
        var result = new List<UserAssignedItem>();
        if (string.IsNullOrEmpty(employeeName)) return result;
        try
        {
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            // 분장표준처리에서 해당 날짜 담당자 목록 → 내 이름으로 배정된 항목
            var managers = AnalysisRequestService.GetManagersByDate(dateStr);
            var assignedItems = managers
                .Where(kv => kv.Value.Trim() == employeeName.Trim())
                .Select(kv => kv.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (assignedItems.Count == 0)
            {
                Log($"GetUserAssignedItems({dateStr}, {employeeName}): 배정 없음");
                return result;
            }

            // 분장표준처리 약칭
            var info = AnalysisRequestService.GetStandardDaysInfo();

            // 분석정보 테이블에서 상세 정보 조회
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `Analyte`, `Category`, `unit`, `ES`, `Method`, `instrument` FROM `분석정보` ORDER BY `Category`, `{DbConnectionFactory.RowId}`";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string analyte = rdr.IsDBNull(0) ? "" : rdr.GetString(0).Trim();
                if (string.IsNullOrEmpty(analyte) || !assignedItems.Contains(analyte)) continue;
                string shortName = info.TryGetValue(analyte, out var meta) ? meta.shortName : analyte;
                result.Add(new UserAssignedItem(
                    analyte,
                    shortName,
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1).Trim(),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2).Trim(),
                    rdr.IsDBNull(3) ? "" : rdr.GetString(3).Trim(),
                    rdr.IsDBNull(4) ? "" : rdr.GetString(4).Trim(),
                    rdr.IsDBNull(5) ? "" : rdr.GetString(5).Trim()
                ));
            }
            Log($"GetUserAssignedItems({dateStr}, {employeeName}): {result.Count}항목");
        }
        catch (Exception ex) { Log($"GetUserAssignedItems 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  캘린더 툴팁용 — 월 범위 내 수질분석센터 시료 목록
    // =========================================================================
    public static Dictionary<DateTime, List<(string 약칭, string 시료명)>> GetSamplesByMonth(string startDate, string endDate)
    {
        var result = new Dictionary<DateTime, List<(string, string)>>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `채취일자`, `약칭`, `시료명` FROM `수질분석센터_결과` WHERE `채취일자` BETWEEN @s AND @e ORDER BY `채취일자`, `약칭`, `시료명`";
            cmd.Parameters.AddWithValue("@s", startDate);
            cmd.Parameters.AddWithValue("@e", endDate);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string dateStr = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString() ?? "";
                if (!DateTime.TryParse(dateStr, out var dt)) continue;
                var date = dt.Date;
                if (!result.TryGetValue(date, out var list)) result[date] = list = new();
                list.Add((
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1).Trim(),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2).Trim()
                ));
            }
        }
        catch (Exception ex) { Log($"GetSamplesByMonth 오류: {ex.Message}"); }
        return result;
    }

    // =========================================================================
    //  캘린더 툴팁용 — 월 범위 내 폐수의뢰 목록
    // =========================================================================
    public static Dictionary<DateTime, List<(string 업체명, string 구분)>> GetWasteRequestsByMonth(string startDate, string endDate)
    {
        var result = new Dictionary<DateTime, List<(string, string)>>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `채수일`, `업체명`, `구분` FROM `비용부담금_결과` WHERE `채수일` BETWEEN @s AND @e ORDER BY `채수일`, `구분`, `업체명`";
            cmd.Parameters.AddWithValue("@s", startDate);
            cmd.Parameters.AddWithValue("@e", endDate);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string dateStr = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString() ?? "";
                if (!DateTime.TryParse(dateStr, out var dt)) continue;
                var date = dt.Date;
                if (!result.TryGetValue(date, out var list)) result[date] = list = new();
                list.Add((
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1).Trim(),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2).Trim()
                ));
            }
        }
        catch (Exception ex) { Log($"GetWasteRequestsByMonth 오류: {ex.Message}"); }
        return result;
    }
}
