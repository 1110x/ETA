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
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "MyTaskDebug.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [MyTaskService] {msg}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    // =========================================================================
    //  직원 이름 조회
    // =========================================================================
    public static string GetEmployeeName(string employeeId)
    {
        if (string.IsNullOrEmpty(employeeId)) return "";
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return "";
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 날짜 정규화
            if (DateTime.TryParse(dateStr, out var dt))
                dateStr = dt.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT {DbConnectionFactory.RowId}, `약칭`, `시료명`, `견적번호` FROM `분석의뢰및결과` WHERE `채취일자` = @d ORDER BY `견적번호`";
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
            cmd2.CommandText = $"SELECT {DbConnectionFactory.RowId}, `{analyteFullName}` FROM `분析의뢰및결과` WHERE `채취일자` = @d";
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
                .OrderBy(i => i.시설명)
                .ToList();
        }
        catch (Exception ex) { Log($"GetFacilityItems 오류: {ex.Message}"); return []; }
    }
}
