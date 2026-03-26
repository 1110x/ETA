using ETA.Views.Pages;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;

namespace ETA.Services;

/// <summary>
/// "분석의뢰및결과" 테이블 접근 서비스
/// TestReportService와 동일한 DbPathHelper.DbPath 사용
/// </summary>
public static class AnalysisRequestService
{
    private static string DbPath => DbPathHelper.DbPath;

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
        Log($"DB 경로: {DbPath}  존재={File.Exists(DbPath)}");
        if (!File.Exists(DbPath)) { Log("DB 파일 없음"); return list; }

        using var conn = new SqliteConnection($"Data Source={DbPath}");
        conn.Open();

        using var chk = conn.CreateCommand();
        chk.CommandText =
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='분석의뢰및결과'";
        long tableExists = Convert.ToInt64(chk.ExecuteScalar()!);
        Log($"테이블 존재: {tableExists}");
        if (tableExists == 0) return list;

        using var cnt = conn.CreateCommand();
        cnt.CommandText = "SELECT COUNT(*) FROM \"분석의뢰및결과\"";
        long rowCount = Convert.ToInt64(cnt.ExecuteScalar()!);
        Log($"전체 행 수: {rowCount}");
        if (rowCount == 0) return list;

        using var sample = conn.CreateCommand();
        sample.CommandText =
            "SELECT \"채취일자\", \"약칭\", \"시료명\", \"견적번호\" FROM \"분석의뢰및결과\" LIMIT 3";
        using var srdr = sample.ExecuteReader();
        while (srdr.Read())
            Log($"  샘플: 채취일자={srdr.GetValue(0)} 약칭={srdr.GetValue(1)}" +
                $" 시료명={srdr.GetValue(2)} 견적번호={srdr.GetValue(3)}");

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT rowid,
                   COALESCE(""약칭"",     ''),
                   COALESCE(""시료명"",   ''),
                   COALESCE(""견적번호"", ''),
                   COALESCE(""채취일자"", '')
            FROM   ""분석의뢰및결과""
            ORDER  BY ""채취일자"" DESC";

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
        if (!File.Exists(DbPath)) return dict;

        using var conn = new SqliteConnection($"Data Source={DbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM \"분석의뢰및결과\" WHERE rowid = @id";
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
    //    표준처리기한  |    7     |          9            |      2       | ...
    //    약칭          |   SS     |         BOD           |     pH       | ...
    //    2025-07-01    |  김지은  |        정준하         |    유경지    | ...
    //    2025-07-02    |  김지은  |        정준하         |    유경지    | ...
    //
    //  반환: 컬럼 전체명 → (처리일수, 약식명)
    //    예: "생물화학적 산소요구량" → (9, "BOD")
    // =====================================================================
    public static Dictionary<string, (int days, string shortName)> GetStandardDaysInfo()
    {
        var result = new Dictionary<string, (int, string)>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(DbPath)) { Log("GetStandardDaysInfo: DB 없음"); return result; }

        try
        {
            using var conn = new SqliteConnection($"Data Source={DbPath}");
            conn.Open();

            using var chk = conn.CreateCommand();
            chk.CommandText =
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='분장표준처리'";
            if (Convert.ToInt64(chk.ExecuteScalar()!) == 0)
            { Log("GetStandardDaysInfo: 분장표준처리 테이블 없음"); return result; }

            // 처음 3행: 표준처리기한, 약칭, (날짜행)
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM \"분장표준처리\" LIMIT 3";
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
        if (!File.Exists(DbPath)) return result;

        try
        {
            using var conn = new SqliteConnection($"Data Source={DbPath}");
            conn.Open();

            // 날짜 정규화: yyyy-MM-dd
            string dateKey = sampleDate.Trim();
            if (DateTime.TryParse(dateKey, out var dt))
                dateKey = dt.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM \"분장표준처리\" WHERE \"항목명\" = @date LIMIT 1";
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
}
