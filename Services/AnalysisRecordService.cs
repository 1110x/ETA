using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ETA.Services;

/// <summary>분석기록부 Excel 파일 생성 서비스</summary>
public static class AnalysisRecordService
{
    private static string DbPath => DbPathHelper.DbPath;

    private static readonly string DataRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ETA", "Data");

    private static string TemplatePath =>
        Path.Combine(DataRoot, "Templates", "분석기록부.xlsx");

    // ── 출력 폴더: 바탕화면\시험의뢰서 출력\yyyy-MM-dd ───────────────────
    public static string GetOutputDir()
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var dir = Path.Combine(desktop, "시험의뢰서 출력", DateTime.Today.ToString("yyyy-MM-dd"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    // ── 항목별 그룹 조회 ─────────────────────────────────────────────────
    /// <summary>Key=분석항목 컬럼명, Value=해당 항목이 "O"인 행 목록</summary>
    public static Dictionary<string, List<Dictionary<string, string>>>
        GetAnalysisGroups(string 견적번호)
    {
        var groups = new Dictionary<string, List<Dictionary<string, string>>>(
            StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(DbPath)) return groups;
        try
        {
            var rows         = GetOrderRows(견적번호);
            var analysisCols = OrderRequestService.GetAnalysisColumns();

            foreach (var row in rows)
                foreach (var col in analysisCols)
                {
                    if (!row.TryGetValue(col, out var val) || val != "O") continue;
                    if (!groups.ContainsKey(col)) groups[col] = new();
                    groups[col].Add(row);
                }
        }
        catch (Exception ex) { Log($"GetAnalysisGroups 오류: {ex.Message}"); }
        return groups;
    }

    /// <summary>방류기준표 전체 로드: 항목명 → (기준유형 → 기준값)</summary>
    public static Dictionary<string, Dictionary<string, string>> Load방류기준표()
        => Get방류기준표Internal();

    // ── 분석자별 파일 생성 ───────────────────────────────────────────────
    /// <summary>
    /// 분장표준처리에서 targetDate 기준 담당자를 조회해,
    /// 분석자별로 파일 하나씩 생성 (시트 = 담당 항목).
    /// 분장 미등록 항목은 "미배정" 파일로 묶음.
    /// 반환값: (담당자명, 파일경로) 목록
    /// </summary>
    public static List<(string Assignee, string FilePath)> GenerateByAssignee(
        Dictionary<string, List<Dictionary<string, string>>> groups,
        Dictionary<string, Dictionary<string, string>> 방류기준표,
        string outputDir,
        DateTime targetDate)
    {
        var results = new List<(string, string)>();
        if (!File.Exists(TemplatePath)) { Log($"템플릿 없음: {TemplatePath}"); return results; }

        // 1) 분석자별로 그룹 재구성
        var byAssignee = new Dictionary<string, Dictionary<string, List<Dictionary<string, string>>>>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var kv in groups)
        {
            string analyteName = kv.Key;
            string assignee    = GetAssigneeForAnalyte(analyteName, targetDate) ?? "미배정";
            if (!byAssignee.ContainsKey(assignee))
                byAssignee[assignee] = new(StringComparer.OrdinalIgnoreCase);
            byAssignee[assignee][analyteName] = kv.Value;
        }

        // 2) 분석자별 파일 생성
        foreach (var kvA in byAssignee)
        {
            string assignee   = kvA.Key;
            var    analyteMap = kvA.Value;
            try
            {
                string fileName = $"{assignee} 기록부 - {targetDate:yyyy MM dd}.xlsx";
                string filePath = Path.Combine(outputDir, fileName);

                using var wb = new XLWorkbook(TemplatePath);
                var templateSheet = wb.Worksheets.First();

                bool isFirst = true;
                foreach (var kv in analyteMap)
                {
                    IXLWorksheet ws;
                    string sheetName = kv.Key.Length > 31 ? kv.Key[..31] : kv.Key;
                    if (isFirst)
                    {
                        ws = templateSheet;
                        ws.Name = sheetName;
                        isFirst = false;
                    }
                    else
                    {
                        ws = templateSheet.CopyTo(sheetName);
                        int last = ws.LastRowUsed()?.RowNumber() ?? 3;
                        for (int r = 4; r <= last; r++)
                            for (int c = 1; c <= 6; c++)
                                ws.Cell(r, c).Value = Blank.Value;
                    }
                    FillSheet(ws, kv.Key, kv.Value, 방류기준표);
                }

                wb.SaveAs(filePath);
                Log($"생성: {fileName} ({analyteMap.Count}개 시트)");
                results.Add((assignee, filePath));
            }
            catch (Exception ex) { Log($"GenerateByAssignee 오류 ({assignee}): {ex.Message}"); }
        }
        return results;
    }

    // ── 내부: 분석자/날짜로 담당자 조회 ─────────────────────────────────
    private static string? GetAssigneeForAnalyte(string analyteName, DateTime date)
    {
        if (!File.Exists(DbPath)) return null;
        try
        {
            using var conn = new SqliteConnection($"Data Source={DbPath}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM \"분장표준처리\" WHERE \"항목명\" = @date";
            cmd.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));
            using var rdr = cmd.ExecuteReader();
            if (!rdr.Read()) return null;
            for (int i = 1; i < rdr.FieldCount; i++)
            {
                if (!string.Equals(rdr.GetName(i).Trim(), analyteName,
                    StringComparison.OrdinalIgnoreCase)) continue;
                var val = rdr.IsDBNull(i) ? null : rdr.GetValue(i)?.ToString()?.Trim();
                return string.IsNullOrEmpty(val) ? null : val;
            }
        }
        catch (Exception ex) { Log($"GetAssigneeForAnalyte 오류: {ex.Message}"); }
        return null;
    }

    // ── 내부: 시트에 데이터 행 채우기 ────────────────────────────────────
    private static void FillSheet(
        IXLWorksheet ws,
        string colName,
        List<Dictionary<string, string>> rows,
        Dictionary<string, Dictionary<string, string>> 방류기준표)
    {
        방류기준표.TryGetValue(colName, out var 기준Dict);
        int rowIdx = 4;
        foreach (var data in rows)
        {
            data.TryGetValue("견적번호",              out var 견적번호);
            data.TryGetValue("약칭",                  out var 약칭);
            data.TryGetValue("시료명",                out var 시료명);
            data.TryGetValue("방류허용기준 적용유무", out var 기준유형);

            string 방류기준 = "";
            if (기준Dict != null)
            {
                if (!string.IsNullOrEmpty(기준유형) &&
                    기준Dict.TryGetValue(기준유형.Trim(), out var matched) &&
                    !string.IsNullOrEmpty(matched))
                    방류기준 = matched;
                else
                    방류기준 = 기준Dict.Values.FirstOrDefault() ?? "";
            }

            ws.Cell(rowIdx, 1).Value = 견적번호 ?? "";
            ws.Cell(rowIdx, 2).Value = 약칭     ?? "";
            ws.Cell(rowIdx, 3).Value = 시료명   ?? "";
            ws.Cell(rowIdx, 6).Value = 방류기준;
            rowIdx++;
        }
    }

    // ── 내부 헬퍼 ────────────────────────────────────────────────────────
    private static List<Dictionary<string, string>> GetOrderRows(string 견적번호)
    {
        var result = new List<Dictionary<string, string>>();
        if (!File.Exists(DbPath)) return result;
        try
        {
            using var conn = new SqliteConnection($"Data Source={DbPath}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM \"분석의뢰및결과\" WHERE \"견적번호\" = @no";
            cmd.Parameters.AddWithValue("@no", 견적번호);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < rdr.FieldCount; i++)
                    row[rdr.GetName(i)] = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString() ?? "";
                result.Add(row);
            }
        }
        catch (Exception ex) { Log($"GetOrderRows 오류: {ex.Message}"); }
        return result;
    }

    private static Dictionary<string, Dictionary<string, string>> Get방류기준표Internal()
    {
        var map = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(DbPath)) return map;
        try
        {
            using var conn = new SqliteConnection($"Data Source={DbPath}");
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText =
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='방류기준표'";
            if (Convert.ToInt64(chk.ExecuteScalar()!) == 0) return map;

            var cols = new List<string>();
            using (var pc = conn.CreateCommand())
            {
                pc.CommandText = "PRAGMA table_info(\"방류기준표\")";
                using var pr = pc.ExecuteReader();
                while (pr.Read()) cols.Add(pr.GetString(1));
            }
            if (cols.Count < 2) return map;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM \"방류기준표\"";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var 구분 = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(구분)) continue;
                var inner = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i < cols.Count; i++)
                    inner[cols[i]] = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                map[구분] = inner;
            }
        }
        catch (Exception ex) { Log($"방류기준표 오류: {ex.Message}"); }
        return map;
    }

    private static void Log(string msg)
        => Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] [AnalysisRecord] {msg}");
}
