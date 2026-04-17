using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 지에스칼텍스 전용 GS 발송양식 Excel 생성
/// VBA 매크로 'GS조회' / 'GS보냄' 로직 포팅
/// </summary>
public static class GsReportService
{
    private static readonly string TemplatePath =
        Path.Combine("Data", "Templates", "GS_발송양식_template.xlsx");

    /// <summary>
    /// 시료명에서 배출구번호 추출 (점 앞까지)
    /// 예: "#W1.(제1폐수처리장/처리수)" → "#W1"
    /// </summary>
    public static string ExtractWellNo(string 시료명)
    {
        if (string.IsNullOrEmpty(시료명)) return "";
        int dotIdx = 시료명.IndexOf('.');
        return dotIdx > 0 ? 시료명.Substring(0, dotIdx).Trim() : 시료명.Trim();
    }

    /// <summary>
    /// GS_CODE 테이블에서 오염물질명(한글)으로 코드, 공식명칭 조회
    /// </summary>
    private static Dictionary<string, (int code, string name)> LoadCodeMap()
    {
        var map = new Dictionary<string, (int, string)>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "GS_CODE")) return map;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT 코드, 오염물질명 FROM `GS_CODE`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var code = r.GetInt32(0);
                var name = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                if (!string.IsNullOrEmpty(name))
                    map[name] = (code, name);
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[GS] LoadCodeMap 실패: {ex.Message}"); }
        return map;
    }

    /// <summary>
    /// GS_CODE 테이블에서 특정 배출구의 허가기준값 맵 로드
    /// wellKey: "#W23" 형태. 실패 시 '#' 제거 후 재시도
    /// </summary>
    private static Dictionary<string, string> LoadStandardMap(string wellKey)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var cols = DbConnectionFactory.GetColumnNames(conn, "GS_CODE");

            // 컬럼 매칭: 정확 → # 제거 → 대소문자 무시
            string? matchCol = cols.FirstOrDefault(c => c == wellKey);
            if (matchCol == null && wellKey.StartsWith("#"))
                matchCol = cols.FirstOrDefault(c => c == wellKey.Substring(1));
            if (matchCol == null)
            {
                var norm = (string s) => s.Replace("#", "").Trim().ToLower();
                matchCol = cols.FirstOrDefault(c => norm(c) == norm(wellKey));
            }
            if (matchCol == null) return map;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `오염물질명`, `{matchCol}` FROM `GS_CODE`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var key = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                var val = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(key)) map[key] = val;
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[GS] LoadStandardMap 실패: {ex.Message}"); }
        return map;
    }

    /// <summary>
    /// GS_CODE.오염물질명과 분석결과 항목명 매칭
    /// 정확 일치 → 공백/특수문자 제거 후 재시도
    /// </summary>
    private static string? MatchCodeName(string 항목명, IEnumerable<string> codeNames)
    {
        var match = codeNames.FirstOrDefault(n => n == 항목명);
        if (match != null) return match;
        var norm = (string s) => s.Replace(" ", "").Replace("(", "").Replace(")", "")
                                  .Replace(",", "").Replace("-", "").ToLower();
        var target = norm(항목명);
        return codeNames.FirstOrDefault(n => norm(n) == target);
    }

    /// <summary>
    /// SampleRequest 리스트로 GS 발송양식 Excel 파일 생성
    /// 반환: 생성된 파일 경로
    /// </summary>
    public static string GenerateReport(List<SampleRequest> samples, Dictionary<string, AnalysisItem>? meta = null)
    {
        if (samples == null || samples.Count == 0)
            throw new InvalidOperationException("시료 없음");

        if (!File.Exists(TemplatePath))
            throw new FileNotFoundException($"템플릿 없음: {TemplatePath}");

        var codeMap = LoadCodeMap();
        var codeNames = codeMap.Keys.ToList();

        using var wb = new XLWorkbook(TemplatePath);
        var ws = wb.Worksheet("GS 발송양식");

        int row = 3;
        foreach (var sample in samples)
        {
            var wellNo = ExtractWellNo(sample.시료명);
            var stdMap = LoadStandardMap(wellNo);

            foreach (var (항목, 값) in sample.분석결과)
            {
                if (string.IsNullOrWhiteSpace(값)) continue;

                var matchedName = MatchCodeName(항목, codeNames);
                if (matchedName == null) continue;  // GS_CODE에 없는 항목은 스킵

                var (code, name) = codeMap[matchedName];

                ws.Cell(row, 1).Value = sample.분석종료일;  // 측정일시
                ws.Cell(row, 2).Value = wellNo;             // 배출구번호
                // 3: 사업장관리번호 (비워둠)
                ws.Cell(row, 4).Value = code;               // 오염물질코드
                ws.Cell(row, 5).Value = name;               // 오염물질명
                ws.Cell(row, 9).Value = 값;                  // 방류수 농도
                if (stdMap.TryGetValue(matchedName, out var stdVal))
                    ws.Cell(row, 11).Value = stdVal;         // 허가배출기준

                // 검사방법(기기명) — 분석정보의 instrument + ES 코드 조합
                if (meta != null && meta.TryGetValue(항목, out var m))
                {
                    var line = (m.instrument ?? "");
                    if (!string.IsNullOrEmpty(m.ES)) line = $"{line} / {m.ES}".Trim(' ', '/');
                    ws.Cell(row, 12).Value = line;
                }

                row++;
            }
        }

        // 결과 저장 경로
        var outDir = Path.Combine("Data", "Exports");
        Directory.CreateDirectory(outDir);
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var outPath = Path.Combine(outDir, $"GS_발송양식_{ts}.xlsx");
        wb.SaveAs(outPath);
        return outPath;
    }
}
