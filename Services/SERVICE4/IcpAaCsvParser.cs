using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Thermo iCE 3000 AA (원자흡광광도계) CSV 파일 파서
/// 단일 원소 (주로 Hg) Vapour 모드 결과 파일
///
/// CSV 구조:
///   "ANALYSIS INFORMATION" 행 → 다음 행에 날짜(col4), 기기명(col5)
///   "RESULTS FOR ELEMENT","원소명" → 원소 설정
///   결과 헤더 행 → 건너뜀
///   단위 행(Abs/%) → 건너뜀
///   데이터 행: col1=시료명, col2=RESULT TYPE, col5=FLAGS, col6=CONC.(µg/L)
///     - RESULT TYPE == "Mean" 인 행만 사용
///     - 건너뛸 시료: Blank, Standard 1~6
///     - FLAG "C" 또는 음수 농도 = 불검출 → Result = ""
///     - 단위 µg/L → mg/L 변환 (/1000)
/// </summary>
public static class IcpAaCsvParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    private static readonly HashSet<string> SkipSampleIds =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "Blank", "Standard 1", "Standard 2", "Standard 3",
            "Standard 4", "Standard 5", "Standard 6"
        };

    private static bool IsQcSample(string name) =>
        name.StartsWith("MBK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("FBK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("CCV", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("-CCV", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("BK", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("BLANK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("DW", StringComparison.OrdinalIgnoreCase);

    /// <summary>Thermo iCE 3000 AA CSV 파일 형식인지 확인</summary>
    public static bool IsAaCsv(string path)
    {
        try
        {
            var lines = File.ReadLines(path).Take(5).ToArray();
            return lines.Any(l => l.Contains("ANALYSIS INFORMATION")) &&
                   lines.Any(l => l.Contains("iCE") || l.Contains("AA"));
        }
        catch { return false; }
    }

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = false };

        string analysisDate   = DateTime.Now.ToString("yyyy-MM-dd");
        string instrumentName = "iCE 3000 AA";
        string elementName    = activeItems.FirstOrDefault() ?? "Hg";

        bool   inResults             = false;
        bool   skipNextAsHeader      = false;
        bool   skipNextAsUnits       = false;
        bool   nextIsAnalysisInfoRow = false;

        try
        {
            foreach (var rawLine in File.ReadLines(path))
            {
                var fields = SplitCsvLine(rawLine);
                if (fields.Length == 0) continue;

                // ANALYSIS INFORMATION 다음 행에서 날짜·기기명 추출
                if (nextIsAnalysisInfoRow)
                {
                    nextIsAnalysisInfoRow = false;
                    if (fields.Length > 4 && !string.IsNullOrWhiteSpace(fields[4]))
                        analysisDate = fields[4].Trim();
                    if (fields.Length > 5 && !string.IsNullOrWhiteSpace(fields[5]))
                        instrumentName = fields[5].Trim();
                    continue;
                }

                if (fields[0] == "ANALYSIS INFORMATION")
                {
                    nextIsAnalysisInfoRow = true;
                    continue;
                }

                // RESULTS FOR ELEMENT 섹션 시작
                if (fields[0] == "RESULTS FOR ELEMENT")
                {
                    if (fields.Length > 1 && !string.IsNullOrWhiteSpace(fields[1]))
                        elementName = fields[1].Trim();
                    inResults       = true;
                    skipNextAsHeader = true;
                    continue;
                }

                // 결과 헤더 행 ("","SAMPLE ID","RESULT TYPE",...) 건너뜀
                if (skipNextAsHeader)
                {
                    skipNextAsHeader = false;
                    skipNextAsUnits  = true;
                    continue;
                }

                // 단위 행 ("","","","Abs","%",...) 건너뜀
                if (skipNextAsUnits)
                {
                    skipNextAsUnits = false;
                    continue;
                }

                if (!inResults) continue;

                // 데이터 행: col0="" col1=시료명 col2=RESULT TYPE col5=FLAGS col6=CONC.
                if (fields.Length < 7) continue;
                if (fields[0] != "") continue;  // 섹션 구분자 행

                string sampleId  = fields.Length > 1 ? fields[1].Trim() : "";
                string resultType = fields.Length > 2 ? fields[2].Trim() : "";
                string flags      = fields.Length > 5 ? fields[5].Trim() : "";
                string concRaw    = fields.Length > 6 ? fields[6].Trim() : "";

                if (!resultType.Equals("Mean", StringComparison.OrdinalIgnoreCase)) continue;
                if (string.IsNullOrWhiteSpace(sampleId)) continue;
                if (SkipSampleIds.Contains(sampleId)) continue;

                // 농도 계산: µg/L → mg/L, 음수 또는 FLAG=C → 불검출
                string result = "";
                if (!string.IsNullOrWhiteSpace(concRaw) &&
                    double.TryParse(concRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out double concUgL))
                {
                    if (flags.Equals("C", StringComparison.OrdinalIgnoreCase) || concUgL < 0)
                        result = "";  // 불검출
                    else
                        result = (concUgL / 1000.0).ToString("G6", CultureInfo.InvariantCulture); // µg/L → mg/L
                }

                rows.Add(new ExcelRow
                {
                    시료명      = sampleId,
                    CompoundName = elementName,
                    Result      = result,
                    Source      = SourceType.미분류,
                    Status      = MatchStatus.대기,
                    IsControl   = IsQcSample(sampleId),
                });
            }
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"AA CSV 파싱 오류: {ex.Message}", ex);
        }

        docInfo.분석방법 = instrumentName;
        docInfo.결과표시 = "mg/L";

        return new ParseResult(rows, docInfo, analysisDate, "ICP_AA_CSV");
    }

    // ── CSV 한 줄 파싱 (전체 쌍따옴표 감싸기 형식) ────────────────────────────
    private static string[] SplitCsvLine(string line)
    {
        if (string.IsNullOrEmpty(line)) return Array.Empty<string>();
        var parts = line.Split(new[] { "\",\"" }, StringSplitOptions.None);
        for (int i = 0; i < parts.Length; i++)
            parts[i] = parts[i].Trim('"');
        return parts;
    }
}
