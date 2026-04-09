using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Shimadzu UV-1800 ASCII 파일 파서 (T-N, T-P, Phenols 등)
/// 형식: "Sample ID","Type","Conc","WL880.0","Comments"
/// </summary>
public static class ShimadzuUvAsciiParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = true };
        string itemAbbr = activeItems.FirstOrDefault() ?? "UVVIS";

        // 인코딩 처리 (EUC-KR 또는 UTF-8)
        string content = ReadFileWithProperEncoding(path);
        var lines = content.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        if (lines.Length < 2)
            throw new ArgumentException("파일이 비어있거나 형식이 올바르지 않습니다.");

        // 헤더 검증
        var header = lines[0].Trim();
        if (!header.Contains("Sample ID") || !header.Contains("Conc") || !header.Contains("WL"))
            throw new ArgumentException("Shimadzu UV-1800 ASCII 형식이 아닙니다.");

        // 표준용액 및 QC 샘플 분석
        var standards = new List<(string name, double conc, double abs)>();
        var ccvs = new List<(string name, double conc, double abs)>();

        // 데이터 행 파싱
        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (string.IsNullOrEmpty(line)) continue;

            var parts = ParseCsvLine(line);
            if (parts.Count < 4) continue;

            var sampleId = parts[0].Trim('"');
            var type = parts[1].Trim('"');
            var concStr = parts[2];
            var wlStr = parts[3];

            if (!double.TryParse(concStr, out var conc)) continue;
            if (!double.TryParse(wlStr, out var absValue)) continue;

            // 블랭크 및 QC 샘플 필터링
            if (IsBlank(sampleId) || IsCcv(sampleId))
            {
                if (IsCcv(sampleId))
                    ccvs.Add((sampleId, conc, absValue));
                continue;
            }

            // 표준용액 처리
            if (IsStandard(sampleId))
            {
                standards.Add((sampleId, conc, absValue));
                continue;
            }

            // 실제 시료 데이터로 ExcelRow 생성
            string cleanedName = CleanSampleName(sampleId);
            if (string.IsNullOrWhiteSpace(cleanedName)) continue;

            var row = new ExcelRow
            {
                시료명 = cleanedName,
                D1 = absValue.ToString("F4"), // 흡광도 → D1 필드 사용
                Fxy = conc.ToString("F4"),    // 농도 → Fxy 필드 사용
                Result = resultFormatter?.Invoke(conc.ToString("F4"), itemAbbr) ?? conc.ToString("F4"),
                Source = SourceType.미분류,
                Status = MatchStatus.대기
            };

            rows.Add(row);
        }

        // 표준곡선 정보 생성 (표준용액 기준)
        BuildCalibrationInfo(docInfo, standards);

        // CCV QC 정보 저장 (기존 필드 재활용)
        if (ccvs.Any())
        {
            var avgCcv = ccvs.Average(x => x.conc);
            docInfo.관련근거 = $"CCV: {avgCcv:F4} mg/L"; // 관련근거 필드에 QC 정보 저장
        }

        return new ParseResult(rows, docInfo, DateTime.Now.ToString("yyyy-MM-dd"), "SHIMADZU_UV1800");
    }

    /// <summary>
    /// 파일 인코딩을 자동 감지하여 읽기
    /// </summary>
    private static string ReadFileWithProperEncoding(string path)
    {
        try
        {
            // UTF-8 시도
            var utf8Content = File.ReadAllText(path, Encoding.UTF8);
            if (!utf8Content.Contains('�')) return utf8Content;

            // EUC-KR 시도 (한글 Windows 기본)
            var euckrEncoding = Encoding.GetEncoding("EUC-KR");
            return File.ReadAllText(path, euckrEncoding);
        }
        catch
        {
            // 최후의 수단: 기본 인코딩
            return File.ReadAllText(path);
        }
    }

    /// <summary>
    /// CSV 라인 파싱 (따옴표 처리)
    /// </summary>
    private static List<string> ParseCsvLine(string line)
    {
        var parts = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];

            if (c == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                parts.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        parts.Add(current.ToString());
        return parts;
    }

    /// <summary>
    /// 블랭크 샘플 판단
    /// </summary>
    private static bool IsBlank(string sampleId) =>
        sampleId.ToUpper().Contains("MBK") ||
        sampleId.ToUpper().Contains("FBK") ||
        sampleId.ToUpper().Contains("BLANK");

    /// <summary>
    /// CCV (Continuing Calibration Verification) 샘플 판단
    /// </summary>
    private static bool IsCcv(string sampleId) =>
        sampleId.ToUpper().Contains("CCV");

    /// <summary>
    /// 표준용액 판단
    /// </summary>
    private static bool IsStandard(string sampleId) =>
        sampleId.Contains("표준용액") ||
        sampleId.Contains("ST") ||
        sampleId.ToUpper().Contains("STD");

    /// <summary>
    /// 시료명 정리 (괄호 제거, 공백 정리)
    /// </summary>
    private static string CleanSampleName(string sampleId)
    {
        string cleaned = sampleId;

        // 괄호 및 괄호 내용 제거
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\([^)]*\)", "");

        // 연속 공백을 하나로 줄이기
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s+", " ");

        return cleaned.Trim();
    }

    /// <summary>
    /// 표준용액 데이터로부터 검량선 정보 생성
    /// </summary>
    private static void BuildCalibrationInfo(ExcelDocInfo docInfo, List<(string name, double conc, double abs)> standards)
    {
        if (standards.Count < 2) return;

        // 농도별 정렬
        var sortedStds = standards.OrderBy(s => s.conc).ToList();

        // 검량선 농도점들 저장
        docInfo.Standard_Points = sortedStds.Take(5).Select(s => s.conc.ToString("F3")).ToArray();
        docInfo.Abs_Values = sortedStds.Take(5).Select(s => s.abs.ToString("F4")).ToArray();

        // 선형 회귀로 기울기/절편 계산
        if (sortedStds.Count >= 2)
        {
            var (slope, intercept, r2) = CalculateLinearRegression(
                sortedStds.Select(s => s.conc).ToArray(),
                sortedStds.Select(s => s.abs).ToArray());

            docInfo.Standard_Slope = slope.ToString("F6");
            docInfo.Standard_Intercept = intercept.ToString("F6");
            docInfo.Abs_R2 = r2.ToString("F5");
        }
    }

    /// <summary>
    /// 선형 회귀 계산 (최소제곱법)
    /// </summary>
    private static (double slope, double intercept, double r2) CalculateLinearRegression(double[] x, double[] y)
    {
        if (x.Length != y.Length || x.Length < 2)
            return (0, 0, 0);

        int n = x.Length;
        double sumX = x.Sum();
        double sumY = y.Sum();
        double sumXY = x.Zip(y, (xi, yi) => xi * yi).Sum();
        double sumXX = x.Sum(xi => xi * xi);
        double sumYY = y.Sum(yi => yi * yi);

        double slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
        double intercept = (sumY - slope * sumX) / n;

        // R² 계산
        double ssRes = y.Zip(x, (yi, xi) => Math.Pow(yi - (slope * xi + intercept), 2)).Sum();
        double ssTot = y.Sum(yi => Math.Pow(yi - sumY / n, 2));
        double r2 = 1 - (ssRes / ssTot);

        return (slope, intercept, r2);
    }
}