using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Agilent Cary-3500 UV CSV 파일 파서 (T-N, T-P, Phenols 등)
/// 구조: METHOD → SAMPLES → Standard → Sample 섹션으로 구성
/// 검량선 회귀식과 표준용액 정보 자동 추출
/// </summary>
public static class AgilentCaryUvCsvParser
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

        var lines = File.ReadAllLines(path);

        // 섹션별 데이터 추출
        var methodInfo = ExtractMethodInfo(lines);
        var standardData = ExtractStandardData(lines);
        var sampleData = ExtractSampleData(lines);
        var calibrationInfo = ExtractCalibrationInfo(lines);

        // 검량선 정보 설정
        SetCalibrationInfo(docInfo, standardData, calibrationInfo);

        // 실제 시료 데이터 처리
        foreach (var sample in sampleData)
        {
            if (IsBlank(sample.Name) || IsCcv(sample.Name)) continue;

            string cleanedName = CleanSampleName(sample.Name);
            if (string.IsNullOrWhiteSpace(cleanedName)) continue;

            var row = new ExcelRow
            {
                시료명 = cleanedName,
                D1 = sample.Abs.ToString("F6"),     // 흡광도
                Fxy = sample.Conc.ToString("F6"),   // 농도
                Result = resultFormatter?.Invoke(sample.Conc.ToString("F6"), itemAbbr) ?? sample.Conc.ToString("F6"),
                Source = SourceType.미분류,
                Status = MatchStatus.대기
            };

            rows.Add(row);
        }

        return new ParseResult(rows, docInfo, DateTime.Now.ToString("yyyy-MM-dd"), "AGILENT_CARY3500");
    }

    /// <summary>
    /// METHOD 섹션에서 기기 설정 정보 추출
    /// </summary>
    private static Dictionary<string, string> ExtractMethodInfo(string[] lines)
    {
        var methodInfo = new Dictionary<string, string>();
        bool inMethodSection = false;

        foreach (var line in lines)
        {
            if (line.Trim() == "METHOD")
            {
                inMethodSection = true;
                continue;
            }

            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("SAMPLES"))
            {
                inMethodSection = false;
            }

            if (inMethodSection && line.Contains(","))
            {
                var parts = line.Split(',', 2);
                if (parts.Length == 2)
                {
                    methodInfo[parts[0].Trim()] = parts[1].Trim();
                }
            }
        }

        return methodInfo;
    }

    /// <summary>
    /// Standard 섹션에서 표준용액 데이터 추출
    /// </summary>
    private static List<(string Name, double Conc, double Abs)> ExtractStandardData(string[] lines)
    {
        var standards = new List<(string, double, double)>();
        bool inStandardSection = false;

        foreach (var line in lines)
        {
            if (line.Trim().StartsWith("Standard,"))
            {
                inStandardSection = true;
                continue;
            }

            if (inStandardSection)
            {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("Sample,"))
                {
                    break;
                }

                if (line.Contains("y =")) continue; // 회귀식 라인 건너뛰기

                var parts = line.Split(',');
                if (parts.Length >= 3 && parts[0].StartsWith("Standard"))
                {
                    if (double.TryParse(parts[1], out var conc) &&
                        double.TryParse(parts[2], out var abs))
                    {
                        standards.Add((parts[0].Trim(), conc, abs));
                    }
                }
            }
        }

        return standards;
    }

    /// <summary>
    /// Sample 섹션에서 시료 데이터 추출
    /// </summary>
    private static List<(string Name, double Conc, double Abs)> ExtractSampleData(string[] lines)
    {
        var samples = new List<(string, double, double)>();
        bool inSampleSection = false;

        foreach (var line in lines)
        {
            if (line.Trim().StartsWith("Sample,"))
            {
                inSampleSection = true;
                continue;
            }

            if (inSampleSection && line.Contains(","))
            {
                var parts = line.Split(',');
                if (parts.Length >= 3)
                {
                    var name = parts[0].Trim().Trim('"');

                    if (double.TryParse(parts[1], out var conc) &&
                        double.TryParse(parts[2], out var abs))
                    {
                        samples.Add((name, conc, abs));
                    }
                }
            }
        }

        return samples;
    }

    /// <summary>
    /// 검량선 회귀식 정보 추출
    /// </summary>
    private static (double Slope, double Intercept, double R2) ExtractCalibrationInfo(string[] lines)
    {
        foreach (var line in lines)
        {
            // "620.00 (nm): ","y = 1.6302 * x + 0.0010, r² = 0.9939" 형식
            if (line.Contains("y =") && line.Contains("r²"))
            {
                return ParseRegressionEquation(line);
            }
        }

        return (0, 0, 0);
    }

    /// <summary>
    /// 회귀식 문자열에서 기울기, 절편, R² 추출
    /// </summary>
    private static (double Slope, double Intercept, double R2) ParseRegressionEquation(string equation)
    {
        try
        {
            // "y = 1.6302 * x + 0.0010, r² = 0.9939" 형식 파싱
            var match = Regex.Match(equation, @"y\s*=\s*([-+]?\d*\.?\d+)\s*\*\s*x\s*([+-])\s*([-+]?\d*\.?\d+).*r²\s*=\s*([-+]?\d*\.?\d+)");

            if (match.Success)
            {
                var slope = double.Parse(match.Groups[1].Value);
                var interceptSign = match.Groups[2].Value == "+" ? 1 : -1;
                var intercept = double.Parse(match.Groups[3].Value) * interceptSign;
                var r2 = double.Parse(match.Groups[4].Value);

                return (slope, intercept, r2);
            }
        }
        catch
        {
            // 파싱 실패 시 기본값 반환
        }

        return (0, 0, 0);
    }

    /// <summary>
    /// 검량선 정보를 ExcelDocInfo에 설정
    /// </summary>
    private static void SetCalibrationInfo(
        ExcelDocInfo docInfo,
        List<(string Name, double Conc, double Abs)> standards,
        (double Slope, double Intercept, double R2) calibration)
    {
        if (standards.Any())
        {
            // 농도 순 정렬하여 표준점들 저장
            var sortedStds = standards.OrderBy(s => s.Conc).Take(5).ToList();

            docInfo.Standard_Points = sortedStds.Select(s => s.Conc.ToString("F3")).ToArray();
            docInfo.Abs_Values = sortedStds.Select(s => s.Abs.ToString("F6")).ToArray();
        }

        // 회귀식 정보 저장
        docInfo.Standard_Slope = calibration.Slope.ToString("F6");
        docInfo.Standard_Intercept = calibration.Intercept.ToString("F6");
        docInfo.Abs_R2 = calibration.R2.ToString("F4");

        // 분석방법 정보 저장
        docInfo.분석방법 = "Agilent Cary-3500 UV-Vis";
        docInfo.결과표시 = "mg/L";
    }

    /// <summary>
    /// 블랭크 샘플 판단
    /// </summary>
    private static bool IsBlank(string sampleName) =>
        sampleName.ToUpper().Contains("MBK") ||
        sampleName.ToUpper().Contains("FBK") ||
        sampleName.ToUpper().Contains("BLANK");

    /// <summary>
    /// CCV 샘플 판단
    /// </summary>
    private static bool IsCcv(string sampleName) =>
        sampleName.ToUpper().Contains("CCV");

    /// <summary>
    /// 시료명 정리
    /// </summary>
    private static string CleanSampleName(string sampleName)
    {
        string cleaned = sampleName;

        // 번호 접미사 제거 (-1, -2 등)
        cleaned = Regex.Replace(cleaned, @"-\d+$", "");

        // 괄호 및 괄호 내용 제거
        cleaned = Regex.Replace(cleaned, @"\([^)]*\)", "");

        // 연속 공백을 하나로 줄이기
        cleaned = Regex.Replace(cleaned, @"\s+", " ");

        return cleaned.Trim();
    }
}