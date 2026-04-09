using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// LCMS 과불화화합물(PFAS) PDF 분석결과 파서
/// PFOA, PFOS, PFBS 등 과불화화합물 3종 분석 결과 추출
/// </summary>
public static class LcmsPfasPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    // 일반적인 과불화화합물 항목들
    private static readonly string[] CommonPfasCompounds =
    {
        "PFOA", "PFOS", "PFBS", "PFNA", "PFDA", "PFHxS", "PFBA", "PFHpA", "PFUnA", "PFDoA"
    };

    // 한글-영문 매핑
    private static readonly Dictionary<string, string> PfasKoreanMap = new()
    {
        ["퍼플루오로옥탄산"] = "PFOA",
        ["퍼플루오로옥탄술폰산"] = "PFOS",
        ["퍼플루오로부탄술폰산"] = "PFBS",
        ["과불화옥탄산"] = "PFOA",
        ["과불화옥탄술폰산"] = "PFOS",
        ["과불화부탄술폰산"] = "PFBS"
    };

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = false };
        string itemAbbr = activeItems.FirstOrDefault() ?? "PFAS";

        try
        {
            // PDF 텍스트 추출
            string pdfText = ExtractPdfText(path);

            // 시료 데이터 추출
            var sampleResults = ExtractPfasSampleResults(pdfText);

            // 검출한계 정보 추출
            var detectionLimits = ExtractPfasDetectionLimits(pdfText);

            // 분석 조건 정보 추출
            ExtractPfasAnalysisConditions(pdfText, docInfo);

            // 각 시료별 ExcelRow 생성
            foreach (var sample in sampleResults)
            {
                string cleanedName = CleanSampleName(sample.SampleName);
                if (string.IsNullOrWhiteSpace(cleanedName)) continue;

                // 화합물별로 개별 행 생성 또는 대표값 사용
                string representativeCompound = GetRepresentativePfasCompound(sample.Compounds);
                var compoundData = sample.Compounds.FirstOrDefault(c => c.Compound == representativeCompound);

                if (compoundData != null)
                {
                    var row = new ExcelRow
                    {
                        시료명 = cleanedName,
                        D1 = compoundData.Concentration.ToString("F6"),  // 농도
                        D2 = detectionLimits.ContainsKey(compoundData.Compound)
                            ? detectionLimits[compoundData.Compound].ToString("F6")
                            : "",  // 검출한계
                        Fxy = compoundData.Compound,  // 화합물명
                        Result = resultFormatter?.Invoke(compoundData.Concentration.ToString("F6"), itemAbbr)
                                ?? compoundData.Concentration.ToString("F6"),
                        Source = SourceType.미분류,
                        Status = MatchStatus.대기
                    };

                    rows.Add(row);
                }
            }

            // 문서 정보 설정
            SetPfasDocumentInfo(docInfo, detectionLimits);

            return new ParseResult(rows, docInfo, ExtractPfasAnalysisDate(pdfText), "LCMS_PFAS");
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"LCMS PFAS PDF 파싱 오류: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// PDF에서 텍스트 추출 (향후 PDF 라이브러리 통합 예정)
    /// </summary>
    private static string ExtractPdfText(string path)
    {
        var fileName = Path.GetFileName(path);

        // 임시 시뮬레이션 데이터 (실제 구현 시 PDF 텍스트 추출로 교체)
        return @"
        LCMS Analysis Report - PFAS Compounds
        Analysis Date: 2026-02-26
        Instrument: Agilent 6470 Triple Quad LC/MS

        Sample Results (ng/L):
        Sample ID    PFOA    PFOS    PFBS    PFNA
        여수-01     2.45    1.28    0.85    <0.5
        4단계-01    1.15    0.92    <0.5    <0.5
        세풍-01     3.22    2.10    1.45    0.68

        Detection Limits (ng/L):
        PFOA: 0.5, PFOS: 0.5, PFBS: 0.5, PFNA: 0.5

        Method: EPA 537.1 Modified
        Column: C18 (2.1 x 100mm, 3.5μm)
        Mobile Phase: 2mM Ammonium acetate / Methanol
        ";
    }

    /// <summary>
    /// PDF 텍스트에서 시료별 PFAS 농도 결과 추출
    /// </summary>
    private static List<PfasSampleResult> ExtractPfasSampleResults(string pdfText)
    {
        var results = new List<PfasSampleResult>();
        var lines = pdfText.Split('\n');

        bool inDataSection = false;
        string[] headers = null;

        foreach (var line in lines)
        {
            var cleanLine = line.Trim();
            if (string.IsNullOrEmpty(cleanLine)) continue;

            // 헤더 행 감지
            if (cleanLine.Contains("Sample") && CommonPfasCompounds.Any(c => cleanLine.Contains(c)))
            {
                headers = cleanLine.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                inDataSection = true;
                continue;
            }

            // 데이터 행 처리
            if (inDataSection && headers != null)
            {
                var parts = cleanLine.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 2 && IsValidSampleName(parts[0]))
                {
                    var sampleResult = new PfasSampleResult { SampleName = parts[0] };

                    for (int i = 1; i < Math.Min(parts.Length, headers.Length); i++)
                    {
                        if (i < headers.Length && CommonPfasCompounds.Contains(headers[i]))
                        {
                            if (TryParsePfasConcentration(parts[i], out double conc))
                            {
                                sampleResult.Compounds.Add(new PfasCompoundResult
                                {
                                    Compound = headers[i],
                                    Concentration = conc
                                });
                            }
                        }
                    }

                    if (sampleResult.Compounds.Any())
                    {
                        results.Add(sampleResult);
                    }
                }
            }
        }

        return results;
    }

    /// <summary>
    /// PFAS 농도값 파싱 (&lt;0.5 등의 형식 처리)
    /// </summary>
    private static bool TryParsePfasConcentration(string value, out double concentration)
    {
        concentration = 0;

        if (string.IsNullOrWhiteSpace(value)) return false;

        // "<0.5" 형식 처리
        if (value.StartsWith("<"))
        {
            var numPart = value.Substring(1);
            if (double.TryParse(numPart, out concentration))
            {
                concentration = concentration / 2; // 검출한계 미만은 절반값으로 처리
                return true;
            }
        }

        return double.TryParse(value, out concentration);
    }

    /// <summary>
    /// PFAS 검출한계 정보 추출
    /// </summary>
    private static Dictionary<string, double> ExtractPfasDetectionLimits(string pdfText)
    {
        var limits = new Dictionary<string, double>();
        var lines = pdfText.Split('\n');

        foreach (var line in lines)
        {
            if (line.Contains("Detection Limit") || line.Contains("검출한계"))
            {
                // "PFOA: 0.5, PFOS: 0.5" 형식 파싱
                var matches = Regex.Matches(line, @"(PF[A-Z]+):\s*([\d.]+)");
                foreach (Match match in matches)
                {
                    var compound = match.Groups[1].Value;
                    if (double.TryParse(match.Groups[2].Value, out var limit))
                    {
                        limits[compound] = limit;
                    }
                }
            }
        }

        return limits;
    }

    /// <summary>
    /// PFAS 분석 조건 정보 추출
    /// </summary>
    private static void ExtractPfasAnalysisConditions(string pdfText, ExcelDocInfo docInfo)
    {
        var lines = pdfText.Split('\n');

        foreach (var line in lines)
        {
            if (line.Contains("Instrument:") || line.Contains("분석기기:"))
            {
                var start = line.IndexOf("Instrument:");
                if (start == -1) start = line.IndexOf("분석기기:");
                docInfo.분석방법 = line.Substring(start + 11).Trim();
            }
            else if (line.Contains("Method:"))
            {
                docInfo.관련근거 = line.Substring(line.IndexOf("Method:") + 7).Trim();
            }
        }

        if (string.IsNullOrEmpty(docInfo.분석방법))
        {
            docInfo.분석방법 = "LC-MS/MS";
        }

        if (string.IsNullOrEmpty(docInfo.관련근거))
        {
            docInfo.관련근거 = "EPA 537.1 Modified";
        }

        docInfo.결과표시 = "ng/L";
    }

    /// <summary>
    /// PFAS 분석일자 추출
    /// </summary>
    private static string ExtractPfasAnalysisDate(string pdfText)
    {
        var datePatterns = new[]
        {
            @"Analysis Date:\s*(\d{4}-\d{2}-\d{2})",
            @"분석일자:\s*(\d{4}-\d{2}-\d{2})",
            @"(\d{4}-\d{2}-\d{2})"
        };

        foreach (var pattern in datePatterns)
        {
            var match = Regex.Match(pdfText, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }

        return DateTime.Now.ToString("yyyy-MM-dd");
    }

    /// <summary>
    /// 대표 PFAS 화합물 선택 (PFOA > PFOS > PFBS 순서)
    /// </summary>
    private static string GetRepresentativePfasCompound(List<PfasCompoundResult> compounds)
    {
        var priority = new[] { "PFOA", "PFOS", "PFBS", "PFNA", "PFDA" };

        foreach (var compound in priority)
        {
            if (compounds.Any(c => c.Compound == compound))
                return compound;
        }

        return compounds.FirstOrDefault()?.Compound ?? "Unknown";
    }

    /// <summary>
    /// 유효한 시료명인지 확인
    /// </summary>
    private static bool IsValidSampleName(string name)
    {
        return !string.IsNullOrWhiteSpace(name) &&
               !name.Equals("Sample", StringComparison.OrdinalIgnoreCase) &&
               !CommonPfasCompounds.Contains(name);
    }

    /// <summary>
    /// 시료명 정리
    /// </summary>
    private static string CleanSampleName(string sampleName)
    {
        return sampleName.Trim().Replace("-", " ");
    }

    /// <summary>
    /// PFAS 문서 정보 설정
    /// </summary>
    private static void SetPfasDocumentInfo(ExcelDocInfo docInfo, Dictionary<string, double> detectionLimits)
    {
        if (detectionLimits.Any())
        {
            var limitInfo = string.Join(", ", detectionLimits.Select(kv => $"{kv.Key}: {kv.Value:F1}"));
            docInfo.관련근거 = $"검출한계 (ng/L): {limitInfo}";
        }
    }

    // 데이터 모델 클래스들
    private class PfasSampleResult
    {
        public string SampleName { get; set; } = "";
        public List<PfasCompoundResult> Compounds { get; set; } = new();
    }

    private class PfasCompoundResult
    {
        public string Compound { get; set; } = "";
        public double Concentration { get; set; }
    }
}