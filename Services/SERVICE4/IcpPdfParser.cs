using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// ICP (Inductively Coupled Plasma) PDF 분석결과 파서
/// 원소별 농도 결과 (Cu, Zn, Pb, Cd, Cr, Ni 등)를 추출하여 처리
/// </summary>
public static class IcpPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    // 일반적인 ICP 분석 원소들
    private static readonly string[] CommonElements =
    {
        "Cu", "Zn", "Pb", "Cd", "Cr", "Ni", "Fe", "Mn", "Al", "As", "Hg", "Se", "Co"
    };

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = false };
        string itemAbbr = activeItems.FirstOrDefault() ?? "ICP";

        try
        {
            // PDF 텍스트 추출
            string pdfText = ExtractPdfText(path);

            // 시료 데이터 추출
            var sampleResults = ExtractSampleResults(pdfText);

            // 검출한계 정보 추출
            var detectionLimits = ExtractDetectionLimits(pdfText);

            // 분석 조건 정보 추출
            ExtractAnalysisConditions(pdfText, docInfo);

            // 각 시료별 ExcelRow 생성
            foreach (var sample in sampleResults)
            {
                string cleanedName = CleanSampleName(sample.SampleName);
                if (string.IsNullOrWhiteSpace(cleanedName)) continue;

                // 원소별로 개별 행 생성 또는 대표값 사용
                string representativeElement = GetRepresentativeElement(sample.Elements);
                var elementData = sample.Elements.FirstOrDefault(e => e.Element == representativeElement);

                if (elementData != null)
                {
                    var row = new ExcelRow
                    {
                        시료명 = cleanedName,
                        D1 = elementData.Concentration.ToString("F6"),  // 농도
                        D2 = detectionLimits.ContainsKey(elementData.Element)
                            ? detectionLimits[elementData.Element].ToString("F6")
                            : "",  // 검출한계
                        Fxy = elementData.Element,  // 원소명
                        Result = resultFormatter?.Invoke(elementData.Concentration.ToString("F6"), itemAbbr)
                                ?? elementData.Concentration.ToString("F6"),
                        Source = SourceType.미분류,
                        Status = MatchStatus.대기
                    };

                    rows.Add(row);
                }
            }

            // 문서 정보 설정
            SetDocumentInfo(docInfo, detectionLimits);

            return new ParseResult(rows, docInfo, ExtractAnalysisDate(pdfText), "ICP_PDF");
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"ICP PDF 파싱 오류: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// PDF에서 텍스트 추출 (향후 PDF 라이브러리 통합 예정)
    /// </summary>
    private static string ExtractPdfText(string path)
    {
        // TODO: PDF 라이브러리 (iText, PDFSharp 등)를 사용하여 텍스트 추출
        // 현재는 시뮬레이션된 데이터를 반환

        var fileName = Path.GetFileName(path);

        // 임시 시뮬레이션 데이터 (실제 구현 시 PDF 텍스트 추출로 교체)
        return @"
        ICP Analysis Report
        Analysis Date: 2026-04-09
        Instrument: Agilent 5110 ICP-OES

        Sample Results (mg/L):
        Sample ID    Cu      Zn      Pb      Cd      Cr      Ni
        여수-01     0.015   0.089   0.005   0.001   0.025   0.012
        4단계-01    0.008   0.045   0.002   <0.001  0.018   0.008
        세풍-01     0.012   0.067   0.003   0.001   0.022   0.010

        Detection Limits (mg/L):
        Cu: 0.001, Zn: 0.005, Pb: 0.001, Cd: 0.001, Cr: 0.002, Ni: 0.001
        ";
    }

    /// <summary>
    /// PDF 텍스트에서 시료별 원소 농도 결과 추출
    /// </summary>
    private static List<SampleResult> ExtractSampleResults(string pdfText)
    {
        var results = new List<SampleResult>();
        var lines = pdfText.Split('\n');

        bool inDataSection = false;
        string[] headers = null;

        foreach (var line in lines)
        {
            var cleanLine = line.Trim();
            if (string.IsNullOrEmpty(cleanLine)) continue;

            // 헤더 행 감지
            if (cleanLine.Contains("Sample") && CommonElements.Any(e => cleanLine.Contains(e)))
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
                    var sampleResult = new SampleResult { SampleName = parts[0] };

                    for (int i = 1; i < Math.Min(parts.Length, headers.Length); i++)
                    {
                        if (i < headers.Length && CommonElements.Contains(headers[i]))
                        {
                            if (TryParseConcentration(parts[i], out double conc))
                            {
                                sampleResult.Elements.Add(new ElementResult
                                {
                                    Element = headers[i],
                                    Concentration = conc
                                });
                            }
                        }
                    }

                    if (sampleResult.Elements.Any())
                    {
                        results.Add(sampleResult);
                    }
                }
            }
        }

        return results;
    }

    /// <summary>
    /// 농도값 파싱 (&lt;0.001 등의 형식 처리)
    /// </summary>
    private static bool TryParseConcentration(string value, out double concentration)
    {
        concentration = 0;

        if (string.IsNullOrWhiteSpace(value)) return false;

        // "<0.001" 형식 처리
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
    /// 검출한계 정보 추출
    /// </summary>
    private static Dictionary<string, double> ExtractDetectionLimits(string pdfText)
    {
        var limits = new Dictionary<string, double>();
        var lines = pdfText.Split('\n');

        foreach (var line in lines)
        {
            if (line.Contains("Detection Limit") || line.Contains("검출한계"))
            {
                // "Cu: 0.001, Zn: 0.005" 형식 파싱
                var matches = Regex.Matches(line, @"([A-Z][a-z]?):\s*([\d.]+)");
                foreach (Match match in matches)
                {
                    var element = match.Groups[1].Value;
                    if (double.TryParse(match.Groups[2].Value, out var limit))
                    {
                        limits[element] = limit;
                    }
                }
            }
        }

        return limits;
    }

    /// <summary>
    /// 분석 조건 정보 추출
    /// </summary>
    private static void ExtractAnalysisConditions(string pdfText, ExcelDocInfo docInfo)
    {
        var lines = pdfText.Split('\n');

        foreach (var line in lines)
        {
            if (line.Contains("Instrument:"))
            {
                docInfo.분석방법 = line.Substring(line.IndexOf("Instrument:") + 11).Trim();
            }
            else if (line.Contains("분석기기:"))
            {
                docInfo.분석방법 = line.Substring(line.IndexOf("분석기기:") + 5).Trim();
            }
        }

        if (string.IsNullOrEmpty(docInfo.분석방법))
        {
            docInfo.분석방법 = "ICP-OES";
        }

        docInfo.결과표시 = "mg/L";
    }

    /// <summary>
    /// 분석일자 추출
    /// </summary>
    private static string ExtractAnalysisDate(string pdfText)
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
    /// 대표 원소 선택 (Cu > Zn > Cr 순서)
    /// </summary>
    private static string GetRepresentativeElement(List<ElementResult> elements)
    {
        var priority = new[] { "Cu", "Zn", "Cr", "Ni", "Pb", "Cd" };

        foreach (var element in priority)
        {
            if (elements.Any(e => e.Element == element))
                return element;
        }

        return elements.FirstOrDefault()?.Element ?? "Unknown";
    }

    /// <summary>
    /// 유효한 시료명인지 확인
    /// </summary>
    private static bool IsValidSampleName(string name)
    {
        return !string.IsNullOrWhiteSpace(name) &&
               !name.Equals("Sample", StringComparison.OrdinalIgnoreCase) &&
               !CommonElements.Contains(name);
    }

    /// <summary>
    /// 시료명 정리
    /// </summary>
    private static string CleanSampleName(string sampleName)
    {
        return sampleName.Trim().Replace("-", " ");
    }

    /// <summary>
    /// 문서 정보 설정
    /// </summary>
    private static void SetDocumentInfo(ExcelDocInfo docInfo, Dictionary<string, double> detectionLimits)
    {
        if (detectionLimits.Any())
        {
            var limitInfo = string.Join(", ", detectionLimits.Select(kv => $"{kv.Key}: {kv.Value:F3}"));
            docInfo.관련근거 = $"검출한계 (mg/L): {limitInfo}";
        }
    }

    // 데이터 모델 클래스들
    private class SampleResult
    {
        public string SampleName { get; set; } = "";
        public List<ElementResult> Elements { get; set; } = new();
    }

    private class ElementResult
    {
        public string Element { get; set; } = "";
        public double Concentration { get; set; }
    }
}