using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Thermo iCE 3000 AA — SOLAAR Data Station V11.03 "SOLAAR AA Report" PDF 파서
/// 단일 원소(주로 Hg, 수은) Vapour 모드 결과
///
/// PDF 구조:
///   Page 1: General Parameters / Spectrometer Parameters / Calibration Parameters
///   Page 2: "Solution Results - Hg" 섹션에 결과 테이블
///     컬럼: Sample ID | Signal(Abs) | Rsd(%) | Conc(µg/L) | Corrected Conc(µg/L)
///     Sample ID 앞에 원소기호 prefix 붙음 (예: "HgBlank", "HgStandard 1", "HgMBK", "Hg1")
///
///   건너뜀: Blank, Standard 1~5
///   정도관리(IsControl): MBK, FBK*, CCV*, BK*, DW*
///   불검출: FLAG "C" 또는 음수 농도 → Result = ""
///   단위 µg/L → mg/L 변환 (/1000)
/// </summary>
public static class AasHgPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    private static readonly Regex DateRx = new(
        @"(\d{4}-\d{2}-\d{2})", RegexOptions.Compiled);

    // 결과 행 패턴:
    //   <SampleId> <Signal> <Rsd> <Conc>[ C] [<CorrectedConc>[ C]]
    //   Signal: -?\d*\.\d+
    //   Rsd: -?\d+(\.\d+)?  또는  >99
    //   Conc: -?\d+(\.\d+)?  + 선택적 " C" 플래그
    private static readonly Regex DataRowRx = new(
        @"^(?<id>\S.*?)\s+(?<signal>-?\d*\.\d+)\s+(?<rsd>-?\d+(?:\.\d+)?|>99)\s+(?<conc>-?\d+(?:\.\d+)?)(?:\s+(?<flag1>[A-Z]))?(?:\s+(?<corr>-?\d+(?:\.\d+)?)(?:\s+(?<flag2>[A-Z]))?)?\s*$",
        RegexOptions.Compiled);

    private static readonly HashSet<string> SkipExact =
        new(StringComparer.OrdinalIgnoreCase) { "Blank" };

    private static readonly Regex SkipStandardRx = new(
        @"^Standard\s+\d+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static bool IsQcSample(string name) =>
        name.StartsWith("MBK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("FBK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("CCV", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("-CCV", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("BK", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("BLANK", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("DW", StringComparison.OrdinalIgnoreCase);

    /// <summary>SOLAAR AA Report PDF 여부 확인</summary>
    public static bool IsAaPdf(string path)
    {
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            if (doc.NumberOfPages == 0) return false;
            string text = string.Join(" ", doc.GetPage(1).GetWords().Select(w => w.Text));
            return text.Contains("SOLAAR", StringComparison.OrdinalIgnoreCase) ||
                   text.Contains("iCE 3000 AA", StringComparison.OrdinalIgnoreCase);
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
        string instrumentName = "Thermo iCE 3000 AA";
        string elementName    = activeItems.FirstOrDefault() ?? "Hg";
        bool   dateFound      = false;
        bool   inResults      = false;

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);

            foreach (var page in doc.GetPages())
            {
                foreach (var line in ExtractLines(page))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // 분석일자 (최초 1회): "2026-03-30 1:18:43" 같은 라인에서 첫 매치
                    if (!dateFound)
                    {
                        var m = DateRx.Match(line);
                        if (m.Success && line.Contains("Record created", StringComparison.OrdinalIgnoreCase))
                        {
                            analysisDate = m.Groups[1].Value;
                            dateFound = true;
                        }
                    }

                    // 원소명 추출: "Element: Hg"
                    var elemMatch = Regex.Match(line, @"Element\s*:?\s*([A-Z][a-z]?)\b");
                    if (elemMatch.Success)
                    {
                        var e = elemMatch.Groups[1].Value;
                        if (e.Length <= 2 && char.IsUpper(e[0])) elementName = e;
                    }

                    // 결과 섹션 진입
                    if (line.Contains("Solution Results", StringComparison.OrdinalIgnoreCase))
                    {
                        inResults = true;
                        continue;
                    }

                    // 다음 헤더로 빠져나가면 종료
                    if (inResults && (line.Contains("Page", StringComparison.OrdinalIgnoreCase) ||
                                      line.Contains("SOLAAR Data Station", StringComparison.OrdinalIgnoreCase)))
                    {
                        // 페이지 푸터/헤더는 무시 — 다음 페이지에서 다시 결과 이어질 수 있으므로 종료 안 함
                        continue;
                    }

                    if (!inResults) continue;

                    // 헤더 행 (Sample ID/Signal/Rsd/Conc, Abs/%/μg/L) 스킵
                    if (line.Contains("Sample ID", StringComparison.OrdinalIgnoreCase)) continue;
                    if (line.StartsWith("Abs", StringComparison.OrdinalIgnoreCase)) continue;
                    if (line.Contains("μg/L") || line.Contains("ug/L") || line.Contains("μg/L")) continue;

                    // 데이터 행 매칭
                    var rm = DataRowRx.Match(line);
                    if (!rm.Success) continue;

                    string rawId = rm.Groups["id"].Value.Trim();

                    // 원소 prefix 제거 (예: "HgMBK" → "MBK", "Hg1" → "1")
                    string sampleId = rawId;
                    if (!string.IsNullOrEmpty(elementName) &&
                        sampleId.StartsWith(elementName, StringComparison.Ordinal))
                    {
                        sampleId = sampleId.Substring(elementName.Length).Trim();
                    }
                    if (string.IsNullOrWhiteSpace(sampleId)) continue;

                    // 표준/블랭크 스킵
                    if (SkipExact.Contains(sampleId)) continue;
                    if (SkipStandardRx.IsMatch(sampleId)) continue;

                    string concRaw  = rm.Groups["conc"].Value;
                    string flag     = rm.Groups["flag1"].Success ? rm.Groups["flag1"].Value
                                    : rm.Groups["flag2"].Success ? rm.Groups["flag2"].Value : "";

                    string result = "";
                    if (!string.IsNullOrWhiteSpace(concRaw) &&
                        double.TryParse(concRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out double concUgL))
                    {
                        if (flag.Equals("C", StringComparison.OrdinalIgnoreCase) || concUgL < 0)
                            result = "";
                        else
                            result = (concUgL / 1000.0).ToString("G6", CultureInfo.InvariantCulture);
                    }

                    rows.Add(new ExcelRow
                    {
                        시료명       = sampleId,
                        CompoundName = elementName,
                        Result       = result,
                        Source       = SourceType.미분류,
                        Status       = MatchStatus.대기,
                        IsControl    = IsQcSample(sampleId),
                    });
                }
            }
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"AA PDF 파싱 오류: {ex.Message}", ex);
        }

        docInfo.분석방법 = instrumentName;
        docInfo.결과표시 = "mg/L";

        return new ParseResult(rows, docInfo, analysisDate, "AA_HG_PDF");
    }

    // PdfPig 단어들을 Y좌표 ±5pt 묶음으로 라인 복원
    private static IEnumerable<string> ExtractLines(UglyToad.PdfPig.Content.Page page)
    {
        var words  = page.GetWords().ToList();
        var groups = new List<List<UglyToad.PdfPig.Content.Word>>();

        foreach (var w in words)
        {
            var g = groups.FirstOrDefault(gr =>
                Math.Abs(gr[0].BoundingBox.Bottom - w.BoundingBox.Bottom) < 5);
            if (g != null) g.Add(w);
            else groups.Add(new List<UglyToad.PdfPig.Content.Word> { w });
        }

        return groups.OrderByDescending(g => g[0].BoundingBox.Bottom)
                     .Select(g => string.Join(" ",
                         g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)).Trim())
                     .Where(l => !string.IsNullOrWhiteSpace(l));
    }
}
