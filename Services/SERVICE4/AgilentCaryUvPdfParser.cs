using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Agilent Cary-3500 UV-Vis PDF 파일 파서 (Concentration report)
/// 구조:
///   1페이지 = 설정 (파일명, 파장, Fit type 등)
///   2페이지 = 검량선 그래프 + Zero Collection
///   3페이지 = Standard 표 + 회귀식 (y = a * x + b, r²)
///   4페이지~ = Sample 표 (시료명, Conc, Abs, [flag], ZeroAbs)
/// </summary>
public static class AgilentCaryUvPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    private static readonly HashSet<string> ControlKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "MBK", "FBK", "BK", "CCV", "DW", "BLANK" };

    private static readonly HashSet<string> FlagTokens = new(StringComparer.OrdinalIgnoreCase)
        { "o", "u", "r", "n" };

    private static readonly string LogPath = Path.Combine(
        Directory.GetCurrentDirectory(), "Logs", "Users", Environment.UserName, "AgilentCaryUV.log");

    private static void Log(string msg)
    {
        if (App.EnableLogging)
        {
            try
            {
                var dir = Path.GetDirectoryName(LogPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(LogPath, $"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            }
            catch { }
        }
    }

    /// <summary>
    /// PDF 파일이 Cary UV 형식인지 확인
    /// </summary>
    public static bool IsCaryUvPdf(string path)
    {
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var firstPage = doc.GetPage(1);
            string text = string.Join(" ", firstPage.GetWords().Select(w => w.Text));
            return text.Contains("Cary", StringComparison.OrdinalIgnoreCase)
                && text.Contains("Concentration", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = true };
        string item = activeItems.FirstOrDefault() ?? "UVVIS";

        if (App.EnableLogging)
        {
            File.WriteAllText(LogPath, $"=== AgilentCaryUvPdf Parse Start: {Path.GetFileName(path)} ===\n");
        }

        using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
        var pages = doc.GetPages().ToList();
        Log($"총 {pages.Count}페이지");

        string? docDate = null;
        double wavelength = 0;

        // 1페이지: 설정
        var p1Lines = ExtractLines(pages[0]);
        Log($"\n--- 1페이지 설정 ---");
        foreach (var l in p1Lines) Log($"  [{l}]");
        ParseSettingsPage(p1Lines, ref wavelength, ref docDate);

        // 2페이지~: Standard/Sample 표
        for (int pi = 1; pi < pages.Count; pi++)
        {
            var lines = ExtractLines(pages[pi]);
            Log($"\n--- 페이지 {pi + 1} ({lines.Count}줄) ---");
            foreach (var l in lines) Log($"  [{l}]");
            ParseDataPage(lines, rows, docInfo, item, resultFormatter);
        }

        // 파장으로 카테고리 결정
        if (wavelength > 0)
            docInfo.DetectedCategory = WavelengthToCategory(wavelength);

        docInfo.분석방법 = "Agilent Cary-3500 UV-Vis";
        docInfo.결과표시 = "mg/L";

        Log($"\n결과: 시료 {rows.Count}건, 파장={wavelength}nm, 카테고리={docInfo.DetectedCategory}");
        Log($"  기울기={docInfo.Standard_Slope}, R²={docInfo.Abs_R2}");

        return new ParseResult(rows, docInfo, docDate, "AGILENT_CARY3500_PDF");
    }

    // ── 텍스트 추출 (ShimadzuUvPdfParser와 동일 방식) ────────────────────────
    private static List<string> ExtractLines(UglyToad.PdfPig.Content.Page page)
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
                     .Select(g => string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)).Trim())
                     .Where(l => !string.IsNullOrWhiteSpace(l))
                     .ToList();
    }

    // ── 1페이지: 파장, 날짜 ──────────────────────────────────────────────────
    private static void ParseSettingsPage(List<string> lines, ref double wavelength, ref string? docDate)
    {
        bool nextIsWavelength = false;
        foreach (var line in lines)
        {
            // "Wavelengths:" 다음 줄이 파장값
            if (line.TrimEnd(':').Equals("Wavelengths", StringComparison.OrdinalIgnoreCase))
            {
                nextIsWavelength = true;
                continue;
            }
            if (nextIsWavelength)
            {
                var wlMatch = Regex.Match(line, @"^([\d.]+)\s*nm$", RegexOptions.IgnoreCase);
                if (wlMatch.Success)
                    double.TryParse(wlMatch.Groups[1].Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out wavelength);
                nextIsWavelength = false;
            }
            // "Wavelengths: 620.00 nm" 가 한 줄인 경우
            var inlineWl = Regex.Match(line, @"Wavelengths[:\s]+([\d.]+)\s*nm", RegexOptions.IgnoreCase);
            if (inlineWl.Success && wavelength == 0)
                double.TryParse(inlineWl.Groups[1].Value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out wavelength);

            // 날짜: "2026-03-30 ..."
            if (docDate == null)
            {
                var dm = Regex.Match(line, @"(\d{4}-\d{2}-\d{2})");
                if (dm.Success) docDate = dm.Groups[1].Value;
            }
        }
    }

    // ── 데이터 페이지: Standard/Sample 표 ────────────────────────────────────
    private static readonly Regex FormulaRe = new(
        @"y\s*=\s*([-\d.]+)\s*\*\s*x\s*([+\-])\s*([\d.]+).*r.=\s*([\d.]+)",
        RegexOptions.IgnoreCase);

    private static void ParseDataPage(List<string> lines, List<ExcelRow> rows,
        ExcelDocInfo docInfo, string item, Func<string, string, string>? formatter)
    {
        bool inStd    = false;
        bool inSample = false;
        var stdConcs  = new List<string>();
        var stdAbs    = new List<string>();

        foreach (var line in lines)
        {
            // 섹션 헤더 감지: "Standard 620.00 (nm)" / "Sample 620.00 (nm)"
            if (Regex.IsMatch(line, @"^Standard\s+[\d.]+\s*\(nm\)", RegexOptions.IgnoreCase))
            {
                inStd = true; inSample = false;
                Log($"  → Standard 섹션 진입");
                continue;
            }
            if (Regex.IsMatch(line, @"^Sample\s+[\d.]+\s*\(nm\)", RegexOptions.IgnoreCase))
            {
                inSample = true; inStd = false;
                Log($"  → Sample 섹션 진입");
                continue;
            }
            if (line.StartsWith("Legend", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Zero Collection", StringComparison.OrdinalIgnoreCase))
            {
                inStd = inSample = false;
                continue;
            }

            // 회귀식: "620.00nm: y = 1.6302 * x + 0.0010, r²=0.9939"
            var fm = FormulaRe.Match(line);
            if (fm.Success)
            {
                docInfo.Standard_Slope = fm.Groups[1].Value;
                double intercept = double.Parse(fm.Groups[3].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                if (fm.Groups[2].Value == "-") intercept = -intercept;
                docInfo.Standard_Intercept = intercept.ToString("F4",
                    System.Globalization.CultureInfo.InvariantCulture);
                docInfo.Abs_R2 = fm.Groups[4].Value;
                Log($"  검량선: slope={docInfo.Standard_Slope}, r²={docInfo.Abs_R2}");
                continue;
            }

            if (inStd)
            {
                // "Standard 1 0.000 0.0000 -0.0921"
                var stm = Regex.Match(line, @"^Standard\s+\d+\s+([-\d.]+)\s+([-\d.]+)",
                    RegexOptions.IgnoreCase);
                if (stm.Success)
                {
                    stdConcs.Add(stm.Groups[1].Value);
                    stdAbs.Add(stm.Groups[2].Value);
                }
            }
            else if (inSample)
            {
                var parsed = ParseSampleLine(line);
                if (parsed == null) continue;
                var (name, conc, abs) = parsed.Value;

                // 음수 농도 → 0
                if (double.TryParse(conc, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var cv) && cv < 0)
                    conc = "0";

                bool isCtrl = ControlKeywords.Any(k =>
                    name.StartsWith(k, StringComparison.OrdinalIgnoreCase));

                Log($"  ROW: name=[{name}] conc={conc} abs={abs} ctrl={isCtrl}");

                rows.Add(new ExcelRow
                {
                    시료명    = name,
                    SN        = "",
                    D1        = abs,
                    Result    = formatter?.Invoke(conc, item) ?? conc,
                    Source    = SourceType.미분류,
                    Status    = MatchStatus.대기,
                    IsControl = isCtrl,
                });
            }
        }

        if (stdConcs.Count > 0 && docInfo.Standard_Points.Length == 0)
        {
            docInfo.Standard_Points = stdConcs.ToArray();
            docInfo.Abs_Values      = stdAbs.ToArray();
        }
    }

    // ── 시료 줄 파싱: {Name} {Conc} {Abs} [{flag}] {ZeroAbs} ─────────────────
    private static (string Name, string Conc, string Abs)? ParseSampleLine(string line)
    {
        var tokens = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length < 4) return null;

        int idx = tokens.Length - 1;

        // ZeroAbs: 마지막 토큰 (숫자)
        if (!IsNum(tokens[idx])) return null;
        idx--; // ZeroAbs 건너뜀

        // 선택적 플래그 (o/u/r/n)
        if (idx >= 0 && IsFlagToken(tokens[idx])) idx--;

        // Abs
        if (idx < 0 || !IsNum(tokens[idx])) return null;
        string abs = tokens[idx--];

        // Conc
        if (idx < 0 || !IsNum(tokens[idx])) return null;
        string conc = tokens[idx--];

        // Name = 나머지
        if (idx < 0) return null;
        string name = string.Join(" ", tokens[..(idx + 1)]).Trim();
        if (string.IsNullOrWhiteSpace(name)) return null;

        return (name, conc, abs);
    }

    private static bool IsNum(string s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out _);

    private static bool IsFlagToken(string s) =>
        FlagTokens.Contains(s) || s.Equals("cf", StringComparison.OrdinalIgnoreCase);

    // ── 파장 → 카테고리 ──────────────────────────────────────────────────────
    private static string? WavelengthToCategory(double wl) => wl switch
    {
        >= 615 and <= 625 => "CN",      // 620nm: 시안
        >= 535 and <= 545 => "CR6",     // 540nm: 6가크롬
        >= 875 and <= 885 => "TP",      // 880nm: T-P
        >= 215 and <= 225 => "TN",      // 220nm: T-N
        >= 455 and <= 465 => "PHENOLS", // 460nm: 페놀류
        >= 360 and <= 370 => "COLOR",   // 365nm: 색도
        _ => null
    };
}
