using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Shimadzu UV-1800 PDF 파일 파서
/// 구조:
///   1페이지 = Standard Table Report (검량선: y=ax±b, r2, ST1~5)
///   2페이지~ = Sample Table Report  (시료: Sample ID, Type, Conc, WL880.0)
/// </summary>
public static class ShimadzuUvPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    private static readonly HashSet<string> ControlKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "MBK", "FBK", "BK", "CCV", "DW", "BLANK" };

    private static readonly string LogPath = Path.Combine(
        Directory.GetCurrentDirectory(), "Logs", "Users", Environment.UserName, "ShimadzuUV.log");

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

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsUVVIS = true };
        string item = activeItems.FirstOrDefault() ?? "UVVIS";

        // 로그 초기화
        if (App.EnableLogging)
        {
            var logDir = Path.GetDirectoryName(LogPath)!;
            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
            File.WriteAllText(LogPath, $"=== ShimadzuUvPdf Parse Start: {Path.GetFileName(path)} ===\n");
        }

        using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
        var pages = doc.GetPages().ToList();
        Log($"총 {pages.Count}페이지");

        string? docDate = null;

        for (int pi = 0; pi < pages.Count; pi++)
        {
            var lines = ExtractLines(pages[pi]);

            // 모든 추출 줄 로깅
            Log($"\n--- 페이지 {pi + 1} ({lines.Count}줄) ---");
            foreach (var l in lines) Log($"  [{l}]");

            if (pi == 0)
            {
                ParseStandardPage(lines, docInfo);
                docDate = ExtractDate(lines);
                // 파일명/폴더명에서 UV 항목 감지 → DetectedCategory
                if (docInfo.DetectedCategory == null)
                    docInfo.DetectedCategory = DetectUvItemFromLines(lines);
                // 감지된 항목으로 item 재설정
                if (!string.IsNullOrEmpty(docInfo.DetectedCategory))
                    item = CatToItemName(docInfo.DetectedCategory);
            }
            else
                ParseSamplePage(lines, rows, item, resultFormatter);
        }

        Log($"\n결과: 시료 {rows.Count}건, 기울기={docInfo.Standard_Slope}, R²={docInfo.Abs_R2}, 날짜={docDate}, 항목={docInfo.DetectedCategory}");
        return new ParseResult(rows, docInfo, docDate, "Shimadzu_UV_PDF");
    }

    // ── 텍스트 추출 ──────────────────────────────────────────────────────────
    private static List<string> ExtractLines(UglyToad.PdfPig.Content.Page page)
    {
        var words  = page.GetWords().ToList();
        var groups = new List<List<UglyToad.PdfPig.Content.Word>>();
        foreach (var w in words)
        {
            var g = groups.FirstOrDefault(gr =>
                Math.Abs(gr[0].BoundingBox.Bottom - w.BoundingBox.Bottom) < 5); // tolerance 5
            if (g != null) g.Add(w);
            else groups.Add(new List<UglyToad.PdfPig.Content.Word> { w });
        }
        return groups.OrderByDescending(g => g[0].BoundingBox.Bottom)
                     .Select(g => string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)).Trim())
                     .Where(l => !string.IsNullOrWhiteSpace(l))
                     .ToList();
    }

    // ── 1페이지: 검량선 ──────────────────────────────────────────────────────
    private static void ParseStandardPage(List<string> lines, ExcelDocInfo docInfo)
    {
        var stdConcs = new List<string>();
        var stdAbs   = new List<string>();

        var eqRe = new Regex(@"y\s*=\s*([-\d.Ee+]+)\s*x\s*([+\-]\s*[\d.Ee+]+)", RegexOptions.IgnoreCase);
        var r2Re = new Regex(@"r2\s*=\s*([\d.]+)", RegexOptions.IgnoreCase);
        var stRe = new Regex(@"ST\d+\s+Standard\s+([-\d.]+)\s+([-\d.]+)", RegexOptions.IgnoreCase);

        foreach (var line in lines)
        {
            var eq = eqRe.Match(line);
            if (eq.Success)
            {
                docInfo.Standard_Slope     = eq.Groups[1].Value;
                docInfo.Standard_Intercept = eq.Groups[2].Value.Replace(" ", "");
                Log($"검량선: slope={docInfo.Standard_Slope}, intercept={docInfo.Standard_Intercept}");
                continue;
            }
            var r2 = r2Re.Match(line);
            if (r2.Success) { docInfo.Abs_R2 = r2.Groups[1].Value; Log($"R²={docInfo.Abs_R2}"); continue; }

            var st = stRe.Match(line);
            if (st.Success)
            {
                stdConcs.Add(st.Groups[1].Value);
                stdAbs.Add(st.Groups[2].Value);
            }
        }

        docInfo.Standard_Points = stdConcs.ToArray();
        docInfo.Abs_Values      = stdAbs.ToArray();
        docInfo.분석방법 = "Shimadzu UV-1800";
    }

    // ── 날짜 추출: "2026-04-07  02:39:44 오" 패턴 ────────────────────────────
    private static readonly Regex DateRe = new(@"(\d{4}-\d{2}-\d{2})", RegexOptions.Compiled);
    private static string? ExtractDate(List<string> lines)
    {
        foreach (var line in lines)
        {
            var m = DateRe.Match(line);
            if (m.Success && DateTime.TryParse(m.Groups[1].Value, out _))
            {
                Log($"날짜 감지: {m.Groups[1].Value}");
                return m.Groups[1].Value;
            }
        }
        return null;
    }

    // ── UV 항목 감지: File Name 경로 또는 페이지 텍스트에서 ────────────────────
    private static readonly (string Keyword, string Category)[] UvKeywords =
    [
        ("Phenols", "PHENOLS"), ("페놀",   "PHENOLS"),
        ("T-P",     "TP"),      ("TP",     "TP"),
        ("T-N",     "TN"),      ("TN",     "TN"),
        ("시안",    "CN"),      ("CN",     "CN"),
        ("6가크롬", "CR6"),     ("Cr6",    "CR6"),
        ("색도",    "COLOR"),
        ("ABS",     "ABS"),
        ("불소",    "FLUORIDE"),
    ];

    private static string? DetectUvItemFromLines(List<string> lines)
    {
        // "File Name:" 줄 우선 검색 (경로에 폴더명/파일명으로 항목 표시)
        var fileNameLine = lines.FirstOrDefault(l =>
            l.StartsWith("File Name", StringComparison.OrdinalIgnoreCase) ||
            l.StartsWith("File:") ||
            l.Contains(":\\") || l.Contains(":/"));

        var searchTargets = fileNameLine != null
            ? new[] { fileNameLine }.Concat(lines)
            : lines;

        foreach (var line in searchTargets)
            foreach (var (kw, cat) in UvKeywords)
                if (line.Contains(kw, StringComparison.OrdinalIgnoreCase))
                {
                    Log($"UV 항목 감지: '{kw}' → {cat}  (줄: {line})");
                    return cat;
                }
        return null;
    }

    private static string CatToItemName(string category) => category switch
    {
        "PHENOLS"  => "Phenols",
        "TP"       => "T-P",
        "TN"       => "T-N",
        "CN"       => "시안",
        "CR6"      => "6가크롬",
        "COLOR"    => "색도",
        "ABS"      => "ABS",
        "FLUORIDE" => "불소",
        _          => "UVVIS",
    };

    // ── 2페이지~: 시료 데이터 ────────────────────────────────────────────────
    // 전략: 줄에 "Unknown"이 포함되면 파싱 시도
    //       Unknown 앞 = 시료명, Unknown 뒤 = [✓] conc wl880
    private static readonly Regex NumRe = new(@"^-?[\d.]+$");

    private static void ParseSamplePage(List<string> lines, List<ExcelRow> rows,
        string item, Func<string, string, string>? formatter)
    {
        foreach (var line in lines)
        {
            // "Unknown" 포함 행만 처리 (Standard는 제외)
            int uIdx = line.IndexOf("Unknown", StringComparison.OrdinalIgnoreCase);
            if (uIdx < 0) continue;

            // Unknown 앞부분에서 시료명 추출
            // 앞부분: "{rowNum} {sampleName}" 또는 "{sampleName}"
            string before = line[..uIdx].Trim();
            // 선두 숫자(행번호) 제거
            before = Regex.Replace(before, @"^\d+\s+", "").Trim();
            string name = before;
            if (string.IsNullOrWhiteSpace(name)) continue;

            // Unknown 뒷부분에서 숫자 추출
            string after = line[(uIdx + 7)..].Trim(); // "Unknown" = 7글자
            // ✓ 또는 체크마크 제거
            after = Regex.Replace(after, @"^[✓✗\u2713\u2717]\s*", "").Trim();

            // 숫자 토큰 추출
            var tokens = after.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                              .Where(t => NumRe.IsMatch(t))
                              .ToArray();

            if (tokens.Length < 2)
            {
                Log($"  SKIP (토큰 부족): [{line}] → after=[{after}] tokens={tokens.Length}");
                continue;
            }

            string conc = tokens[0]; // Conc
            string abs  = tokens[1]; // WL880.0

            // 음수 농도 → 0
            if (double.TryParse(conc, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var cv) && cv < 0)
                conc = "0";

            bool isCtrl = ControlKeywords.Any(k => name.StartsWith(k, StringComparison.OrdinalIgnoreCase))
                       || name.Contains("표준첨가") || name.StartsWith("CCV", StringComparison.OrdinalIgnoreCase);

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
}
