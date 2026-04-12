using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Analytik Jena PQ 9000 ICP-OES / VGA PDF 결과 파서
/// OES 모드: Ba, Cd, Cr, Cu, Fe, Mn, Ni, Pb, Sn, Zn 등 다원소
/// VGA 모드: As, Sb, Se 등 기화원소
///
/// PDF 구조:
///   시료 헤더: "SAMPLENAME  Method: OES(1)  Orig. Date: YYYY-MM-DD HH:MM"
///   검량선 헤더: "Compute calib.  Method: Ba455.403  Date: ..."
///   원소줄: "Ba455.403"  (원소기호+파장, 단독줄)
///   농도줄: "Conc.1  VALUE  [SD RSD CI]  mg/L  [&lt;KAL]"
///   강도줄: "Ints.  MEAN  [SD RSD]"
///   계수줄: "y=a+bx  a=X  b=Y  Method SD: ..."
///   R²줄:   "R²(adj.): X ; ..."
///
///   건너뜀: method blank, BLANK, Reag. blank, ST1~ST5, Compute calib.
///   정도관리(IsControl): MBK, FBK*, CCV*, BK*, DW*
///   불검출: &lt;KAL 있음 또는 값 &lt;= 0
/// </summary>
public static class IcpPdfParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo DocInfo,
        string? DocDate,
        string DetectedFormat);

    // ── Skip 시료 ────────────────────────────────────────────────────────
    private static readonly HashSet<string> SkipSampleExact =
        new(StringComparer.OrdinalIgnoreCase)
        { "method blank", "BLANK", "Reag. blank" };

    private static readonly Regex SkipSampleRx = new(
        @"^ST\d+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex CalibSampleRx = new(
        @"^Compute\s+calib\.$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static bool IsQcSample(string name) =>
        name.StartsWith("MBK",  StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("FBK",  StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("CCV",  StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("BK",   StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("DW",   StringComparison.OrdinalIgnoreCase) ||
        name.Equals("BLANK",    StringComparison.OrdinalIgnoreCase);

    // ── Regex ─────────────────────────────────────────────────────────────
    // 시료 헤더: "SAMPLENAME Method: OES(1) ..."
    // 검량선 헤더: "Compute calib. Method: Ba455.403 ..." → 음수 전방탐색으로 제외
    private static readonly Regex SampleHeaderRx = new(
        @"^(.+?)\s+Method:\s+(?![A-Z][a-z]?\d)", RegexOptions.Compiled);

    // 원소줄: "Ba455.403", "As193.698"
    private static readonly Regex ElementRx = new(
        @"^([A-Z][a-z]?)\d{3}\.\d{3}$", RegexOptions.Compiled);

    // 날짜: "Orig. Date: 2025-11-28"
    private static readonly Regex DateRx = new(
        @"Orig\.\s*Date:\s*(\d{4}-\d{2}-\d{2})", RegexOptions.Compiled);

    // 검량선 헤더: "Compute calib. Method: Ba455.403 Date: ..."
    private static readonly Regex CalibHeaderRx = new(
        @"^Compute\s+calib\.\s+Method:\s+([A-Z][a-z]?\d{3}\.\d{3})", RegexOptions.Compiled);

    // y=a+bx 계수줄: "y=a+bx a=49713.912 b=2904184.1 Method SD: ..."
    private static readonly Regex CalibFormulaRx = new(
        @"y=a\+bx\s+a=([\d.\-eE+]+)\s+b=([\d.\-eE+]+)", RegexOptions.Compiled);

    // R²줄: "R²(adj.): 0.999592802 ; ..."
    private static readonly Regex CalibR2Rx = new(
        @"R.{1,5}adj\..{1,3}:\s*([\d.]+)", RegexOptions.Compiled);

    // 페이지/파일 헤더 무시
    private static readonly Regex PageHeaderRx = new(
        @"^(PQ 9000|ASpect PQ|Operator:|Lab\.:|Results$|Results file:|Instrument:|#\d|Comment:|Line Mean|Single values|Blank unit|Pre-DF:)",
        RegexOptions.Compiled);

    /// <summary>Analytik Jena PQ 9000 PDF 여부 확인</summary>
    public static bool IsIcpOesPdf(string path)
    {
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var page = doc.GetPage(1);
            string text = string.Join(" ", page.GetWords().Select(w => w.Text));
            return text.Contains("PQ 9000", StringComparison.OrdinalIgnoreCase) &&
                   text.Contains("Analytik", StringComparison.OrdinalIgnoreCase);
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
        string instrumentName = "Analytik Jena PQ 9000";
        string gcFormat       = "ICP-OES";
        bool   dateFound      = false;

        // ── 상태 머신 ────────────────────────────────────────────────────
        string? currentSample  = null;
        bool    skipCurrent    = false;
        bool    isStd          = false;   // ST1~5: 표준용액
        bool    isCalib        = false;   // Compute calib. 블록
        bool    isControl      = false;
        string? currentElement = null;
        bool    concCaptured   = false;
        bool    intsCaptured   = false;

        // ── 검정곡선 수집 ─────────────────────────────────────────────────
        // key = 원소기호, value = 누적 표준점 리스트
        var stdData  = new Dictionary<string, List<(string Conc, string Ints)>>();
        var calData  = new Dictionary<string, (string A, string B, string R2)>();
        string? calibElem = null;
        string? calibA    = null;
        string? calibB    = null;

        // 표준용액 수집용: 현재 ST 원소에서 Conc.1 대기
        string? stdConc  = null;

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);

            foreach (var page in doc.GetPages())
            {
                foreach (var line in ExtractLines(page))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (PageHeaderRx.IsMatch(line)) continue;

                    // ── 날짜 추출 (최초 1회) ───────────────────────────
                    if (!dateFound)
                    {
                        var dm = DateRx.Match(line);
                        if (dm.Success) { analysisDate = dm.Groups[1].Value; dateFound = true; }
                    }

                    // ── Compute calib. 헤더 ───────────────────────────
                    var chm = CalibHeaderRx.Match(line);
                    if (chm.Success)
                    {
                        var em = ElementRx.Match(chm.Groups[1].Value);
                        calibElem = em.Success ? em.Groups[1].Value : null;
                        calibA = null; calibB = null;
                        currentSample = null;
                        isCalib = true; skipCurrent = true; isStd = false;
                        currentElement = null; concCaptured = false; intsCaptured = false;
                        continue;
                    }

                    // ── 검량곡선 계수 (y=a+bx) ────────────────────────
                    if (isCalib && calibElem != null)
                    {
                        var fm = CalibFormulaRx.Match(line);
                        if (fm.Success)
                        {
                            calibA = fm.Groups[1].Value;
                            calibB = fm.Groups[2].Value;
                            continue;
                        }
                        var r2m = CalibR2Rx.Match(line);
                        if (r2m.Success && calibA != null && calibB != null)
                        {
                            calData[calibElem] = (calibA, calibB, r2m.Groups[1].Value);
                            calibElem = null; calibA = null; calibB = null;
                            isCalib = false;
                            continue;
                        }
                    }

                    // ── 시료 헤더 ─────────────────────────────────────
                    var sm = SampleHeaderRx.Match(line);
                    if (sm.Success)
                    {
                        string sname   = sm.Groups[1].Value.Trim();
                        currentSample  = sname;
                        currentElement = null;
                        concCaptured   = false;
                        intsCaptured   = false;
                        stdConc        = null;
                        isCalib        = false;

                        isStd       = SkipSampleRx.IsMatch(sname);
                        skipCurrent = SkipSampleExact.Contains(sname) ||
                                      CalibSampleRx.IsMatch(sname);
                        isControl   = !skipCurrent && !isStd && IsQcSample(sname);

                        // ST 블록은 표준용액 수집을 위해 처리하되
                        // 시료결과로는 추가하지 않음
                        if (isStd) skipCurrent = false;
                        continue;
                    }

                    if (skipCurrent) continue;
                    if (currentSample == null) continue;

                    // ── 원소줄 ────────────────────────────────────────
                    var em2 = ElementRx.Match(line);
                    if (em2.Success)
                    {
                        currentElement = em2.Groups[1].Value;
                        concCaptured   = false;
                        intsCaptured   = false;
                        stdConc        = null;
                        continue;
                    }

                    // ── Conc.1 줄 ─────────────────────────────────────
                    if (!concCaptured && currentElement != null &&
                        line.StartsWith("Conc.1", StringComparison.OrdinalIgnoreCase))
                    {
                        var tokens = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                        if (tokens.Length >= 2)
                        {
                            if (isStd)
                            {
                                // 표준용액: 농도값만 저장 (Ints.에서 응답 수집)
                                stdConc = tokens[1];
                            }
                            else
                            {
                                bool isKal = tokens[^1].Equals("<KAL", StringComparison.OrdinalIgnoreCase);
                                string result = "";
                                if (!isKal &&
                                    double.TryParse(tokens[1],
                                        NumberStyles.Float, CultureInfo.InvariantCulture, out double val) &&
                                    val > 0)
                                {
                                    result = val.ToString("G6", CultureInfo.InvariantCulture);
                                }

                                rows.Add(new ExcelRow
                                {
                                    시료명       = currentSample,
                                    CompoundName = currentElement,
                                    Result       = result,
                                    Source       = SourceType.미분류,
                                    Status       = MatchStatus.대기,
                                    IsControl    = isControl,
                                });
                                concCaptured = true;
                            }
                        }
                    }

                    // ── Ints. 줄 (표준용액에서 응답 수집) ────────────────
                    if (isStd && currentElement != null && stdConc != null &&
                        !intsCaptured && line.StartsWith("Ints.", StringComparison.OrdinalIgnoreCase))
                    {
                        var tokens = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                        if (tokens.Length >= 2)
                        {
                            if (!stdData.ContainsKey(currentElement))
                                stdData[currentElement] = new List<(string, string)>();
                            stdData[currentElement].Add((stdConc!, tokens[1]));
                            stdConc      = null;
                            intsCaptured = true;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"ICP PDF 파싱 오류: {ex.Message}", ex);
        }

        // ── 검정곡선 빌드 ─────────────────────────────────────────────────
        // 원소 순서: calData 기준 + stdData에만 있는 원소 추가
        var elemOrder = calData.Keys.ToList();
        foreach (var k in stdData.Keys)
            if (!elemOrder.Contains(k)) elemOrder.Add(k);

        if (elemOrder.Count > 0)
        {
            docInfo.IsGcMode  = true;
            docInfo.GcFormat  = gcFormat;

            foreach (var elem in elemOrder)
            {
                var cal = new GcCompoundCalInfo { Name = elem };

                if (calData.TryGetValue(elem, out var cd))
                {
                    cal.Slope     = cd.B;
                    cal.Intercept = cd.A;
                    cal.R         = cd.R2;
                }

                if (stdData.TryGetValue(elem, out var pts))
                {
                    cal.StdConcs = pts.Select(p => p.Conc).ToArray();
                    cal.StdResps = pts.Select(p => p.Ints).ToArray();
                }

                docInfo.GcCompoundCals.Add(cal);
            }
        }

        docInfo.분석방법 = instrumentName;
        docInfo.결과표시 = "mg/L";

        return new ParseResult(rows, docInfo, analysisDate, "ICP_OES_PDF");
    }

    // ── PdfPig 줄 추출 (Y 좌표 ±5pt로 단어 묶기) ───────────────────────────
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
