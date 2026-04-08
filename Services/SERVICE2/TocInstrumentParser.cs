using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ETA.Services.SERVICE2;

public class TocInstrumentRow
{
    public string RawName  { get; set; } = "";
    public string SN       { get; set; } = ""; // 정규화된 SN (03-11-04 형식)
    public string Area     { get; set; } = "";
    public string Conc     { get; set; } = ""; // 최종 농도 (mg/L)
    public string Method   { get; set; } = "NPOC"; // NPOC 또는 TCIC
    public bool   IsControl { get; set; } = false;
    // TCIC 전용: TC/IC 분리값 (TOC_TCIC_DATA 컬럼 매핑)
    public string TCAU     { get; set; } = ""; // TC absorbance / area
    public string TCcon    { get; set; } = ""; // TC 농도 (mg/L)
    public string ICAU     { get; set; } = ""; // IC absorbance / area
    public string ICcon    { get; set; } = ""; // IC 농도 (mg/L)
    public string Dilution { get; set; } = ""; // 희석배수
}

public class TocCalibrationData
{
    public string Slope_TC     { get; set; } = ""; // TC 기울기 (NPOC=TC만 사용)
    public string Intercept_TC { get; set; } = "";
    public string Slope_IC     { get; set; } = ""; // IC 기울기 (가감법만)
    public string Intercept_IC { get; set; } = "";
    public string R2_TC        { get; set; } = "";
    public string R2_IC        { get; set; } = "";
    public string Method       { get; set; } = "NPOC";
    // 표준점 (ST-1~ST-5): 농도 / 면적
    public string[] StdConcs    { get; set; } = []; // TC 공칭 농도
    public string[] StdAreas    { get; set; } = []; // TC 측정 면적(AU)
    public string[] StdConcs_IC { get; set; } = []; // IC 공칭 농도 (TCIC 전용)
    public string[] StdAreas_IC { get; set; } = []; // IC 측정 면적(AU) (TCIC 전용)
    public bool   HasData      => !string.IsNullOrEmpty(Slope_TC);
}

/// <summary>TOC 분석기(스칼라/시마즈) 출력 파일 파서</summary>
public static class TocInstrumentParser
{
    // 바탕/정도관리 시료 키워드 (SN 매칭 제외)
    static readonly HashSet<string> ControlKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "DW", "BK", "MBK", "FBK", "CCV", "ST", "NPOC", "TCIC", "BLANK" };

    // ── 포맷 감지 ──────────────────────────────────────────────────────────

    public enum TocFileFormat { Unknown, Skalar, ShimadzuSimple, ShimadzuDetail, JenaCalibration, JenaAnalysisPdf, TcicXlsxNote }

    public static TocFileFormat DetectFormat(string path)
    {
        try
        {
            // PDF → 예나 분석표
            if (path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                return TocFileFormat.JenaAnalysisPdf;

            // XLSX → TOC(TCIC) NOTE 레거시 양식 확인
            if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var wb = new XLWorkbook(fs);
                    var ws = wb.Worksheets.FirstOrDefault();
                    if (ws != null)
                    {
                        var r5a = ws.Cell(5, 1).GetString().Trim();
                        var r6a = ws.Cell(6, 1).GetString().Trim();
                        var r7a = ws.Cell(7, 1).GetString().Trim();
                        // 시그니처: R5='Standard (TC)', R6='AU', R7='Standard (IC)'
                        if (r5a.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            r5a.IndexOf("TC",       StringComparison.OrdinalIgnoreCase) >= 0 &&
                            r6a.Equals("AU",        StringComparison.OrdinalIgnoreCase) &&
                            r7a.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            r7a.IndexOf("IC",       StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return TocFileFormat.TcicXlsxNote;
                        }
                    }
                }
                catch { }
                return TocFileFormat.Unknown;
            }

            // UTF-16 LE BOM 확인 → 스칼라
            var bom = new byte[2];
            using (var fs = File.OpenRead(path))
                fs.Read(bom, 0, 2);
            if (bom[0] == 0xFF && bom[1] == 0xFE)
                return TocFileFormat.Skalar;

            // 텍스트로 읽어 시마즈 [Header] 구조 확인 (CP949 또는 UTF-8)
            string firstLine;
            try { firstLine = File.ReadLines(path, Encoding.UTF8).FirstOrDefault() ?? ""; }
            catch { firstLine = File.ReadLines(path, Encoding.Latin1).FirstOrDefault() ?? ""; }

            if (firstLine.TrimStart().StartsWith("[Header]", StringComparison.OrdinalIgnoreCase))
            {
                // 디테일 여부: "Inj. No." 컬럼 포함 여부
                var lines = ReadShimadzuLines(path);
                bool isDetail = lines.Any(l => l.Contains("Inj. No.") || l.Contains("Inj.No."));
                return isDetail ? TocFileFormat.ShimadzuDetail : TocFileFormat.ShimadzuSimple;
            }

            // 예나 검량선 CSV: 첫 줄이 "Cal_"로 시작하고 세미콜론 구분자
            if (firstLine.StartsWith("Cal_", StringComparison.OrdinalIgnoreCase) && firstLine.Contains(';'))
                return TocFileFormat.JenaCalibration;
        }
        catch { }
        return TocFileFormat.Unknown;
    }

    // ── 메인 파서 ──────────────────────────────────────────────────────────

    public static (TocFileFormat format, List<TocInstrumentRow> rows, TocCalibrationData? cal) Parse(string path)
    {
        var fmt = DetectFormat(path);
        TocCalibrationData? cal = null;
        List<TocInstrumentRow> rows;
        switch (fmt)
        {
            case TocFileFormat.Skalar:
                rows = ParseSkalar(path, out cal);
                break;
            case TocFileFormat.ShimadzuSimple:
                rows = ParseShimadzuSimple(path, out cal);
                break;
            case TocFileFormat.ShimadzuDetail:
                rows = ParseShimadzuDetail(path, out cal);
                break;
            case TocFileFormat.JenaCalibration:
                (rows, cal) = ParseJena(path);
                break;
            case TocFileFormat.JenaAnalysisPdf:
                (rows, cal) = ParseJenaPdf(path);
                break;
            case TocFileFormat.TcicXlsxNote:
                (rows, cal) = ParseTcicXlsx(path);
                break;
            default:
                rows = new List<TocInstrumentRow>();
                break;
        }
        return (fmt, rows, cal);
    }

    // ── 스칼라 CSV (UTF-16 LE) ──────────────────────────────────────────────

    static List<TocInstrumentRow> ParseSkalar(string path, out TocCalibrationData? cal)
    {
        cal = null;
        var result      = new List<TocInstrumentRow>();
        var wsNomConcs  = new List<double>(); // 공칭 농도 (이름에서 추출)
        var wsAreas     = new List<double>(); // 측정 면적
        try
        {
            var text = File.ReadAllText(path, Encoding.Unicode);
            var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines.Skip(1)) // 헤더 건너뜀
            {
                var fields = SplitCsvLine(line);
                if (fields.Count < 6) continue;

                // 컬럼: [0]=순번, [1]=위치, [2]=시료명, [3]=타입(U/WS), [4]=면적, [5]=Conc, [6]=단위
                var sampleName = fields[2].Trim();
                var typeStr    = fields[3].Trim();

                if (typeStr.Equals("WS", StringComparison.OrdinalIgnoreCase))
                {
                    // 공칭 농도: 이름에서 숫자 추출 ("NPOC 5 mg/L" → 5.0, "NPOC 0" → 0.0)
                    var nomMatch = Regex.Match(sampleName, @"([\d]+(?:\.[\d]+)?)(?:\s*mg)?");
                    if (nomMatch.Success &&
                        double.TryParse(nomMatch.Groups[1].Value, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var wNomConc) &&
                        double.TryParse(fields[4].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var wArea))
                    {
                        wsNomConcs.Add(wNomConc);
                        wsAreas.Add(wArea);
                    }
                    continue;
                }

                if (!typeStr.Equals("U", StringComparison.OrdinalIgnoreCase)) continue;

                bool isCtrl = IsControlName(sampleName);
                var sn = NormalizeSN(sampleName);

                result.Add(new TocInstrumentRow
                {
                    RawName   = sampleName,
                    SN        = sn,
                    Area      = fields[4].Trim(),
                    Conc      = fields[5].Trim(),
                    Method    = "NPOC",
                    IsControl = isCtrl,
                });
            }

            // 선형 회귀: x=공칭농도, y=면적 → Area = a*Conc + b
            if (wsNomConcs.Count >= 2)
            {
                var (a, b, r2) = LinearRegression(wsNomConcs, wsAreas);
                cal = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = "NPOC",
                    StdConcs     = wsNomConcs.Select(c => c.ToString("F2")).ToArray(),
                    StdAreas     = wsAreas.Select(v => v.ToString("F4")).ToArray(),
                };
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[TocParser/Skalar] {ex.Message}"); }
        return result;
    }

    // ── 시마즈 요약 (값) ──────────────────────────────────────────────────

    static List<TocInstrumentRow> ParseShimadzuSimple(string path, out TocCalibrationData? cal)
    {
        cal = null;
        var result = new List<TocInstrumentRow>();
        try
        {
            var lines  = ReadShimadzuLines(path);
            int dataIdx = lines.FindIndex(l => l.TrimStart().StartsWith("[Data]"));
            if (dataIdx < 0) return result;

            // 헤더행 파싱
            var header = lines[dataIdx + 1].Split('\t');
            int idxType   = IndexOf(header, "Type");
            int idxAnal   = IndexOf(header, "Anal.");
            int idxName   = IndexOf(header, "Sample Name");
            int idxResult = IndexOf(header, "Result(NPOC)", "Result(TOC)", "Result(TC)");
            int idxArea   = IndexOf(header, "Area");
            int idxConc   = IndexOf(header, "Conc.");
            int idxExcl   = IndexOf(header, "Excluded");
            int idxSplNo  = IndexOf(header, "Spl. No.", "Spl.No.");

            if (idxName < 0 || idxResult < 0) return result;

            string method = IndexOf(header, "Result(NPOC)") >= 0 ? "NPOC" : "TCIC";
            var stdPoints = new Dictionary<int, (double conc, double area)>();

            for (int i = dataIdx + 2; i < lines.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = lines[i].Split('\t');
                if (cols.Length <= idxResult) continue;

                var typeStr = idxType >= 0 ? cols[idxType].Trim() : "";
                var excluded = idxExcl >= 0 && idxExcl < cols.Length ? cols[idxExcl].Trim() : "0";

                // Standard 행 → 검량선 데이터 수집
                if (typeStr.Equals("Standard", StringComparison.OrdinalIgnoreCase))
                {
                    if (excluded == "1") continue;
                    var stdAreaStr = idxArea >= 0 && idxArea < cols.Length ? cols[idxArea].Trim() : "";
                    var stdConcStr = idxConc >= 0 && idxConc < cols.Length ? cols[idxConc].Trim() : "";
                    int splNo = 0;
                    if (idxSplNo >= 0 && idxSplNo < cols.Length)
                        int.TryParse(cols[idxSplNo].Trim(), out splNo);

                    if (double.TryParse(stdConcStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var sc) &&
                        double.TryParse(stdAreaStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var sa))
                    {
                        stdPoints[splNo] = (sc, sa);
                    }
                    continue;
                }

                if (!typeStr.Equals("Unknown", StringComparison.OrdinalIgnoreCase)) continue;

                var sampleName = idxName < cols.Length ? cols[idxName].Trim() : "";
                var concStr    = cols[idxResult].Trim();

                bool isCtrl = IsControlName(sampleName);
                var sn = NormalizeSN(sampleName);

                result.Add(new TocInstrumentRow
                {
                    RawName   = sampleName,
                    SN        = sn,
                    Area      = "",
                    Conc      = concStr,
                    Method    = method,
                    IsControl = isCtrl,
                });
            }

            // Standard 포인트로 검량선 계산
            if (stdPoints.Count >= 2)
            {
                var sortedPts = stdPoints.OrderBy(p => p.Key).ToList();
                var concs = sortedPts.Select(p => p.Value.conc).ToList();
                var areas = sortedPts.Select(p => p.Value.area).ToList();
                var (a, b, r2) = LinearRegression(concs, areas);
                cal = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = method,
                    StdConcs     = concs.Select(c => c.ToString("F2")).ToArray(),
                    StdAreas     = areas.Select(v => v.ToString("F4")).ToArray(),
                };
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[TocParser/ShimadzuSimple] {ex.Message}"); }
        return result;
    }

    // ── 시마즈 상세 (주입 단위) ───────────────────────────────────────────

    static List<TocInstrumentRow> ParseShimadzuDetail(string path, out TocCalibrationData? cal)
    {
        cal = null;
        var result = new List<TocInstrumentRow>();
        try
        {
            var lines   = ReadShimadzuLines(path);
            int dataIdx = lines.FindIndex(l => l.TrimStart().StartsWith("[Data]"));
            if (dataIdx < 0) return result;

            var header  = lines[dataIdx + 1].Split('\t');
            int idxType = IndexOf(header, "Type");
            int idxName = IndexOf(header, "Sample Name");
            int idxSplNo = IndexOf(header, "Spl. No.", "Spl.No.");
            int idxArea  = IndexOf(header, "Area");
            int idxConc  = IndexOf(header, "Conc.");
            int idxExcl  = IndexOf(header, "Excluded");
            int idxAnal  = IndexOf(header, "Anal.");

            if (idxName < 0) return result;

            // ── 메서드 감지: 컬럼 헤더 우선, 없으면 첫 데이터 행의 Anal. 값 ──
            string fileMethod = "NPOC";
            // Detail 포맷: "TC Conc." / "IC Conc." 또는 "Result(NPOC)" 헤더로 감지
            if (IndexOf(header, "Result(NPOC)") >= 0)
                fileMethod = "NPOC";
            else if (IndexOf(header, "Result(TC)", "Result(TOC)", "Result(TCIC)") >= 0 ||
                     IndexOf(header, "TC Conc.", "NPOC Conc.") >= 0)
                fileMethod = "TCIC";
            else if (idxAnal >= 0)
            {
                // 헤더에 명시 없음 → 첫 번째 Unknown 행의 Anal. 값으로 판별
                for (int k = dataIdx + 2; k < lines.Count && k < dataIdx + 30; k++)
                {
                    if (string.IsNullOrWhiteSpace(lines[k])) continue;
                    var c0 = lines[k].Split('\t');
                    if (idxType >= 0 && idxType < c0.Length &&
                        !c0[idxType].Trim().Equals("Unknown", StringComparison.OrdinalIgnoreCase)) continue;
                    if (idxAnal < c0.Length)
                    {
                        var av = c0[idxAnal].Trim();
                        if (av.Contains("TC", StringComparison.OrdinalIgnoreCase) &&
                            av.Contains("IC", StringComparison.OrdinalIgnoreCase))
                            fileMethod = "TCIC";
                        break;
                    }
                }
            }

            // 그룹: 시료명별로 마지막 비제외 주입의 Conc를 사용
            var groups = new Dictionary<string, (string area, string conc)>(StringComparer.OrdinalIgnoreCase);
            // Standard 검량선 포인트: Spl.No.별 마지막 비제외 주입
            var stdPoints = new Dictionary<int, (double conc, double area)>();

            for (int i = dataIdx + 2; i < lines.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = lines[i].Split('\t');
                if (idxName >= cols.Length) continue;

                var typeStr = idxType >= 0 && idxType < cols.Length ? cols[idxType].Trim() : "";
                var excluded = idxExcl >= 0 && idxExcl < cols.Length ? cols[idxExcl].Trim() : "0";

                // Standard 행 → 검량선 데이터 수집
                if (typeStr.Equals("Standard", StringComparison.OrdinalIgnoreCase))
                {
                    if (excluded == "1") continue;
                    var stdAreaStr = idxArea >= 0 && idxArea < cols.Length ? cols[idxArea].Trim() : "";
                    var stdConcStr = idxConc >= 0 && idxConc < cols.Length ? cols[idxConc].Trim() : "";
                    int splNo = 0;
                    if (idxSplNo >= 0 && idxSplNo < cols.Length)
                        int.TryParse(cols[idxSplNo].Trim(), out splNo);

                    if (double.TryParse(stdConcStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var sc) &&
                        double.TryParse(stdAreaStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var sa))
                    {
                        stdPoints[splNo] = (sc, sa); // 같은 Spl.No.면 마지막 주입 사용
                    }
                    continue;
                }

                if (!typeStr.Equals("Unknown", StringComparison.OrdinalIgnoreCase)) continue;
                if (excluded == "1") continue;

                var sampleName = cols[idxName].Trim();
                var area = idxArea >= 0 && idxArea < cols.Length ? cols[idxArea].Trim() : "";
                var conc = idxConc >= 0 && idxConc < cols.Length ? cols[idxConc].Trim() : "";

                if (!string.IsNullOrEmpty(conc))
                    groups[sampleName] = (area, conc);
            }

            foreach (var (name, (area, conc)) in groups)
            {
                result.Add(new TocInstrumentRow
                {
                    RawName   = name,
                    SN        = NormalizeSN(name),
                    Area      = area,
                    Conc      = conc,
                    Method    = fileMethod,
                    IsControl = IsControlName(name),
                });
            }

            // Standard 포인트로 검량선 계산
            if (stdPoints.Count >= 2)
            {
                var sortedPts = stdPoints.OrderBy(p => p.Key).ToList();
                var concs = sortedPts.Select(p => p.Value.conc).ToList();
                var areas = sortedPts.Select(p => p.Value.area).ToList();
                var (a, b, r2) = LinearRegression(concs, areas);
                cal = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = fileMethod,
                    StdConcs     = concs.Select(c => c.ToString("F2")).ToArray(),
                    StdAreas     = areas.Select(v => v.ToString("F4")).ToArray(),
                };
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[TocParser/ShimadzuDetail] {ex.Message}"); }
        return result;
    }

    // ── SN 정규화 ──────────────────────────────────────────────────────────

    /// <summary>
    /// 3_11_4 / 03_11_04 / 03-11-04 / [세풍]03-11-04 → 03-11-04 (또는 [세풍]03-11-04)
    /// 매칭 불가 시 "" 반환
    /// </summary>
    public static string NormalizeSN(string raw)
    {
        var trimmed = raw.Trim();
        string prefix = "";

        // 대괄호 접두사 처리 ([세풍], [율촌] 등)
        var bracketM = Regex.Match(trimmed, @"^\[([^\]]+)\](.+)$");
        if (bracketM.Success)
        {
            prefix  = $"[{bracketM.Groups[1].Value}]";
            trimmed = bracketM.Groups[2].Value.Trim();
        }

        // MM-DD-NN 또는 MM_DD_NN 패턴 (각 부분 1~2자리)
        var m = Regex.Match(trimmed, @"^(\d{1,2})[-_](\d{1,2})[-_](\d{1,2})$");
        if (!m.Success) return "";

        return prefix
             + $"{int.Parse(m.Groups[1].Value):D2}-"
             + $"{int.Parse(m.Groups[2].Value):D2}-"
             + $"{int.Parse(m.Groups[3].Value):D2}";
    }

    // ── 예나 CSV (검량선 + 시료 데이터) ──────────────────────────────────────
    // 형식: 시료명;농도(mg/l);Volume;평균면적;주입1;주입2;...
    // IC_xx → IC 검량선, TC_xx → TC 검량선, 그 외 → 시료 데이터

    static (List<TocInstrumentRow> rows, TocCalibrationData cal) ParseJena(string path)
    {
        var cal    = new TocCalibrationData();
        var rows   = new List<TocInstrumentRow>();
        try
        {
            string[] allLines;
            try   { allLines = File.ReadAllLines(path, Encoding.GetEncoding(949)); }
            catch { allLines = File.ReadAllLines(path, Encoding.UTF8); }

            var tcConcs = new List<double>();
            var tcAreas = new List<double>();
            var icConcs = new List<double>();
            var icAreas = new List<double>();

            foreach (var line in allLines.Skip(2))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var cols = line.Split(';');
                if (cols.Length < 4) continue;

                var name = cols[0].Trim();
                if (string.IsNullOrEmpty(name)) continue;

                bool parseConc = double.TryParse(cols[1].Trim(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var conc);
                bool parseArea = double.TryParse(cols[3].Trim(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var area);

                if (name.StartsWith("IC_", StringComparison.OrdinalIgnoreCase))
                {
                    if (parseConc && parseArea) { icConcs.Add(conc); icAreas.Add(area); }
                }
                else if (name.StartsWith("TC_", StringComparison.OrdinalIgnoreCase))
                {
                    if (parseConc && parseArea) { tcConcs.Add(conc); tcAreas.Add(area); }
                }
                else if (!name.StartsWith("Cal_", StringComparison.OrdinalIgnoreCase)
                      && !IsControlName(name))
                {
                    // 시료 데이터 행
                    rows.Add(new TocInstrumentRow
                    {
                        RawName   = name,
                        SN        = NormalizeSN(name),
                        Area      = parseArea ? area.ToString("F4") : cols[3].Trim(),
                        Conc      = parseConc ? conc.ToString("F4") : cols[1].Trim(),
                        Method    = "NPOC", // 아래에서 확정
                        IsControl = false,
                    });
                }
            }

            bool hasTC = tcConcs.Count >= 2;
            bool hasIC = icConcs.Count >= 2;

            if (hasTC)
            {
                var (a, b, r2) = LinearRegression(tcConcs, tcAreas);
                cal.Slope_TC     = a.ToString("F4");
                cal.Intercept_TC = b.ToString("F4");
                cal.R2_TC        = r2.ToString("F5");
                cal.StdConcs     = tcConcs.Select(c => c.ToString("F2")).ToArray();
                cal.StdAreas     = tcAreas.Select(v => v.ToString("F4")).ToArray();
            }
            if (hasIC)
            {
                var (a, b, r2) = LinearRegression(icConcs, icAreas);
                cal.Slope_IC     = a.ToString("F4");
                cal.Intercept_IC = b.ToString("F4");
                cal.R2_IC        = r2.ToString("F5");
                cal.StdConcs_IC  = [.. icConcs.Select(c => c.ToString("F2"))];
                cal.StdAreas_IC  = [.. icAreas.Select(v => v.ToString("F4"))];
            }

            cal.Method = (hasTC && hasIC) ? "TCIC" : "NPOC";
            // 시료 행 메서드 업데이트
            foreach (var r in rows) r.Method = cal.Method;
        }
        catch (Exception ex) { Debug.WriteLine($"[TocParser/Jena] {ex.Message}"); }
        return (rows, cal);
    }

    // ── TOC(TCIC) NOTE 레거시 xlsx ────────────────────────────────────────
    // 양식: 1개 시트, 좌(col 1-8) / 우(col 9-16) 2블록
    //   R5: 'Standard (TC)' | conc1..conc5 | k1(slope) | k0(intercept)
    //   R6: 'AU'            | area1..area5 | R²
    //   R7: 'Standard (IC)' | conc1..conc5 | k1        | k0
    //   R8: 'AU'            | area1..area5 | R²
    //   R9: 헤더 (시료명/시료량/TC/IC/TC-IC/희석배수/Result/Remark)
    //   R10~: 시료 데이터 (시료량 무시, Result=col7, SN=Remark/col8)

    static (List<TocInstrumentRow> rows, TocCalibrationData? cal) ParseTcicXlsx(string path)
    {
        var rows = new List<TocInstrumentRow>();
        TocCalibrationData? cal = null;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            var ws = wb.Worksheets.FirstOrDefault();
            if (ws == null) return (rows, null);

            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 9;

            string slopeTC = "", interTC = "", r2TC = "";
            string slopeIC = "", interIC = "", r2IC = "";
            var stdTcConcs = new List<double>();
            var stdTcAreas = new List<double>();
            var stdIcConcs = new List<double>();
            var stdIcAreas = new List<double>();

            // 좌(offset 0)/우(offset 8) 두 블록 처리
            foreach (int off in new[] { 0, 8 })
            {
                var r5tag = ws.Cell(5, 1 + off).GetString().Trim();
                var r6tag = ws.Cell(6, 1 + off).GetString().Trim();
                var r7tag = ws.Cell(7, 1 + off).GetString().Trim();
                // 블록 시그니처 검증 (빈 우블록 스킵)
                if (r5tag.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) < 0 ||
                    !r6tag.Equals("AU", StringComparison.OrdinalIgnoreCase) ||
                    r7tag.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                // ─ TC 검량선 (R5/R6, col 2~6) ─
                var tcC = new List<double>();
                var tcA = new List<double>();
                for (int c = 2; c <= 6; c++)
                {
                    var cs = ws.Cell(5, c + off).GetString().Trim();
                    var asv = ws.Cell(6, c + off).GetString().Trim();
                    if (double.TryParse(cs, NumberStyles.Any, CultureInfo.InvariantCulture, out var cv) &&
                        double.TryParse(asv, NumberStyles.Any, CultureInfo.InvariantCulture, out var av))
                    {
                        tcC.Add(cv);
                        tcA.Add(av);
                    }
                }
                var bTcSlope = ws.Cell(5, 7 + off).GetString().Trim();
                var bTcInter = ws.Cell(5, 8 + off).GetString().Trim();
                var bTcR2    = ws.Cell(6, 7 + off).GetString().Trim();

                // ─ IC 검량선 (R7/R8, col 2~6) ─
                var icC = new List<double>();
                var icA = new List<double>();
                for (int c = 2; c <= 6; c++)
                {
                    var cs = ws.Cell(7, c + off).GetString().Trim();
                    var asv = ws.Cell(8, c + off).GetString().Trim();
                    if (double.TryParse(cs, NumberStyles.Any, CultureInfo.InvariantCulture, out var cv) &&
                        double.TryParse(asv, NumberStyles.Any, CultureInfo.InvariantCulture, out var av))
                    {
                        icC.Add(cv);
                        icA.Add(av);
                    }
                }
                var bIcSlope = ws.Cell(7, 7 + off).GetString().Trim();
                var bIcInter = ws.Cell(7, 8 + off).GetString().Trim();
                var bIcR2    = ws.Cell(8, 7 + off).GetString().Trim();

                // 첫 번째로 데이터 있는 블록의 검량선 채택
                if (string.IsNullOrEmpty(slopeTC) && !string.IsNullOrEmpty(bTcSlope))
                {
                    slopeTC = bTcSlope; interTC = bTcInter; r2TC = bTcR2;
                    stdTcConcs = tcC; stdTcAreas = tcA;
                }
                if (string.IsNullOrEmpty(slopeIC) && !string.IsNullOrEmpty(bIcSlope))
                {
                    slopeIC = bIcSlope; interIC = bIcInter; r2IC = bIcR2;
                    stdIcConcs = icC; stdIcAreas = icA;
                }

                // ─ 시료 데이터 (R10~) ─
                //   col 1=시료명 / 2=시료량(무시) / 3=TC(mg/L) / 4=IC(mg/L) /
                //   5=TC-IC(mg/L) / 6=희석배수 / 7=Result(㎎/L) / 8=Remark
                for (int r = 10; r <= lastRow; r++)
                {
                    var name   = ws.Cell(r, 1 + off).GetString().Trim();
                    if (string.IsNullOrEmpty(name)) continue;

                    var tcStr     = ws.Cell(r, 3 + off).GetString().Trim();
                    var icStr     = ws.Cell(r, 4 + off).GetString().Trim();
                    var dilStr    = ws.Cell(r, 6 + off).GetString().Trim();
                    var resultStr = ws.Cell(r, 7 + off).GetString().Trim();
                    var remark    = ws.Cell(r, 8 + off).GetString().Trim();
                    if (string.IsNullOrEmpty(resultStr) && string.IsNullOrEmpty(remark)) continue;

                    bool isCtrl = IsControlName(name);
                    var sn = NormalizeSN(remark);
                    if (string.IsNullOrEmpty(sn)) sn = NormalizeSN(name);

                    rows.Add(new TocInstrumentRow
                    {
                        RawName   = name,
                        SN        = sn,
                        Area      = "", // 면적 없음 (xlsx 양식엔 raw AU 없음)
                        Conc      = resultStr,
                        Method    = "TCIC",
                        IsControl = isCtrl,
                        TCAU      = "", // xlsx에 raw AU 없음
                        TCcon     = tcStr,
                        ICAU      = "",
                        ICcon     = icStr,
                        Dilution  = dilStr,
                    });
                }
            }

            if (!string.IsNullOrEmpty(slopeTC) || !string.IsNullOrEmpty(slopeIC))
            {
                cal = new TocCalibrationData
                {
                    Slope_TC     = slopeTC,
                    Intercept_TC = interTC,
                    R2_TC        = r2TC,
                    Slope_IC     = slopeIC,
                    Intercept_IC = interIC,
                    R2_IC        = r2IC,
                    Method       = "TCIC",
                    StdConcs     = [.. stdTcConcs.Select(c => c.ToString("F2", CultureInfo.InvariantCulture))],
                    StdAreas     = [.. stdTcAreas.Select(v => v.ToString("F4", CultureInfo.InvariantCulture))],
                    StdConcs_IC  = [.. stdIcConcs.Select(c => c.ToString("F2", CultureInfo.InvariantCulture))],
                    StdAreas_IC  = [.. stdIcAreas.Select(v => v.ToString("F4", CultureInfo.InvariantCulture))],
                };
            }

            Debug.WriteLine($"[TocParser/TcicXlsxNote] 파싱 완료: rows={rows.Count}, cal={cal?.Method}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TocParser/TcicXlsxNote] 파싱 오류: {ex.Message}");
        }
        return (rows, cal);
    }

    /// <summary>레거시 NOTE xlsx에서 R1B의 분석일을 추출 (없으면 null)</summary>
    public static string? ExtractTcicXlsxDate(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            var ws = wb.Worksheets.FirstOrDefault();
            if (ws == null) return null;
            var cell = ws.Cell(1, 2);
            if (cell.TryGetValue<DateTime>(out var dt))
                return dt.ToString("yyyy-MM-dd");
            var s = cell.GetString().Trim();
            if (DateTime.TryParse(s, out dt))
                return dt.ToString("yyyy-MM-dd");
        }
        catch { }
        return null;
    }

    /// <summary>최소자승 선형 회귀: x=농도, y=면적 → (기울기a, 절편b, R²)</summary>
    static (double a, double b, double r2) LinearRegression(List<double> xs, List<double> ys)
    {
        int n = xs.Count;
        double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
        for (int i = 0; i < n; i++)
        { sumX += xs[i]; sumY += ys[i]; sumXY += xs[i]*ys[i]; sumX2 += xs[i]*xs[i]; }
        double denom = n * sumX2 - sumX * sumX;
        if (Math.Abs(denom) < 1e-12) return (0, sumY / n, 0);
        double a = (n * sumXY - sumX * sumY) / denom;
        double b = (sumY - a * sumX) / n;
        // R² 계산
        double meanY = sumY / n, ssTot = 0, ssRes = 0;
        for (int i = 0; i < n; i++)
        { double diff = ys[i] - meanY; ssTot += diff*diff; double res = ys[i] - (a*xs[i]+b); ssRes += res*res; }
        double r2 = ssTot < 1e-12 ? 1.0 : 1.0 - ssRes / ssTot;
        return (a, b, r2);
    }

    // ── 유틸리티 ──────────────────────────────────────────────────────────

    static bool IsControlName(string name)
    {
        var upper = name.Trim().ToUpperInvariant();
        return ControlKeywords.Any(k => upper.StartsWith(k) || upper.Equals(k));
    }

    static List<string> ReadShimadzuLines(string path)
    {
        // CP949(한국어 시마즈 파일) 우선 — Program.cs에서 RegisterProvider 등록 후 사용 가능
        try   { return File.ReadAllLines(path, Encoding.GetEncoding(949)).ToList(); }
        catch { }
        try   { return File.ReadAllLines(path, Encoding.UTF8).ToList(); }
        catch { return File.ReadAllLines(path, Encoding.Latin1).ToList(); }
    }

    static int IndexOf(string[] arr, params string[] candidates)
    {
        foreach (var c in candidates)
        {
            int i = Array.FindIndex(arr, h => h.Trim().Equals(c, StringComparison.OrdinalIgnoreCase));
            if (i >= 0) return i;
        }
        return -1;
    }

    /// <summary>간단한 CSV 분할 (큰따옴표 지원)</summary>
    static List<string> SplitCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuote = false;
        var current = new System.Text.StringBuilder();
        foreach (char c in line)
        {
            if (c == '"') { inQuote = !inQuote; }
            else if (c == ',' && !inQuote) { result.Add(current.ToString()); current.Clear(); }
            else current.Append(c);
        }
        result.Add(current.ToString());
        return result;
    }

    // PdfPig가 추출한 시료명의 불필요한 공백을 정리
    // "4 단계농축기" → "4단계농축기", "( 기존 )" → "(기존)", "1202_ 금호 1" → "1202_금호1"
    private static string NormalizeJenaName(string raw)
    {
        var s = Regex.Replace(raw.Trim(), @"\s+", " ");
        // 괄호 주변 공백 제거
        s = Regex.Replace(s, @"\(\s*", "(");
        s = Regex.Replace(s, @"\s*\)", ")");
        // 숫자/언더바와 한글 사이 공백 제거
        s = Regex.Replace(s, @"([\d_])\s+([가-힣])", "$1$2");
        s = Regex.Replace(s, @"([가-힣])\s+([\d_])", "$1$2");
        return s;
    }

    // 값 토큰 → mg/l 실수 변환. "245.6µg/l" → 0.2456, "1.43g/l" → 1430
    private static double? ParseJenaConcValue(string num, string unit)
    {
        if (!double.TryParse(num,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (unit.Equals("mg", StringComparison.OrdinalIgnoreCase)) return v;
        if (unit.Equals("g", StringComparison.OrdinalIgnoreCase)) return v * 1000.0;
        // µg / μg / ug
        return v / 1000.0;
    }

    // "1.166E4AU", "9,204AU", "0!" → 실수 AU 값
    private static double? ParseJenaAreaValue(string token)
    {
        if (token.StartsWith("0!")) return 0.0;
        var cleaned = token.Replace("AU", "", StringComparison.OrdinalIgnoreCase)
                            .Replace(",", "");
        return double.TryParse(cleaned,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    // ── 예나 PDF 분석표 파서 (NPOC 직독법 / 가감법 TCIC 공통) ─────────────────
    // PDF 구조 (예나 multi N/C 3100):
    //
    //  [NPOC 직독법]
    //   시료ID 행:  "{시료명} 2026-04-02 오전 11:16:10 4"
    //   매소드 행:  "NPOC(처리장시료) NPOC 200.0µl 1 (1) 1 안에서 1 시료"
    //   NPOC 행:    "NPOC 45.29mg/l 1.166E4AU c = ( 7.686E-4*I + 0.098505)/V 1"
    //
    //  [가감법 TCIC]
    //   시료ID 행:  "{시료명} 2025-12-11 오전 10:28:50 6"
    //   매소드 행:  "가감법(시료용) IC; TC 200.0µl 2 (2-3) 1 안에서 5 시료"  ← 끝의 "5"가 DF
    //   TOC 행:     "TOC 158.8mg/l"                          (계산값, SD/면적 없음)
    //   IC 행:      "IC 424.7mg/l ± 1.01% 2.806E4AU c = ( 6.049E-4*I + 9.690E-3)/V 1"
    //   TC 행:      "TC 583.5mg/l ± 0.29% 3.561E4AU c = ( 6.453E-4*I + 0.36166)/V 1"
    //
    //  [검량선 보고서]
    //   - NPOC 모드: 1페이지 (보정 채널: NPOC)
    //   - 가감법 모드: 2페이지 (보정 채널: IC / TC 각각 독립 검정곡선)
    //   각 채널마다 k0/k1 및 5개 표준점 (0, 5, 10, 50, 100 mg/l 등)
    private static (List<TocInstrumentRow> rows, TocCalibrationData? cal) ParseJenaPdf(string path)
    {
        var rows = new List<TocInstrumentRow>();
        TocCalibrationData? cal = null;
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            Debug.WriteLine($"[JenaPdf] 페이지 수: {doc.NumberOfPages}");

            // 페이지별로 줄 단위 추출 (페이지 경계가 검량선 섹션 분리에 필요)
            var pageLines = new List<List<string>>();
            int pageIdx = 0;
            foreach (var page in doc.GetPages())
            {
                pageIdx++;
                var words = page.GetWords().ToList();
                Debug.WriteLine($"[JenaPdf] 페이지 {pageIdx}: 단어 {words.Count}개");

                var lineGroups = new List<List<UglyToad.PdfPig.Content.Word>>();
                foreach (var w in words)
                {
                    var grp = lineGroups.FirstOrDefault(g =>
                        Math.Abs(g[0].BoundingBox.Bottom - w.BoundingBox.Bottom) < 4);
                    if (grp != null) grp.Add(w);
                    else lineGroups.Add(new List<UglyToad.PdfPig.Content.Word> { w });
                }
                lineGroups = lineGroups.OrderByDescending(g => g[0].BoundingBox.Bottom).ToList();
                var lines = new List<string>();
                foreach (var grp in lineGroups)
                {
                    var sorted = grp.OrderBy(w => w.BoundingBox.Left);
                    var line = string.Join(" ", sorted.Select(w => w.Text)).Trim();
                    if (!string.IsNullOrWhiteSpace(line)) lines.Add(line);
                }
                pageLines.Add(lines);
            }
            var allLines = pageLines.SelectMany(p => p).ToList();
            Debug.WriteLine($"[JenaPdf] 총 줄 수: {allLines.Count}");

            // ───── 검량선 파싱 (페이지별, 채널별) ─────
            // "보정 채널 : IC" 또는 "보정 채널 : TC" 또는 "보정 채널 : NPOC"
            var channelRe = new Regex(@"보정\s*채널\s*:\s*(\S+)");
            // "k0 = 0.098505 k1 = 7.686E-4"
            var kLineRe = new Regex(@"k0\s*=\s*([0-9eE.+\-]+)\s+k1\s*=\s*([0-9eE.+\-]+)");
            // "1 1-1 0.000mg/l 0µg 0! 0.197mg/l"
            var stdRe = new Regex(
                @"^(\d+)\s+\d+-\d+\s+([0-9]+(?:\.[0-9]+)?)\s*mg/l\s+[0-9]+(?:\.[0-9]+)?\s*[\u00B5\u03BC]g\s+(\S+)");

            var calByCh = new Dictionary<string, (string k0, string k1, List<double> concs, List<double> areas)>();
            foreach (var lines in pageLines)
            {
                string currentCh = "";
                foreach (var line in lines)
                {
                    var cm = channelRe.Match(line);
                    if (cm.Success)
                    {
                        currentCh = cm.Groups[1].Value;
                        if (!calByCh.ContainsKey(currentCh))
                            calByCh[currentCh] = ("", "", new List<double>(), new List<double>());
                        continue;
                    }
                    if (string.IsNullOrEmpty(currentCh)) continue;

                    var km = kLineRe.Match(line);
                    if (km.Success)
                    {
                        var cur = calByCh[currentCh];
                        calByCh[currentCh] = (km.Groups[1].Value, km.Groups[2].Value, cur.concs, cur.areas);
                        continue;
                    }

                    var sm = stdRe.Match(line);
                    if (sm.Success &&
                        double.TryParse(sm.Groups[2].Value,
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var concVal))
                    {
                        var areaVal = ParseJenaAreaValue(sm.Groups[3].Value) ?? 0.0;
                        calByCh[currentCh].concs.Add(concVal);
                        calByCh[currentCh].areas.Add(areaVal);
                    }
                }
            }

            // 검량선 데이터 조립
            bool hasIC = calByCh.ContainsKey("IC") && calByCh["IC"].concs.Count > 0;
            bool hasTC = calByCh.ContainsKey("TC") && calByCh["TC"].concs.Count > 0;
            bool hasNpoc = calByCh.ContainsKey("NPOC") && calByCh["NPOC"].concs.Count > 0;

            if (hasIC || hasTC || hasNpoc)
            {
                cal = new TocCalibrationData();
                if (hasIC && hasTC)
                {
                    // 가감법: IC/TC 각각의 검정곡선
                    cal.Method = "TCIC";
                    var ic = calByCh["IC"];
                    var tc = calByCh["TC"];
                    cal.Slope_IC = ic.k1;
                    cal.Intercept_IC = ic.k0;
                    cal.StdConcs_IC = ic.concs.Select(c => c.ToString("F2",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.StdAreas_IC = ic.areas.Select(v => v.ToString("F4",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.Slope_TC = tc.k1;
                    cal.Intercept_TC = tc.k0;
                    cal.StdConcs = tc.concs.Select(c => c.ToString("F2",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.StdAreas = tc.areas.Select(v => v.ToString("F4",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                }
                else
                {
                    // NPOC 직독법
                    cal.Method = "NPOC";
                    var npoc = hasNpoc ? calByCh["NPOC"] : (hasTC ? calByCh["TC"] : calByCh["IC"]);
                    cal.Slope_TC = npoc.k1;
                    cal.Intercept_TC = npoc.k0;
                    cal.StdConcs = npoc.concs.Select(c => c.ToString("F2",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.StdAreas = npoc.areas.Select(v => v.ToString("F4",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                }
            }

            // ───── 시료 파싱 ─────
            // 시료 ID 행: "{name} YYYY-MM-DD 오전/오후 HH:MM:SS {rack#}"
            var sampleIdRe = new Regex(
                @"^(?<name>.+?)\s+\d{4}-\d{2}-\d{2}\s+(?:오전|오후)\s+\d{1,2}:\d{2}:\d{2}\s+\d+\s*$");

            // 값 행 패턴
            // "TOC 158.8mg/l" (가감법 계산값)
            var tocLineRe = new Regex(
                @"^TOC\s+([0-9]+(?:\.[0-9]+)?)\s*(mg|g|[\u00B5\u03BC]g|ug)\s*/\s*l",
                RegexOptions.IgnoreCase);
            // "IC 424.7mg/l ± 1.01% 2.806E4AU ..." / "TC 583.5mg/l ..."
            var icTcLineRe = new Regex(
                @"^(IC|TC)\s+([0-9]+(?:\.[0-9]+)?)\s*(mg|g|[\u00B5\u03BC]g|ug)\s*/\s*l",
                RegexOptions.IgnoreCase);
            // "NPOC 45.29mg/l ..." (NPOC 직독법)
            var npocValueRe = new Regex(
                @"^NPOC\s+([0-9]+(?:\.[0-9]+)?)\s*(mg|g|[\u00B5\u03BC]g|ug)\s*/\s*l",
                RegexOptions.IgnoreCase);
            // 매소드 행에서 희석배수 추출: "... 안에서 5 시료"
            var dfRe = new Regex(@"안에서\s+(\d+)\s+시료");
            // AU 면적 추출
            var auRe = new Regex(
                @"([0-9]+(?:[.,][0-9]+)?(?:E[+\-]?[0-9]+)?)\s*AU",
                RegexOptions.IgnoreCase);

            for (int i = 0; i < allLines.Count; i++)
            {
                var m = sampleIdRe.Match(allLines[i]);
                if (!m.Success) continue;

                var name = NormalizeJenaName(m.Groups["name"].Value);

                // 바탕/정도관리 시료 스킵
                if (name.Equals("BK", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("CCV", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("DW", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("MBK", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("FBK", StringComparison.OrdinalIgnoreCase))
                    continue;

                // 다음 6줄 이내에서 값 라인 탐색 (다음 시료 ID 만나면 중단)
                double? tocVal = null;
                double? tcArea = null;
                bool isTcic = false;

                for (int j = i + 1; j < Math.Min(i + 6, allLines.Count); j++)
                {
                    if (sampleIdRe.IsMatch(allLines[j])) break;

                    // 가감법 TOC 라인 (계산값)
                    var tm = tocLineRe.Match(allLines[j]);
                    if (tm.Success)
                    {
                        isTcic = true;
                        tocVal = ParseJenaConcValue(tm.Groups[1].Value, tm.Groups[2].Value);
                        continue;
                    }

                    // 가감법 IC/TC 라인
                    var icm = icTcLineRe.Match(allLines[j]);
                    if (icm.Success)
                    {
                        var isTC = icm.Groups[1].Value.Equals("TC", StringComparison.OrdinalIgnoreCase);
                        if (isTC)
                        {
                            var am = auRe.Match(allLines[j]);
                            if (am.Success) tcArea = ParseJenaAreaValue(am.Groups[1].Value + "AU");
                            break; // TC가 시료당 마지막 값 라인
                        }
                        continue;
                    }

                    // NPOC 직독법
                    var nm = npocValueRe.Match(allLines[j]);
                    if (nm.Success)
                    {
                        tocVal = ParseJenaConcValue(nm.Groups[1].Value, nm.Groups[2].Value);
                        var am = auRe.Match(allLines[j]);
                        if (am.Success) tcArea = ParseJenaAreaValue(am.Groups[1].Value + "AU");
                        break;
                    }
                }

                if (tocVal == null) continue;

                rows.Add(new TocInstrumentRow
                {
                    RawName = name,
                    SN = name,
                    Conc = tocVal.Value.ToString("0.###",
                        System.Globalization.CultureInfo.InvariantCulture),
                    Area = tcArea?.ToString("0.####",
                        System.Globalization.CultureInfo.InvariantCulture) ?? "",
                    Method = isTcic ? "TCIC" : "NPOC",
                });
            }

            // 검량선의 메서드를 시료에 맞춰 보정 (혼합 상황 방지)
            if (cal != null && rows.Count > 0)
            {
                bool anyTcic = rows.Any(r => r.Method == "TCIC");
                bool anyNpoc = rows.Any(r => r.Method == "NPOC");
                if (anyTcic && !anyNpoc) cal.Method = "TCIC";
                else if (anyNpoc && !anyTcic) cal.Method = "NPOC";
            }

            Debug.WriteLine($"[JenaPdf] 파싱 완료: 검량선={cal?.Method}, 시료={rows.Count}건");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[JenaPdf] 파싱 오류: {ex.Message}");
        }
        return (rows, cal);
    }
}
