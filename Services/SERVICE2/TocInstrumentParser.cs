using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE2;

public class TocInstrumentRow
{
    public string RawName  { get; set; } = "";
    public string SN       { get; set; } = ""; // 정규화된 SN (03-11-04 형식)
    public string Area     { get; set; } = "";
    public string Conc     { get; set; } = ""; // 최종 농도 (mg/L)
    public string Method   { get; set; } = "NPOC"; // NPOC 또는 TCIC
    public bool   IsControl { get; set; } = false;
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

    public enum TocFileFormat { Unknown, Skalar, ShimadzuSimple, ShimadzuDetail, JenaCalibration }

    public static TocFileFormat DetectFormat(string path)
    {
        try
        {
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
}
