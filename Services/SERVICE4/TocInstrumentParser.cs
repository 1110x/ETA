using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ETA.Services.SERVICE4;

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
    // 기기출력값 (Shimadzu PDF에서 직접 읽은 slope/intercept/R² + Mean Area)
    // null이면 기기출력값 없음 (TXT only)
    public string? Slope_TC_Inst     { get; set; }
    public string? Intercept_TC_Inst { get; set; }
    public string? R2_TC_Inst        { get; set; }
    public string? Slope_IC_Inst     { get; set; }
    public string? Intercept_IC_Inst { get; set; }
    public string? R2_IC_Inst        { get; set; }
    public string[] StdAreas_Inst    { get; set; } = []; // TC Mean Area (기기값)
    public string[] StdAreas_IC_Inst { get; set; } = []; // IC Mean Area (기기값)
    public bool HasInstrumentCal => Slope_TC_Inst != null;
}

/// <summary>TOC 분석기(스칼라/시마즈) 출력 파일 파서</summary>
public static class TocInstrumentParser
{
    // 바탕/정도관리 시료 키워드 (SN 매칭 제외)
    static readonly HashSet<string> ControlKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "DW", "BK", "MBK", "FBK", "CCV", "ST", "NPOC", "TCIC", "BLANK" };

    // ── 포맷 감지 ──────────────────────────────────────────────────────────

    public enum TocFileFormat { Unknown, Skalar, ShimadzuSimple, ShimadzuDetail, ShimadzuPdf, JenaCalibration, JenaAnalysisPdf, TcicXlsxNote, SkalarCalibrationPdf, SkalarNpocPdf }

    public static TocFileFormat DetectFormat(string path)
    {
        try
        {
            // PDF → 내용으로 스칼라 분석 PDF / 스칼라 검량선 전용 PDF / 예나 분석표 구별
            if (path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
                    var pages = doc.GetPages().ToList();
                    var firstText = pages.Count > 0 ? pages[0].Text : "";
                    bool isSkalar = firstText.Contains("메소드 이름") && firstText.Contains("캘리브레이션 타입");
                    if (isSkalar)
                    {
                        // 시료 데이터 페이지 존재 여부 확인 (SI. No. ID 패턴)
                        bool hasSamples = pages.Skip(1).Any(p => p.Text.Contains("SI. No."));
                        return hasSamples ? TocFileFormat.SkalarNpocPdf : TocFileFormat.SkalarCalibrationPdf;
                    }
                    // 시마즈 TOC-L Report PDF: "TOC-Control" 또는 "TOC-L" + "Report" 키워드
                    bool isShimadzu = firstText.Contains("TOC-Control") ||
                                      (firstText.Contains("TOC-L") && firstText.Contains("Report"));
                    if (isShimadzu) return TocFileFormat.ShimadzuPdf;
                }
                catch { }
                return TocFileFormat.JenaAnalysisPdf;
            }

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

            // ── 시마즈 [Header] 구조 우선 확인 (스칼라 패턴보다 먼저) ──
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

            // 예나 검량선 CSV: 첫 줄이 "Cal_"로 시작하고 세미콜론 구분자 (스칼라 패턴보다 먼저)
            if (firstLine.StartsWith("Cal_", StringComparison.OrdinalIgnoreCase) && firstLine.Contains(';'))
                return TocFileFormat.JenaCalibration;

            // 스칼라 CSV 감지: UTF-16 BOM 또는 내용 기반
            var bom = new byte[2];
            using (var fs = File.OpenRead(path))
                fs.Read(bom, 0, 2);
            if (bom[0] == 0xFF && bom[1] == 0xFE)
                return TocFileFormat.Skalar;

            // BOM 없는 경우: 스칼라 고유 패턴으로 감지 (시마즈/예나와 겹치지 않는 패턴만 사용)
            try
            {
                var encodings = new[] { Encoding.UTF8, Encoding.Unicode, Encoding.UTF32, Encoding.ASCII };
                foreach (var encoding in encodings)
                {
                    try
                    {
                        string content;
                        using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var reader = new StreamReader(fileStream, encoding))
                            content = reader.ReadToEnd();

                        // 스칼라 고유 패턴 (시마즈/예나와 겹치지 않는 것만)
                        var patterns = new[]
                        {
                            "스칼라", "가감법",
                            "시료명", "농도", "면적", "타입",
                            ",WS,", "\"WS\"", " WS ,",
                            ",U,", "\"U,\"",
                            "mg/l", "mg/L",
                        };
                        int matchCount = 0;
                        foreach (var p in patterns)
                            if (content.Contains(p) && ++matchCount >= 2)
                                return TocFileFormat.Skalar;
                    }
                    catch { continue; }
                }
            }
            catch { }

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
            case TocFileFormat.ShimadzuPdf:
                (rows, cal) = ParseShimadzuPdf(path);
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
            case TocFileFormat.SkalarCalibrationPdf:
                (rows, cal) = ParseSkalarCalibrationPdf(path);
                break;
            case TocFileFormat.SkalarNpocPdf:
                rows = ParseSkalarNpocPdf(path, out cal);
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

        // 로그 파일 경로
        string logPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Logs", "Users", Environment.UserName, "SCALAR_TCIC.log");
        void WriteLog(string message)
        {
            try
            {
                // 디렉토리 자동 생성
                if (App.EnableLogging)
                {
                    var logDir = System.IO.Path.GetDirectoryName(logPath);
                    if (!Directory.Exists(logDir))
                        Directory.CreateDirectory(logDir);

                    string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [ParseSkalar] {message}";
                    File.AppendAllText(logPath, logLine + "\n");
                }
            }
            catch { }
        }

        try
        {
            WriteLog($"파싱 시작: {System.IO.Path.GetFileName(path)}");

            // 여러 인코딩 시도 (파일 감지와 동일한 방식)
            string text = "";
            var encodings = new[]
            {
                Encoding.Unicode,     // UTF-16 LE (기본)
                Encoding.UTF8,        // UTF-8
                Encoding.GetEncoding(949),  // CP949 (한국어)
                Encoding.UTF32,       // UTF-32
                Encoding.ASCII        // ASCII
            };

            foreach (var encoding in encodings)
            {
                try
                {
                    WriteLog($"{encoding.EncodingName} 인코딩 시도...");
                    // 읽기 전용 공유 모드로 파일 열기 (다른 프로그램에서 열려있어도 읽기 가능)
                    using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(fileStream, encoding))
                    {
                        text = reader.ReadToEnd();
                    }
                    WriteLog($"{encoding.EncodingName} 파일 읽기 완료, 길이: {text.Length}");

                    // 파일 내용 일부 확인 (처음 200자)
                    var preview = text.Length > 200 ? text.Substring(0, 200) : text;
                    WriteLog($"{encoding.EncodingName} 내용 미리보기: {preview.Replace('\n', ' ').Replace('\r', ' ')}");

                    // 스칼라 특징 패턴 확인 (더 관대한 패턴)
                    var patterns = new[] { "mg/l", "Conc", ",U,", "\"U\"", "TC 0", "TC 2", "TC 5", "IC 0", "IC 2", "IC 5", ",S,", "\"S\"" };
                    int matchCount = 0;
                    foreach (var pattern in patterns)
                    {
                        if (text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            matchCount++;
                            WriteLog($"{encoding.EncodingName} 패턴 '{pattern}' 매칭 성공");
                        }
                    }

                    // 스칼라 CSV의 기본 구조 확인 (쉼표 구분, 따옴표 사용)
                    bool hasBasicStructure = text.Contains(",") && text.Contains("\"") &&
                                            (text.Contains("TC") || text.Contains("IC"));

                    if (matchCount > 0 || hasBasicStructure)
                    {
                        WriteLog($"{encoding.EncodingName} 인코딩 성공! 매칭 패턴: {matchCount}개, 기본구조: {hasBasicStructure}");
                        break; // 성공적인 인코딩 발견
                    }
                    else
                    {
                        WriteLog($"{encoding.EncodingName} 패턴 매칭 실패 (0개), 기본구조: {hasBasicStructure}");
                        text = ""; // 실패시 텍스트 초기화
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"{encoding.EncodingName} 인코딩 오류: {ex.Message}");
                    continue; // 해당 인코딩 실패시 다음 시도
                }
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                WriteLog($"모든 인코딩 실패: {path}");
                return result;
            }

            WriteLog("인코딩 성공, CSV 파싱 시작...");

            var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2)
            {
                WriteLog("CSV 라인이 너무 적음");
                return result;
            }

            // 헤더 분석
            var headerFields = SplitCsvLine(lines[0]);
            WriteLog($"헤더 필드 개수: {headerFields.Count}");
            WriteLog($"헤더: {string.Join(" | ", headerFields)}");

            // 한글/영문 헤더 모두 지원하는 컬럼 인덱스 찾기 (위치 기반 추가)
            int idxSerial = FindColumnIndex(headerFields, "시리얼넘버", "Serial", "순번");
            int idxPosition = FindColumnIndex(headerFields, "위치", "Position");
            int idxSampleId = FindColumnIndex(headerFields, "샘플 ID", "Sample ID", "시료명", "Sample Name");
            int idxType = FindColumnIndex(headerFields, "시료 유형", "Sample Type", "타입", "Type");
            int idxArea = FindColumnIndex(headerFields, "평균 피크 면적", "Average Peak Area", "면적", "Area");
            int idxConc = FindColumnIndex(headerFields, "Conc", "농도", "Concentration");
            int idxUnit = FindColumnIndex(headerFields, "측정 단위", "Measurement Unit", "단위", "Unit");

            // 인코딩 문제로 헤더가 깨진 경우 위치 기반 대체 (스칼라 표준 형식)
            if (idxSampleId < 0 && headerFields.Count >= 3) idxSampleId = 2; // 3번째 컬럼
            if (idxType < 0 && headerFields.Count >= 4) idxType = 3; // 4번째 컬럼
            if (idxArea < 0 && headerFields.Count >= 5) idxArea = 4; // 5번째 컬럼
            if (idxConc < 0 && headerFields.Count >= 6) idxConc = 5; // 6번째 컬럼

            WriteLog($"컬럼 인덱스 - 시료명:{idxSampleId}, 타입:{idxType}, 면적:{idxArea}, 농도:{idxConc}");

            if (idxSampleId < 0 || idxType < 0 || idxArea < 0 || idxConc < 0)
            {
                WriteLog("필수 컬럼을 찾을 수 없음");
                return result;
            }

            foreach (var line in lines.Skip(1)) // 헤더 건너뜀
            {
                var fields = SplitCsvLine(line);
                if (fields.Count <= Math.Max(Math.Max(idxSampleId, idxType), Math.Max(idxArea, idxConc)))
                {
                    WriteLog($"필드 개수 부족: {fields.Count}");
                    continue;
                }

                var sampleName = fields[idxSampleId].Trim();
                var typeStr    = fields[idxType].Trim();

                WriteLog($"처리 중: {sampleName}, 타입: {typeStr}");

                // 표준품 처리 (WS 또는 S)
                if (typeStr.Equals("WS", StringComparison.OrdinalIgnoreCase) ||
                    typeStr.Equals("S", StringComparison.OrdinalIgnoreCase))
                {
                    // IC 표준품: "IC 2.5" → 2.5, "IC 0" → 0.0
                    var icMatch = Regex.Match(sampleName, @"IC\s*([\d]+(?:\.[\d]+)?)");
                    if (icMatch.Success &&
                        double.TryParse(icMatch.Groups[1].Value, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var icNomConc) &&
                        double.TryParse(fields[idxArea].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var icArea))
                    {
                        wsNomConcs.Add(icNomConc);
                        wsAreas.Add(icArea);
                        // 로그 제거 - 성능 최적화
                        continue;
                    }

                    // TC 표준품: "TC 5" → 5.0, "TC 0" → 0.0
                    var tcMatch = Regex.Match(sampleName, @"TC\s*([\d]+(?:\.[\d]+)?)");
                    if (tcMatch.Success &&
                        double.TryParse(tcMatch.Groups[1].Value, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var tcNomConc) &&
                        double.TryParse(fields[idxArea].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var tcArea))
                    {
                        wsNomConcs.Add(tcNomConc);
                        wsAreas.Add(tcArea);
                        // 로그 제거 - 성능 최적화
                        continue;
                    }

                    // 기타 표준품: 이름에서 숫자 추출 ("NPOC 5 mg/L" → 5.0, "NPOC 0" → 0.0)
                    var nomMatch = Regex.Match(sampleName, @"([\d]+(?:\.[\d]+)?)(?:\s*mg)?");
                    if (nomMatch.Success &&
                        double.TryParse(nomMatch.Groups[1].Value, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var nomConc) &&
                        double.TryParse(fields[idxArea].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var area))
                    {
                        wsNomConcs.Add(nomConc);
                        wsAreas.Add(area);
                        // 로그 제거 - 성능 최적화
                    }
                    continue;
                }

                // Unknown 시료만 처리 (U 타입)
                if (!typeStr.Equals("U", StringComparison.OrdinalIgnoreCase))
                {
                    WriteLog($"Unknown 타입이 아님: {typeStr}");
                    continue;
                }

                bool isCtrl = IsControlName(sampleName);
                var sn = NormalizeSN(sampleName);

                result.Add(new TocInstrumentRow
                {
                    RawName   = sampleName,
                    SN        = sn,
                    Area      = fields[idxArea].Trim(),
                    Conc      = fields[idxConc].Trim(),
                    Method    = "NPOC",
                    IsControl = isCtrl,
                });
            }

            // 선형 회귀: x=공칭농도, y=면적 → Area = a*Conc + b
            if (wsNomConcs.Count >= 2)
            {
                WriteLog($"검정곡선 생성: 표준품 {wsNomConcs.Count}개점");
                var (a, b, r2) = LinearRegressionZeroShift(wsNomConcs, wsAreas);
                cal = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = "NPOC",
                    StdConcs     = wsNomConcs.Select(c => c.ToString("F2")).ToArray(),
                    StdAreas     = wsAreas.Select(v => v.ToString("F3")).ToArray(),
                };
                WriteLog($"검정곡선 완료: Slope={a:F4}, R²={r2:F5}");
            }
            else
            {
                WriteLog($"검정곡선 생성 실패: 표준품 {wsNomConcs.Count}개점 (최소 2개 필요)");
            }
        }
        catch (Exception ex) { }
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
                var (a, b, r2) = LinearRegressionZeroShift(concs, areas);
                cal = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = method,
                    StdConcs     = concs.Select(c => c.ToString("F2")).ToArray(),
                    StdAreas     = areas.Select(v => v.ToString("F3")).ToArray(),
                };
            }
        }
        catch (Exception ex) { }
        return result;
    }

    // ── 시마즈 상세 (주입 단위) ───────────────────────────────────────────
    // 수정: (시료명, Result문자열) 키로 분석런 그룹화
    //   - 같은 런의 다중 주입 → 비제외 Area 평균, 농도는 Result 컬럼에서 파싱
    //   - 같은 이름의 복수 분석 → Result값이 다르므로 각각 별도 행 생성

    static List<TocInstrumentRow> ParseShimadzuDetail(string path, out TocCalibrationData? cal)
    {
        cal = null;
        var result = new List<TocInstrumentRow>();
        try
        {
            var lines   = ReadShimadzuLines(path);
            int dataIdx = lines.FindIndex(l => l.TrimStart().StartsWith("[Data]"));
            if (dataIdx < 0) return result;

            var header     = lines[dataIdx + 1].Split('\t');
            int idxType    = IndexOf(header, "Type");
            int idxAnal    = IndexOf(header, "Anal.");
            int idxName    = IndexOf(header, "Sample Name");
            int idxSplNo   = IndexOf(header, "Spl. No.", "Spl.No.");
            int idxArea    = IndexOf(header, "Area");
            int idxConc    = IndexOf(header, "Conc.");
            int idxExcl    = IndexOf(header, "Excluded");
            int idxAnalInj = IndexOf(header, "Analysis(Inj.)", "Analysis");
            int idxResCol  = IndexOf(header, "Result");
            int idxDil     = IndexOf(header, "Manual Dilution");

            if (idxName < 0) return result;

            // ── 메서드 감지 ──
            bool hasTc = false, hasIc = false;
            for (int k = dataIdx + 2; k < lines.Count && k < dataIdx + 60; k++)
            {
                if (string.IsNullOrWhiteSpace(lines[k])) continue;
                var c0 = lines[k].Split('\t');
                var t0 = idxType >= 0 && idxType < c0.Length ? c0[idxType].Trim() : "";
                if (!t0.Equals("Standard", StringComparison.OrdinalIgnoreCase)) continue;
                var a0 = idxAnal >= 0 && idxAnal < c0.Length ? c0[idxAnal].Trim().ToUpper() : "";
                if (a0 == "TC") hasTc = true;
                if (a0 == "IC") hasIc = true;
                if (hasTc && hasIc) break;
            }
            string fileMethod = (hasTc && hasIc) ? "TCIC" : "NPOC";

            // ── 검량선 포인트: Spl.No.별 비제외 주입 Area 목록 (다중 주입 평균) ──
            var stdListsTC = new Dictionary<int, (double conc, List<double> areas)>();
            var stdListsIC = new Dictionary<int, (double conc, List<double> areas)>();

            // ── 분석런 그룹화: 키=(시료명, Result문자열) ──
            // 동일 런의 다중 주입은 모두 같은 Result 문자열을 공유
            var runOrder = new List<(string name, string resKey)>(); // 출현 순서
            var tcAreas  = new Dictionary<(string, string), List<double>>(); // 비제외 TC Area 목록
            var icAreas  = new Dictionary<(string, string), List<double>>(); // 비제외 IC Area 목록
            var runMeta  = new Dictionary<(string, string), (string tcConc, string icConc, string tocConc,
                                                              string npocConc, string npocArea, string dil)>();

            // Result 컬럼 파싱 정규식
            var reTocConc  = new Regex(@"TOC:(-?[0-9.]+)ppm",  RegexOptions.IgnoreCase);
            var reTcConc   = new Regex(@"TC:(-?[0-9.]+)ppm",   RegexOptions.IgnoreCase);
            var reIcConc   = new Regex(@"IC:(-?[0-9.]+)ppm",   RegexOptions.IgnoreCase);
            var reNpocConc = new Regex(@"NPOC:(-?[0-9.]+)ppm", RegexOptions.IgnoreCase);

            for (int i = dataIdx + 2; i < lines.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = lines[i].Split('\t');
                if (idxName >= cols.Length) continue;

                var typeStr  = idxType >= 0 && idxType < cols.Length ? cols[idxType].Trim() : "";
                var analStr  = idxAnal >= 0 && idxAnal < cols.Length ? cols[idxAnal].Trim().ToUpper() : "";
                var excluded = idxExcl >= 0 && idxExcl < cols.Length ? cols[idxExcl].Trim() : "0";

                // ── Standard → 검량선 포인트 수집 (제외된 주입 스킵) ──
                if (typeStr.Equals("Standard", StringComparison.OrdinalIgnoreCase))
                {
                    if (excluded == "1") continue;
                    var stdArea = idxArea >= 0 && idxArea < cols.Length ? cols[idxArea].Trim() : "";
                    var stdConc = idxConc >= 0 && idxConc < cols.Length ? cols[idxConc].Trim() : "";
                    int splNo = 0;
                    if (idxSplNo >= 0 && idxSplNo < cols.Length)
                        int.TryParse(cols[idxSplNo].Trim(), out splNo);
                    if (double.TryParse(stdConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var sc) &&
                        double.TryParse(stdArea, NumberStyles.Any, CultureInfo.InvariantCulture, out var sa))
                    {
                        var dict = analStr == "IC" ? stdListsIC : stdListsTC;
                        if (!dict.ContainsKey(splNo)) dict[splNo] = (sc, new List<double>());
                        dict[splNo].areas.Add(sa);
                    }
                    continue;
                }

                if (!typeStr.Equals("Unknown", StringComparison.OrdinalIgnoreCase)) continue;

                var name    = cols[idxName].Trim();
                var area    = idxArea >= 0 && idxArea < cols.Length ? cols[idxArea].Trim() : "";
                var conc    = idxConc >= 0 && idxConc < cols.Length ? cols[idxConc].Trim() : "";
                var resStr  = idxResCol >= 0 && idxResCol < cols.Length ? cols[idxResCol].Trim() : "";
                var dilStr  = idxDil >= 0 && idxDil < cols.Length ? cols[idxDil].Trim() : "1";
                var injType = idxAnalInj >= 0 && idxAnalInj < cols.Length
                    ? cols[idxAnalInj].Trim().ToUpper() : analStr;

                // 런 키: (시료명, Result문자열) — 동일 런은 Result 값이 같음
                var runKey = (name, resStr);
                if (!runMeta.ContainsKey(runKey))
                {
                    // 메타 초기화: Result 컬럼에서 농도 파싱
                    string tcC = "", icC = "", tocC = "", npocC = "", npocA = "";
                    if (fileMethod == "TCIC")
                    {
                        var mToc = reTocConc.Match(resStr);
                        var mTc  = reTcConc.Match(resStr);
                        var mIc  = reIcConc.Match(resStr);
                        tocC = mToc.Success ? mToc.Groups[1].Value : "";
                        tcC  = mTc.Success  ? mTc.Groups[1].Value  : "";
                        icC  = mIc.Success  ? mIc.Groups[1].Value  : "";
                    }
                    else
                    {
                        // NPOC: Result="NPOC:x.xxxppm" 또는 Conc. 직접 사용
                        var mNpoc = reNpocConc.Match(resStr);
                        npocC = mNpoc.Success ? mNpoc.Groups[1].Value : conc;
                        npocA = area; // 첫 주입 Area (보통 NPOC는 1회 주입)
                    }
                    runMeta[runKey] = (tcC, icC, tocC, npocC, npocA, dilStr);
                    tcAreas[runKey] = new List<double>();
                    icAreas[runKey] = new List<double>();
                    runOrder.Add(runKey);
                }

                // 제외된 주입은 Area 누적에서 제외
                if (excluded == "1") continue;

                if (fileMethod == "TCIC")
                {
                    if (injType == "TC" && double.TryParse(area, NumberStyles.Any, CultureInfo.InvariantCulture, out var ta))
                        tcAreas[runKey].Add(ta);
                    else if (injType == "IC" && double.TryParse(area, NumberStyles.Any, CultureInfo.InvariantCulture, out var ia))
                        icAreas[runKey].Add(ia);
                }
                // NPOC는 npocArea를 첫 비제외 주입값으로 갱신
                else if (double.TryParse(area, NumberStyles.Any, CultureInfo.InvariantCulture, out var na))
                {
                    var m = runMeta[runKey];
                    runMeta[runKey] = (m.tcConc, m.icConc, m.tocConc, m.npocConc, area, m.dil);
                }
            }

            // ── 분석런별 행 생성 ──
            foreach (var key in runOrder)
            {
                var (name, _) = key;
                var meta = runMeta[key];

                if (fileMethod == "TCIC")
                {
                    // TC/IC 평균 Area
                    var tcList = tcAreas[key];
                    var icList = icAreas[key];
                    string tcAreaStr = tcList.Count > 0
                        ? (tcList.Sum() / tcList.Count).ToString("F3", CultureInfo.InvariantCulture) : "";
                    string icAreaStr = icList.Count > 0
                        ? (icList.Sum() / icList.Count).ToString("F3", CultureInfo.InvariantCulture) : "";

                    // 최종 TOC: Result에서 추출, 없으면 TC-IC 계산
                    string finalConc = meta.tocConc;
                    if (string.IsNullOrEmpty(finalConc) &&
                        double.TryParse(meta.tcConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var tv) &&
                        double.TryParse(meta.icConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var iv))
                        finalConc = (tv - iv).ToString("F4", CultureInfo.InvariantCulture);

                    result.Add(new TocInstrumentRow
                    {
                        RawName   = name,
                        SN        = NormalizeSN(name),
                        Area      = tcAreaStr,
                        Conc      = finalConc,
                        TCAU      = tcAreaStr,
                        TCcon     = meta.tcConc,
                        ICAU      = icAreaStr,
                        ICcon     = meta.icConc,
                        Method    = "TCIC",
                        IsControl = IsControlName(name),
                        Dilution  = meta.dil != "1" ? meta.dil : "",
                    });
                }
                else
                {
                    result.Add(new TocInstrumentRow
                    {
                        RawName   = name,
                        SN        = NormalizeSN(name),
                        Area      = meta.npocArea,
                        Conc      = meta.npocConc,
                        Method    = "NPOC",
                        IsControl = IsControlName(name),
                        Dilution  = meta.dil != "1" ? meta.dil : "",
                    });
                }
            }

            // Spl.No.별 Area 목록 → 평균값으로 stdPoints 생성
            var stdPointsTC = stdListsTC.ToDictionary(
                p => p.Key,
                p => (conc: p.Value.conc, area: p.Value.areas.Count > 0 ? p.Value.areas.Average() : 0.0));
            var stdPointsIC = stdListsIC.ToDictionary(
                p => p.Key,
                p => (conc: p.Value.conc, area: p.Value.areas.Count > 0 ? p.Value.areas.Average() : 0.0));

            // TC 검량선 계산
            TocCalibrationData? calData = null;
            if (stdPointsTC.Count >= 2)
            {
                var pts = stdPointsTC.OrderBy(p => p.Key).ToList();
                var (a, b, r2) = LinearRegressionZeroShift(pts.Select(p => p.Value.conc).ToList(),
                                                            pts.Select(p => p.Value.area).ToList());
                calData = new TocCalibrationData
                {
                    Slope_TC     = a.ToString("F4"),
                    Intercept_TC = b.ToString("F4"),
                    R2_TC        = r2.ToString("F5"),
                    Method       = fileMethod,
                    StdConcs     = pts.Select(p => p.Value.conc.ToString("F4")).ToArray(),
                    StdAreas     = pts.Select(p => p.Value.area.ToString("F3")).ToArray(),
                };
            }
            // IC 검량선 계산
            if (stdPointsIC.Count >= 2)
            {
                var pts = stdPointsIC.OrderBy(p => p.Key).ToList();
                var (a, b, r2) = LinearRegressionZeroShift(pts.Select(p => p.Value.conc).ToList(),
                                                            pts.Select(p => p.Value.area).ToList());
                if (calData == null) calData = new TocCalibrationData { Method = fileMethod };
                calData.Slope_IC     = a.ToString("F4");
                calData.Intercept_IC = b.ToString("F4");
                calData.R2_IC        = r2.ToString("F5");
                calData.StdConcs_IC  = pts.Select(p => p.Value.conc.ToString("F4")).ToArray();
                calData.StdAreas_IC  = pts.Select(p => p.Value.area.ToString("F3")).ToArray();
            }
            cal = calData;
        }
        catch (Exception ex) { }
        return result;
    }

    // ── 시마즈 TOC-L PDF (분석결과 보고서) ───────────────────────────────────
    // 가감법(TCIC): 시료별 페이지 구조
    //   Sample Name: {name}
    //   Unknown TOC {dil} TOC:{toc}ppm TC:{tc}ppm IC:{ic}ppm
    //   Anal.: TC → Mean Area {area} / Mean Conc. {conc}ppm
    //   Anal.: IC → Mean Area {area} / Mean Conc. {conc}ppm
    // NPOC: 요약 표 (Unknown NPOC {name} {dil} NPOC:{conc}ppm)
    private static (List<TocInstrumentRow> rows, TocCalibrationData? cal) ParseShimadzuPdf(string path)
    {
        var rows = new List<TocInstrumentRow>();
        TocCalibrationData? cal = null;

        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "Logs", "ShimadzuTocPdfDebug.log");
        void Log(string msg) { if (App.EnableLogging) { try { File.AppendAllText(logPath, $"{DateTime.Now:HH:mm:ss} {msg}\n"); } catch { } } }

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            Log($"=== ParseShimadzuPdf: {Path.GetFileName(path)}, 페이지={doc.NumberOfPages} ===");

            // ── 페이지별 줄 단위 텍스트 추출 ──
            var pageLines = new List<List<string>>();
            foreach (var page in doc.GetPages())
            {
                var words = page.GetWords().ToList();
                var lineGroups = new List<List<UglyToad.PdfPig.Content.Word>>();
                foreach (var w in words)
                {
                    var grp = lineGroups.FirstOrDefault(g =>
                        Math.Abs(g[0].BoundingBox.Bottom - w.BoundingBox.Bottom) < 5);
                    if (grp != null) grp.Add(w);
                    else lineGroups.Add(new List<UglyToad.PdfPig.Content.Word> { w });
                }
                var lines = lineGroups
                    .OrderByDescending(g => g[0].BoundingBox.Bottom)
                    .Select(g => string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)).Trim())
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .ToList();
                pageLines.Add(lines);
            }

            var allLines = pageLines.SelectMany(p => p).ToList();

            // ── 방법 감지: 결과 요약 라인으로 판단 ──
            // TCIC: "Unknown TOC {dil} TOC:{x}ppm TC:{y}ppm IC:{z}ppm"
            // NPOC: "Unknown NPOC {dil} NPOC:{x}ppm"
            bool isTcic = allLines.Any(l =>
                Regex.IsMatch(l, @"Unknown\s+TOC\s+[0-9.]+\s+TOC:", RegexOptions.IgnoreCase));
            string fileMethod = isTcic ? "TCIC" : "NPOC";
            Log($"방법 감지: {fileMethod}");

            // ── 검량선 파싱 ──
            // 시마즈 PDF: "Standard TC" 또는 "Standard IC" 로 섹션 전환
            // "Slope: 5.553 Area 62.381" / "Intercept 0.000" / "r^2 0.9994 50"
            var slopeRe     = new Regex(@"^Slope:\s*([0-9.]+)", RegexOptions.IgnoreCase);
            var interceptRe = new Regex(@"^Intercept\s+([0-9.]+)", RegexOptions.IgnoreCase);
            var r2Re        = new Regex(@"^r\^2\s+([0-9.]+)", RegexOptions.IgnoreCase);
            // 표준점 Mean Area 추출:
            //   "Conc: {x}ppm"   → 현재 표준점 농도
            //   "Mean Area {v}"  → 해당 표준점의 기기 Mean Area (별도 줄)
            var concRe     = new Regex(@"^Conc:\s*([0-9.]+)ppm", RegexOptions.IgnoreCase);
            var meanAreaCalRe = new Regex(@"^Mean Area\s+([0-9.]+)", RegexOptions.IgnoreCase);

            string slope_TC = "", intercept_TC = "", r2_TC = "";
            string slope_IC = "", intercept_IC = "", r2_IC = "";
            bool inIcCal = false;

            // 표준점: (conc, meanArea) 목록
            var stdInstTC = new List<(double conc, double area)>();
            var stdInstIC = new List<(double conc, double area)>();
            double? curStdConc = null;

            // 검량선 섹션 스캔: 전체 파일 순회 (Sample Name: TC_ST/IC_ST 도 검량선 페이지에 있으므로 break 금지)
            bool inCalSection = true; // 검량선 페이지인 동안만 slope 수집
            foreach (var line in allLines)
            {
                if (line.Equals("Standard TC", StringComparison.OrdinalIgnoreCase)) { inIcCal = false; inCalSection = true; curStdConc = null; continue; }
                if (line.Equals("Standard IC", StringComparison.OrdinalIgnoreCase)) { inIcCal = true;  inCalSection = true; curStdConc = null; continue; }
                // NPOC 검량선도 "Standard NPOC" 또는 "Standard TC" 로 시작
                if (line.StartsWith("Standard NPOC", StringComparison.OrdinalIgnoreCase)) { inIcCal = false; inCalSection = true; curStdConc = null; continue; }
                // 실제 시료 섹션 시작 감지 (Unknown TOC/NPOC 결과 라인)
                if (Regex.IsMatch(line, @"^Unknown\s+(TOC|NPOC)\s+", RegexOptions.IgnoreCase))
                    { inCalSection = false; continue; }

                if (!inCalSection) continue;

                // 표준점 농도 감지: "Conc: 1.000ppm"
                var cm = concRe.Match(line);
                if (cm.Success && double.TryParse(cm.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var sc))
                { curStdConc = sc; continue; }

                // Mean Area 직접 감지: "Mean Area 1.298"
                var mam = meanAreaCalRe.Match(line);
                if (mam.Success && curStdConc.HasValue &&
                    double.TryParse(mam.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var ma))
                {
                    if (inIcCal) stdInstIC.Add((curStdConc.Value, ma));
                    else         stdInstTC.Add((curStdConc.Value, ma));
                    curStdConc = null; // 같은 표준점 중복 방지
                    continue;
                }

                var sm = slopeRe.Match(line);
                if (sm.Success) { if (inIcCal) slope_IC = sm.Groups[1].Value; else slope_TC = sm.Groups[1].Value; }
                var im = interceptRe.Match(line);
                if (im.Success) { if (inIcCal) intercept_IC = im.Groups[1].Value; else intercept_TC = im.Groups[1].Value; }
                var rm = r2Re.Match(line);
                if (rm.Success) { if (inIcCal) r2_IC = rm.Groups[1].Value; else r2_TC = rm.Groups[1].Value; }
            }

            if (!string.IsNullOrEmpty(slope_TC))
            {
                // 기기출력값: PDF 직독 slope/intercept/R² + Mean Area
                // 파싱값: Mean Area → ZeroShift 회귀 계산
                (double parsedSlope_TC, double parsedIntercept_TC, double parsedR2_TC) = (0, 0, 0);
                (double parsedSlope_IC, double parsedIntercept_IC, double parsedR2_IC) = (0, 0, 0);
                if (stdInstTC.Count >= 2)
                    (parsedSlope_TC, parsedIntercept_TC, parsedR2_TC) = LinearRegressionZeroShift(
                        stdInstTC.Select(p => p.conc).ToList(), stdInstTC.Select(p => p.area).ToList());
                if (stdInstIC.Count >= 2)
                    (parsedSlope_IC, parsedIntercept_IC, parsedR2_IC) = LinearRegressionZeroShift(
                        stdInstIC.Select(p => p.conc).ToList(), stdInstIC.Select(p => p.area).ToList());

                cal = new TocCalibrationData
                {
                    Method       = fileMethod,
                    // 파싱값 (기본 표시)
                    Slope_TC     = stdInstTC.Count >= 2 ? parsedSlope_TC.ToString("F4") : slope_TC,
                    Intercept_TC = stdInstTC.Count >= 2 ? parsedIntercept_TC.ToString("F4") : intercept_TC,
                    R2_TC        = stdInstTC.Count >= 2 ? parsedR2_TC.ToString("F5")        : r2_TC,
                    Slope_IC     = stdInstIC.Count >= 2 ? parsedSlope_IC.ToString("F4") : slope_IC,
                    Intercept_IC = stdInstIC.Count >= 2 ? parsedIntercept_IC.ToString("F4") : intercept_IC,
                    R2_IC        = stdInstIC.Count >= 2 ? parsedR2_IC.ToString("F5")        : r2_IC,
                    StdConcs     = stdInstTC.Select(p => p.conc.ToString("F4")).ToArray(),
                    StdAreas     = stdInstTC.Select(p => p.area.ToString("F3")).ToArray(),
                    StdConcs_IC  = stdInstIC.Select(p => p.conc.ToString("F4")).ToArray(),
                    StdAreas_IC  = stdInstIC.Select(p => p.area.ToString("F3")).ToArray(),
                    // 기기출력값 (PDF 직독)
                    Slope_TC_Inst     = slope_TC,
                    Intercept_TC_Inst = intercept_TC,
                    R2_TC_Inst        = r2_TC,
                    Slope_IC_Inst     = slope_IC,
                    Intercept_IC_Inst = intercept_IC,
                    R2_IC_Inst        = r2_IC,
                    StdAreas_Inst     = stdInstTC.Select(p => p.area.ToString("F3")).ToArray(),
                    StdAreas_IC_Inst  = stdInstIC.Select(p => p.area.ToString("F3")).ToArray(),
                };
                Log($"검량선: slope_TC={slope_TC}(기기)/slope_TC_parsed={parsedSlope_TC:F4}(파싱), stdTC={stdInstTC.Count}점");
            }

            // ── 시료 파싱 ──
            if (isTcic)
            {
                // 가감법: allLines 전체를 순차 스캔
                // 패턴:
                //   Sample Name: {name}
                //   ...
                //   Unknown TOC {dil} TOC:{toc}ppm TC:{tc}ppm IC:{ic}ppm   ← 결과 요약
                //   Anal.: TC → Mean Area {area} → Mean Conc. {conc}ppm
                //   Anal.: IC → Mean Area {area} → Mean Conc. {conc}ppm

                var summaryRe = new Regex(
                    @"Unknown\s+TOC\s+[0-9.]+\s+TOC:(-?[0-9.]+)ppm\s+TC:(-?[0-9.]+)ppm\s+IC:(-?[0-9.]+)ppm",
                    RegexOptions.IgnoreCase);
                var sampleNameRe = new Regex(@"^Sample Name:\s*(.+)$", RegexOptions.IgnoreCase);
                var meanAreaRe   = new Regex(@"^Mean Area\s+([0-9.]+)");
                var meanConcRe   = new Regex(@"^Mean Conc\.\s+([0-9.]+)ppm");

                string? curName = null;
                string tocConc = "", tcConc = "", tcArea = "", icConc = "", icArea = "";
                bool inAnalTC = false, inAnalIC = false;

                for (int i = 0; i < allLines.Count; i++)
                {
                    var line = allLines[i];

                    // 시료명 감지
                    var nm = sampleNameRe.Match(line);
                    if (nm.Success)
                    {
                        // 이전 시료 저장
                        if (curName != null && (!string.IsNullOrEmpty(tcConc) || !string.IsNullOrEmpty(icConc)))
                        {
                            Log($"시료: {curName} → toc={tocConc}, tc={tcConc}, tcA={tcArea}, ic={icConc}, icA={icArea}");
                            AddTcicRow(rows, curName, tcConc, tcArea, icConc, icArea, tocConc);
                        }
                        curName = nm.Groups[1].Value.Trim();
                        tocConc = tcConc = tcArea = icConc = icArea = "";
                        inAnalTC = inAnalIC = false;
                        continue;
                    }

                    if (curName == null) continue;

                    // 결과 요약 라인: TOC/TC/IC 농도 한 번에 추출
                    var sm2 = summaryRe.Match(line);
                    if (sm2.Success)
                    {
                        tocConc = sm2.Groups[1].Value;
                        tcConc  = sm2.Groups[2].Value;
                        icConc  = sm2.Groups[3].Value;
                        continue;
                    }

                    // Anal. 섹션 전환
                    if (line.Equals("Anal.: TC", StringComparison.OrdinalIgnoreCase)) { inAnalTC = true;  inAnalIC = false; continue; }
                    if (line.Equals("Anal.: IC", StringComparison.OrdinalIgnoreCase)) { inAnalTC = false; inAnalIC = true;  continue; }

                    // Mean Area / Mean Conc.
                    var am = meanAreaRe.Match(line);
                    if (am.Success) {
                        if (inAnalTC) tcArea = am.Groups[1].Value;
                        else if (inAnalIC) icArea = am.Groups[1].Value;
                        continue;
                    }
                    var cm = meanConcRe.Match(line);
                    if (cm.Success) {
                        if (inAnalTC && string.IsNullOrEmpty(tcConc)) tcConc = cm.Groups[1].Value;
                        else if (inAnalIC && string.IsNullOrEmpty(icConc)) icConc = cm.Groups[1].Value;
                        continue;
                    }
                }
                // 마지막 시료 저장
                if (curName != null && (!string.IsNullOrEmpty(tcConc) || !string.IsNullOrEmpty(icConc)))
                {
                    Log($"시료(마지막): {curName} → toc={tocConc}, tc={tcConc}, ic={icConc}");
                    AddTcicRow(rows, curName, tcConc, tcArea, icConc, icArea, tocConc);
                }
            }
            else
            {
                // NPOC: "Unknown NPOC {name} {dil} NPOC:{conc}ppm" 또는 표 형태
                // 가감법 결과 요약과 동일한 구조: Sample Name + Mean Conc 방식도 지원
                var npocSummaryRe = new Regex(
                    @"Unknown\s+NPOC\s+(.+?)\s+([0-9.]+)\s+NPOC:([0-9.]+)ppm",
                    RegexOptions.IgnoreCase);
                var npocTableRe = new Regex(
                    @"Unknown\s+NPOC\s+(.+?)\s+([0-9.]+)\s+([0-9.]+)\s*ppm",
                    RegexOptions.IgnoreCase);
                var sampleNameRe = new Regex(@"^Sample Name:\s*(.+)$", RegexOptions.IgnoreCase);
                var meanConcRe   = new Regex(@"^Mean Conc\.\s+([0-9.]+)ppm");
                var meanAreaRe   = new Regex(@"^Mean Area\s+([0-9.]+)");

                // 먼저 요약 라인으로 시도
                foreach (var line in allLines)
                {
                    var m = npocSummaryRe.Match(line);
                    if (!m.Success) m = npocTableRe.Match(line);
                    if (m.Success)
                    {
                        var name = m.Groups[1].Value.Trim();
                        var dil  = m.Groups[2].Value;
                        var conc = m.Groups[3].Value;
                        Log($"NPOC row: name='{name}', dil={dil}, conc={conc}");
                        rows.Add(new TocInstrumentRow {
                            RawName   = name, SN = NormalizeSN(name), Conc = conc,
                            Method    = "NPOC", IsControl = IsControlName(name),
                            Dilution  = dil != "1" ? dil : "",
                        });
                    }
                }

                // 표 형식 실패 시: Sample Name + Mean Conc 방식
                if (rows.Count == 0)
                {
                    Log("표 파싱 실패 → Sample Name + Mean Conc 방식");
                    string? curName = null;
                    string npocConc = "", npocArea = "";
                    foreach (var line in allLines)
                    {
                        var nm = sampleNameRe.Match(line);
                        if (nm.Success)
                        {
                            if (curName != null && !string.IsNullOrEmpty(npocConc))
                            {
                                Log($"NPOC (name-based): {curName} = {npocConc}");
                                rows.Add(new TocInstrumentRow {
                                    RawName = curName, SN = NormalizeSN(curName), Conc = npocConc,
                                    Area = npocArea, Method = "NPOC", IsControl = IsControlName(curName),
                                });
                            }
                            curName = nm.Groups[1].Value.Trim();
                            npocConc = npocArea = "";
                            continue;
                        }
                        if (curName == null) continue;
                        var am = meanAreaRe.Match(line);
                        if (am.Success) npocArea = am.Groups[1].Value;
                        var cm = meanConcRe.Match(line);
                        if (cm.Success) npocConc = cm.Groups[1].Value;
                    }
                    if (curName != null && !string.IsNullOrEmpty(npocConc))
                        rows.Add(new TocInstrumentRow {
                            RawName = curName, SN = NormalizeSN(curName), Conc = npocConc,
                            Area = npocArea, Method = "NPOC", IsControl = IsControlName(curName),
                        });
                }
            }

            Log($"파싱 완료: method={fileMethod}, rows={rows.Count}, cal={cal?.Slope_TC}");
        }
        catch (Exception ex)
        {
            Log($"파싱 오류: {ex}");
        }
        return (rows, cal);
    }

    static void AddTcicRow(List<TocInstrumentRow> rows,
        string name, string tcConc, string tcArea, string icConc, string icArea, string tocConc)
    {
        string finalConc = tocConc;
        if (string.IsNullOrEmpty(finalConc))
        {
            if (double.TryParse(tcConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var tv) &&
                double.TryParse(icConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var iv))
                finalConc = (tv - iv).ToString("F4", CultureInfo.InvariantCulture);
        }
        rows.Add(new TocInstrumentRow
        {
            RawName   = name,
            SN        = NormalizeSN(name),
            Conc      = finalConc,
            Area      = tcArea,
            Method    = "TCIC",
            IsControl = IsControlName(name),
            TCAU      = tcArea,
            TCcon     = tcConc,
            ICAU      = icArea,
            ICcon     = icConc,
        });
    }

    /// <summary>Result 컬럼 문자열에서 TOC 최종값 추출. "TOC:x.xppm TC:... IC:..." 형태 파싱.</summary>
    static string ExtractTocValueFromResult(string res, string tcConc, string icConc)
    {
        // "TOC:x.xppm" 패턴
        var m = Regex.Match(res, @"TOC\s*:\s*([\d.]+)", RegexOptions.IgnoreCase);
        if (m.Success) return m.Groups[1].Value;
        // 없으면 TC-IC 직접 계산
        if (double.TryParse(tcConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var tcv) &&
            double.TryParse(icConc, NumberStyles.Any, CultureInfo.InvariantCulture, out var icv))
            return (tcv - icv).ToString("F4");
        return tcConc;
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
            try
            {
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(fileStream, Encoding.GetEncoding(949)))
                {
                    allLines = reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries);
                }
            }
            catch
            {
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(fileStream, Encoding.UTF8))
                {
                    allLines = reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries);
                }
            }

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
                else if (!name.StartsWith("Cal_", StringComparison.OrdinalIgnoreCase))
                {
                    // 시료 데이터 행 (정도관리 시료도 보존 — 기록부 증거)
                    rows.Add(new TocInstrumentRow
                    {
                        RawName   = name,
                        SN        = NormalizeSN(name),
                        Area      = parseArea ? area.ToString("F3") : cols[3].Trim(),
                        Conc      = parseConc ? conc.ToString("F4") : cols[1].Trim(),
                        Method    = "NPOC", // 아래에서 확정
                        IsControl = IsControlName(name),
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
                cal.StdAreas     = tcAreas.Select(v => v.ToString("F3")).ToArray();
            }
            if (hasIC)
            {
                var (a, b, r2) = LinearRegression(icConcs, icAreas);
                cal.Slope_IC     = a.ToString("F4");
                cal.Intercept_IC = b.ToString("F4");
                cal.R2_IC        = r2.ToString("F5");
                cal.StdConcs_IC  = [.. icConcs.Select(c => c.ToString("F2"))];
                cal.StdAreas_IC  = [.. icAreas.Select(v => v.ToString("F3"))];
            }

            cal.Method = (hasTC && hasIC) ? "TCIC" : "NPOC";
            // 시료 행 메서드 업데이트
            foreach (var r in rows) r.Method = cal.Method;
        }
        catch (Exception ex) { }
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
                    StdAreas     = [.. stdTcAreas.Select(v => v.ToString("F3", CultureInfo.InvariantCulture))],
                    StdConcs_IC  = [.. stdIcConcs.Select(c => c.ToString("F2", CultureInfo.InvariantCulture))],
                    StdAreas_IC  = [.. stdIcAreas.Select(v => v.ToString("F3", CultureInfo.InvariantCulture))],
                };
            }

        }
        catch (Exception ex)
        {
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

    // Zero Shift 회귀: 절편 0 강제 (slope = Σ(xi·yi)/Σ(xi²))
    static (double a, double b, double r2) LinearRegressionZeroShift(List<double> xs, List<double> ys)
    {
        int n = xs.Count;
        double sumXY = 0, sumX2 = 0, sumY = 0;
        for (int i = 0; i < n; i++)
        { sumXY += xs[i]*ys[i]; sumX2 += xs[i]*xs[i]; sumY += ys[i]; }
        if (sumX2 < 1e-12) return (0, 0, 0);
        double a = sumXY / sumX2; // slope (intercept=0 강제)
        // R² = 1 - SS_res/SS_tot (평균 기준)
        double meanY = sumY / n, ssTot = 0, ssRes = 0;
        for (int i = 0; i < n; i++)
        { double diff = ys[i] - meanY; ssTot += diff*diff; double res = ys[i] - a*xs[i]; ssRes += res*res; }
        double r2 = ssTot < 1e-12 ? 1.0 : 1.0 - ssRes / ssTot;
        return (a, 0.0, r2);
    }

    // ── 유틸리티 ──────────────────────────────────────────────────────────

    // 헤더에서 컬럼 인덱스 찾기 (여러 후보 이름 지원)
    static int FindColumnIndex(List<string> headerFields, params string[] candidates)
    {
        for (int i = 0; i < headerFields.Count; i++)
        {
            var field = headerFields[i].Trim().Trim('"');
            foreach (var candidate in candidates)
            {
                if (field.Equals(candidate, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
        }
        return -1;
    }

    static bool IsControlName(string name)
    {
        var upper = name.Trim().ToUpperInvariant();
        return ControlKeywords.Any(k => upper.StartsWith(k) || upper.Equals(k));
    }

    static List<string> ReadShimadzuLines(string path)
    {
        // CP949(한국어 시마즈 파일) 우선 — Program.cs에서 RegisterProvider 등록 후 사용 가능
        try
        {
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fileStream, Encoding.GetEncoding(949)))
            {
                return reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries).ToList();
            }
        }
        catch { }
        try
        {
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fileStream, Encoding.UTF8))
            {
                return reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries).ToList();
            }
        }
        catch
        {
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fileStream, Encoding.Latin1))
            {
                return reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries).ToList();
            }
        }
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

        // 파일 기반 로그
        string logPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Logs", "JenaPdfDebug.log");
        void Log(string msg)
        {
            if (App.EnableLogging)
            {
                try { File.AppendAllText(logPath, $"{DateTime.Now:HH:mm:ss} {msg}\n"); } catch { }
            }
        }

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            Log($"=== ParseJenaPdf 시작: {System.IO.Path.GetFileName(path)}, 페이지={doc.NumberOfPages} ===");

            // 페이지별로 줄 단위 추출 (페이지 경계가 검량선 섹션 분리에 필요)
            var pageLines = new List<List<string>>();
            int pageIdx = 0;
            foreach (var page in doc.GetPages())
            {
                pageIdx++;
                var words = page.GetWords().ToList();
                Log($"페이지 {pageIdx}: 단어 {words.Count}개");

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
            Log($"총 줄 수: {allLines.Count}");
            for (int li = 0; li < Math.Min(allLines.Count, 60); li++)
                Log($"  [{li:D3}] {allLines[li]}");

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
                    cal.StdAreas_IC = ic.areas.Select(v => v.ToString("F3",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.Slope_TC = tc.k1;
                    cal.Intercept_TC = tc.k0;
                    cal.StdConcs = tc.concs.Select(c => c.ToString("F2",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                    cal.StdAreas = tc.areas.Select(v => v.ToString("F3",
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
                    cal.StdAreas = npoc.areas.Select(v => v.ToString("F3",
                        System.Globalization.CultureInfo.InvariantCulture)).ToArray();
                }
            }

            // ───── 시료 파싱 ─────
            // 시료 ID 행: "{name} YYYY-MM-DD [오전|오후|AM|PM] HH:MM:SS {rack#}"
            var sampleIdRe = new Regex(
                @"^(?<name>.+?)\s+\d{4}-\d{2}-\d{2}\s+(?:오전|오후|AM|PM|am|pm)\s+\d{1,2}:\d{2}:\d{2}\s+\d+\s*$");

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

            Log($"시료 파싱 시작 (sampleIdRe 탐색)");
            for (int i = 0; i < allLines.Count; i++)
            {
                var m = sampleIdRe.Match(allLines[i]);
                if (!m.Success) continue;

                var name = NormalizeJenaName(m.Groups["name"].Value);
                Log($"  시료ID 매칭: [{i}] name='{name}', line='{allLines[i]}'");

                // 정도관리 시료(BK/CCV/FBK/MBK/DW)도 기록부 증거이므로 보존 — IsControl 플래그만 세팅
                bool isCtrl = IsControlName(name);

                // 다음 10줄 이내에서 값 라인 탐색 (다음 시료 ID 만나면 중단)
                double? tocVal  = null;
                double? tcConc  = null;
                double? tcArea  = null;
                double? icConc  = null;
                double? icArea  = null;
                bool isTcic = false;

                for (int j = i + 1; j < Math.Min(i + 10, allLines.Count); j++)
                {
                    if (sampleIdRe.IsMatch(allLines[j])) break;

                    // 가감법 TOC 라인 (계산값: TOC = TC - IC)
                    var tm = tocLineRe.Match(allLines[j]);
                    if (tm.Success)
                    {
                        isTcic = true;
                        tocVal = ParseJenaConcValue(tm.Groups[1].Value, tm.Groups[2].Value);
                        continue;
                    }

                    // 가감법 IC/TC 라인 — 농도 + AU 면적 모두 수집
                    var icm = icTcLineRe.Match(allLines[j]);
                    if (icm.Success)
                    {
                        isTcic = true;
                        var concVal = ParseJenaConcValue(icm.Groups[2].Value, icm.Groups[3].Value);
                        var am = auRe.Match(allLines[j]);
                        var areaVal = am.Success ? ParseJenaAreaValue(am.Groups[1].Value + "AU") : null;

                        if (icm.Groups[1].Value.Equals("TC", StringComparison.OrdinalIgnoreCase))
                        {
                            tcConc = concVal;
                            tcArea = areaVal;
                            Log($"    TC: conc={tcConc}, area={tcArea}");
                            break; // TC가 시료당 마지막 값 라인
                        }
                        else // IC
                        {
                            icConc = concVal;
                            icArea = areaVal;
                            Log($"    IC: conc={icConc}, area={icArea}");
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

                // TCIC: tocVal이 없으면 TC-IC로 계산
                if (isTcic && tocVal == null && tcConc.HasValue && icConc.HasValue)
                    tocVal = tcConc.Value - icConc.Value;

                if (tocVal == null) { Log($"    tocVal null — 스킵: {name}"); continue; }

                string Fmt(double? v) => v?.ToString("0.####", CultureInfo.InvariantCulture) ?? "";

                rows.Add(new TocInstrumentRow
                {
                    RawName = name,
                    SN      = name,
                    Conc    = tocVal.Value.ToString("0.###", CultureInfo.InvariantCulture),
                    Area    = Fmt(tcArea),
                    Method  = isTcic ? "TCIC" : "NPOC",
                    IsControl = isCtrl,
                    // TCIC 전용
                    TCAU  = Fmt(tcArea),
                    TCcon = Fmt(tcConc),
                    ICAU  = Fmt(icArea),
                    ICcon = Fmt(icConc),
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

            Log($"파싱 완료: 검량선={cal?.Method}, 시료={rows.Count}건");
        }
        catch (Exception ex)
        {
            Log($"파싱 오류: {ex}");
        }
        return (rows, cal);
    }

    // ── 스칼라 NPOC 분석결과 PDF ───────────────────────────────────────────────
    // 구조: 1페이지=검량선(a/b/R²), 2페이지~=시료별 SI.No. 블록
    // 각 블록: SI.No. → {번호} {위치} {U|WS} {희석배수} {샘플ID} → S.No. → {면적} {농도} → 평균행

    static List<TocInstrumentRow> ParseSkalarNpocPdf(string path, out TocCalibrationData? cal)
    {
        cal = null;
        var rows = new List<TocInstrumentRow>();
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var pages = doc.GetPages().ToList();

            // 1페이지: 검량선
            if (pages.Count > 0)
            {
                var (_, calData) = ParseSkalarCalibrationPdf_Text(pages[0].Text);
                cal = calData;
            }

            // 2페이지~: 시료 데이터
            // 라인 추출 함수 (JenaPdf와 동일 방식)
            static List<string> ExtractLines(UglyToad.PdfPig.Content.Page page)
            {
                var words = page.GetWords().ToList();
                var groups = new List<List<UglyToad.PdfPig.Content.Word>>();
                foreach (var w in words)
                {
                    var g = groups.FirstOrDefault(gr => Math.Abs(gr[0].BoundingBox.Bottom - w.BoundingBox.Bottom) < 4);
                    if (g != null) g.Add(w);
                    else groups.Add(new List<UglyToad.PdfPig.Content.Word> { w });
                }
                return groups.OrderByDescending(g => g[0].BoundingBox.Bottom)
                             .Select(g => string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)).Trim())
                             .Where(l => !string.IsNullOrWhiteSpace(l))
                             .ToList();
            }

            // SI. No. ID 블록 파싱
            // 패턴: "SI. No. ID" 헤더 → "{번호} {위치} {타입} {희석배수} {샘플ID}" → skip → "S.No." → skip(주입행) → "{면적} {농도} CV {cv}"
            var siHeaderRe = new Regex(@"^SI\.?\s*No\.?");
            // "{번호} {위치} {U|WS} {희석배수} {나머지=샘플ID}"
            var sampleLineRe = new Regex(@"^(\d+)\s+\w+\s+(U|WS)\s+([\d.]+)\s+(.+)$");
            // 평균행: "{면적} {농도} CV {cv}" (앞에 숫자 없음)
            var avgLineRe = new Regex(@"^([\d.]+)\s+([\d.]+)\s+CV\s+([\d.]+)$");

            string? currentSampleId = null;
            string? currentType = null;
            string? currentDilution = null;
            int parseState = 0; // 0=찾는중, 1=샘플헤더읽음, 2=S.No.기다림, 3=주입행기다림, 4=평균행기다림

            foreach (var page in pages.Skip(1))
            {
                var lines = ExtractLines(page);
                foreach (var line in lines)
                {
                    switch (parseState)
                    {
                        case 0: // SI. No. 헤더 탐색
                            if (siHeaderRe.IsMatch(line)) parseState = 1;
                            break;

                        case 1: // 샘플 헤더 라인 파싱
                            var sm = sampleLineRe.Match(line);
                            if (sm.Success)
                            {
                                currentType     = sm.Groups[2].Value; // U or WS
                                currentDilution = sm.Groups[3].Value;
                                currentSampleId = sm.Groups[4].Value.Trim();
                                parseState = 2;
                            }
                            else parseState = 0;
                            break;

                        case 2: // "S.No." 또는 "NPOC" 라인 건너뜀
                            if (line.Trim() == "S.No.") parseState = 3;
                            // "NPOC 100 : mg/l" 등은 그냥 무시
                            break;

                        case 3: // 주입 데이터 라인 건너뜀 (숫자로 시작)
                            parseState = 4;
                            break;

                        case 4: // 평균행
                            var am = avgLineRe.Match(line);
                            if (am.Success && currentSampleId != null)
                            {
                                string area  = am.Groups[1].Value;
                                string conc  = am.Groups[2].Value;
                                bool isCtrl  = currentType == "WS"
                                    || ControlKeywords.Any(k => currentSampleId.StartsWith(k, StringComparison.OrdinalIgnoreCase));
                                var sn = NormalizeSN(currentSampleId);
                                rows.Add(new TocInstrumentRow
                                {
                                    RawName   = currentSampleId,
                                    SN        = sn,
                                    Area      = area,
                                    Conc      = conc,
                                    Method    = "NPOC",
                                    IsControl = isCtrl,
                                    Dilution  = currentDilution ?? "1",
                                });
                            }
                            currentSampleId = null;
                            parseState = 0;
                            break;
                    }
                }
            }

        }
        catch (Exception ex)
        {
        }
        return rows;
    }

    // 내부 헬퍼: 텍스트에서 검량선 값 추출 (SkalarCalibrationPdf와 공유)
    static (bool isIC, TocCalibrationData? cal) ParseSkalarCalibrationPdf_Text(string text)
    {
        var methodMatch = Regex.Match(text, @"메소드\s*이름\s*:\s*(\S+)");
        bool isIC = methodMatch.Success &&
                    methodMatch.Groups[1].Value.IndexOf("IC", StringComparison.OrdinalIgnoreCase) >= 0;
        var abMatch = Regex.Match(text, @"a\s*=\s*([\d.]+)\s+b\s*=\s*([\d.]+)\s+r\s*=\s*([\d.]+)");
        var r2Match = Regex.Match(text, @"R-Squared\s*=\s*([\d.]+)");
        if (!abMatch.Success) return (isIC, null);
        var cal = new TocCalibrationData { Method = isIC ? "TCIC" : "NPOC" };
        string slope = abMatch.Groups[1].Value, intercept = abMatch.Groups[2].Value, r2 = r2Match.Success ? r2Match.Groups[1].Value : "";
        if (isIC) { cal.Slope_IC = slope; cal.Intercept_IC = intercept; cal.R2_IC = r2; }
        else       { cal.Slope_TC = slope; cal.Intercept_TC = intercept; cal.R2_TC = r2; }
        return (isIC, cal);
    }

    // ── 스칼라 검량선 PDF (TC 또는 IC 단일 채널) ───────────────────────────────

    /// <summary>
    /// 스칼라 TOC 기기의 검량선 PDF 파싱.
    /// TC: 메소드 이름 TOC_Analysis / IC: 메소드 이름 TOC_Analysis_IC
    /// 추출값: a(기울기), b(절편), r, R-Squared
    /// </summary>
    public static (List<TocInstrumentRow> rows, TocCalibrationData? cal) ParseSkalarCalibrationPdf(string path)
    {
        var rows = new List<TocInstrumentRow>();
        TocCalibrationData? cal = null;
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var allText = string.Join("\n", doc.GetPages().Select(p => p.Text));
            var (_, calData) = ParseSkalarCalibrationPdf_Text(allText);
            cal = calData;
        }
        catch (Exception ex)
        {
        }
        return (rows, cal);
    }
}
