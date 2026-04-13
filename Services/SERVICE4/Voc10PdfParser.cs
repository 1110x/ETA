using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

/// <summary>
/// Agilent MassHunter Quantitative Analysis Results Summary Report PDF 파서
/// VOC10, 벤젠 등 다성분 ISTD GC 분석결과
///
/// PDF 구조:
///   페이지 헤더: "Quantitative Analysis Results Summary Report"
///   배치 정보:  Batch Data Path / Analysis Time (건너뜀)
///   반복 섹션:
///     [성분명]                           ← 단독줄
///     Data File  Type  RT  Resp.  ISTD Resp  Final Conc.   ← 컬럼 헤더
///     [데이터 행들]                       ← 파일명 + 타입 + 값
///   멀티라인 파일명: "FBK(PL 25-12-  Blank 1902 ND" / "30).D"
///   종료 조건: "Quant Calibration Report" 섹션 도달
///
///   Type: Cal(검정곡선), Sample(시료), Blank(바탕), QC(정도관리)
///   불검출: Final Conc. == "ND"
///   GcInstrumentFile(VocMulti) 형식으로 반환 — 기존 처리 코드 재활용
/// </summary>
public static class Voc10PdfParser
{
    private static readonly HashSet<string> TypeKeywords =
        new(StringComparer.OrdinalIgnoreCase) { "Cal", "Sample", "Blank", "QC", "Check" };

    private static readonly Regex CalibLevelRx = new(
        @"ST-(\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static bool IsQcRow(string rawName, string type) =>
        type.Equals("Blank", StringComparison.OrdinalIgnoreCase) ||
        type.Equals("QC",    StringComparison.OrdinalIgnoreCase) ||
        rawName.StartsWith("MBK",  StringComparison.OrdinalIgnoreCase) ||
        rawName.StartsWith("FBK",  StringComparison.OrdinalIgnoreCase) ||
        rawName.StartsWith("CCV",  StringComparison.OrdinalIgnoreCase) ||
        rawName.StartsWith("BK",   StringComparison.OrdinalIgnoreCase) ||
        rawName.StartsWith("DW",   StringComparison.OrdinalIgnoreCase);

    // 건너뛸 페이지/배치 헤더 패턴
    private static readonly Regex SkipLineRx = new(
        @"^(Quantitative Analysis Results|Batch Data Path|Analysis Time|Calibration Last|Page \d+ of \d+|Analyst Name|Report Generator|Analyze Quant|Report Quant)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Quant Calibration Report 섹션 — 공식/R² 파싱용
    // "y = 366.446256 * x + 1.841816"  또는  "{name} = 366.446256 * x - 0.123"
    private static readonly Regex FormulaRx = new(
        @"(-?[\d]+(?:[.,][\d]+)?(?:[Ee][+\-]?\d+)?)\s*\*\s*x\s*([+\-]\s*[\d]+(?:[.,][\d]+)?(?:[Ee][+\-]?\d+)?)?",
        RegexOptions.Compiled);
    // "R^2 = 0.997"  또는  "R²= 0.997"
    private static readonly Regex R2Rx = new(
        @"R\^?2\s*=\s*([\d.]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>VOC10 Agilent PDF 여부 확인</summary>
    public static bool IsVoc10Pdf(string path)
    {
        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var page = doc.GetPage(1);
            string text = string.Join(" ", page.GetWords().Select(w => w.Text));
            return text.Contains("Quantitative Analysis Results", StringComparison.OrdinalIgnoreCase) &&
                   text.Contains("Final Conc", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    public static GcInstrumentFile Parse(string path)
    {
        var file = new GcInstrumentFile
        {
            Format     = GcFileFormat.VocMulti,
            SourcePath = path,
        };

        string? currentCompoundName = null;
        GcCompound? currentCompound = null;

        // 멀티라인 파일명 처리
        string? pendingPartialName = null;  // 아직 .D 없는 파일명 앞부분
        string? pendingType        = null;
        string[]? pendingData      = null;

        // 다음에 오는 비헤더 줄이 성분명 후보
        string? compoundCandidate = null;

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);

            foreach (var page in doc.GetPages())
            {
                foreach (var line in ExtractLines(page))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // ── 종료 조건: 검량선 보고서 섹션 ──────────────────
                    if (line.Contains("Quant Calibration Report", StringComparison.OrdinalIgnoreCase))
                        goto Done;

                    // ── 페이지/배치 헤더 건너뜀 ─────────────────────────
                    if (SkipLineRx.IsMatch(line)) continue;

                    // ── 컬럼 헤더줄 ─────────────────────────────────────
                    if (line.StartsWith("Data File", StringComparison.OrdinalIgnoreCase) &&
                        line.Contains("Type", StringComparison.OrdinalIgnoreCase) &&
                        line.Contains("Resp", StringComparison.OrdinalIgnoreCase))
                    {
                        // 바로 직전 후보가 성분명 → 새 성분 시작 (또는 페이지 경계 재개)
                        if (compoundCandidate != null)
                        {
                            currentCompoundName = compoundCandidate;
                            // 이미 존재하면 재사용 (페이지 경계 분할 대응)
                            var existing = file.Compounds.FirstOrDefault(c =>
                                c.Name.Equals(currentCompoundName, StringComparison.OrdinalIgnoreCase));
                            if (existing != null)
                                currentCompound = existing;
                            else
                            {
                                currentCompound = new GcCompound { Name = currentCompoundName, HasIstd = true };
                                file.Compounds.Add(currentCompound);
                            }
                            compoundCandidate = null;
                        }
                        pendingPartialName = null;
                        pendingType        = null;
                        pendingData        = null;
                        continue;
                    }

                    // ── 멀티라인 파일명 완성줄 (".D" 로 끝나는 단독 토큰) ──
                    if (pendingPartialName != null &&
                        line.Trim().EndsWith(".D", StringComparison.OrdinalIgnoreCase) &&
                        !line.Contains("Cal") && !line.Contains("Sample") &&
                        !line.Contains("Blank") && !line.Contains("QC"))
                    {
                        string fullName = pendingPartialName.TrimEnd('-') + line.Trim();
                        EmitRow(currentCompound, fullName, pendingType!, pendingData!);
                        pendingPartialName = null;
                        pendingType        = null;
                        pendingData        = null;
                        continue;
                    }

                    // ── 데이터 행 파싱 ──────────────────────────────────
                    var tokens = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                    int typeIdx = FindTypeIndex(tokens);

                    if (typeIdx >= 0 && currentCompound != null)
                    {
                        string rawName    = string.Join(" ", tokens[..typeIdx]);
                        string type       = tokens[typeIdx];
                        string[] dataTokens = tokens[(typeIdx + 1)..];

                        if (!rawName.EndsWith(".D", StringComparison.OrdinalIgnoreCase))
                        {
                            // 파일명이 불완전 → 다음 줄에 ".D" 이어짐
                            pendingPartialName = rawName;
                            pendingType        = type;
                            pendingData        = dataTokens;
                        }
                        else
                        {
                            EmitRow(currentCompound, rawName, type, dataTokens);
                            pendingPartialName = null;
                        }
                        compoundCandidate = null;
                        continue;
                    }

                    // ── 성분명 후보 ─────────────────────────────────────
                    // .D 없고, 타입 키워드 없고, 짧은 줄
                    if (typeIdx < 0 &&
                        !line.EndsWith(".D", StringComparison.OrdinalIgnoreCase) &&
                        line.Length < 80 &&
                        !line.StartsWith("Page ", StringComparison.OrdinalIgnoreCase))
                    {
                        compoundCandidate = line.Trim();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"VOC10 PDF 파싱 오류: {ex.Message}", ex);
        }

        Done:
        ParseCalibrationFormulas(path, file);
        return file;
    }

    // ── 행 추가 ───────────────────────────────────────────────────────────
    private static void EmitRow(GcCompound? compound, string rawName, string type, string[] data)
    {
        if (compound == null) return;

        // 값 파싱: 4=RT+Resp+ISTD+Final / 3=Resp+ISTD+Final / 2=ISTD+Final / 1=Final
        string rt = "", resp = "", istd = "", final = "ND";
        switch (data.Length)
        {
            case >= 4: rt = data[0]; resp = data[1]; istd = data[2]; final = data[3]; break;
            case 3:                  resp = data[0]; istd = data[1]; final = data[2]; break;
            case 2:                                  istd = data[0]; final = data[1]; break;
            case 1:                                                  final = data[0]; break;
        }

        bool isCal = type.Equals("Cal", StringComparison.OrdinalIgnoreCase);

        var row = new GcInstrumentRow
        {
            RawName   = rawName,
            Type      = type,
            Rt        = rt,
            Resp      = resp,
            IstdResp  = istd,
            FinalConc = final.Equals("ND", StringComparison.OrdinalIgnoreCase) ? "ND" : final,
            IsControl = IsQcRow(rawName, type),
        };

        compound.Rows.Add(row);

        // Cal 행: Calibration 포인트 생성 (Conc는 빈값 — nominal 농도는 ParseCalibrationFormulas에서 설정)
        if (isCal)
        {
            var lm = CalibLevelRx.Match(rawName);
            int level = lm.Success ? int.Parse(lm.Groups[1].Value) : compound.Calibration.Count + 1;
            compound.Calibration.Add(new GcCalibrationPoint
            {
                Level        = level,
                Enable       = "x",
                Conc         = "",   // nominal 농도는 Quant Calibration Report에서 채워짐
                Response     = resp,
                IstdResponse = istd,
            });
        }
    }

    // ── 타입 키워드 인덱스 탐색 ────────────────────────────────────────────
    private static int FindTypeIndex(string[] tokens)
    {
        for (int i = 0; i < tokens.Length; i++)
            if (TypeKeywords.Contains(tokens[i]))
                return i;
        return -1;
    }

    // ── Quant Calibration Report 파싱 ────────────────────────────────────
    // Agilent이 계산한 기울기/절편/R² 값을 PDF 텍스트에서 직접 추출하여
    // ComputeLinearRegression(back-calculated Conc + raw Resp 사용)의 오차를 방지
    private static void ParseCalibrationFormulas(string path, GcInstrumentFile file)
    {
        bool inCal = false;
        GcCompound? cur = null;

        try
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            foreach (var page in doc.GetPages())
            {
                foreach (var line in ExtractLines(page))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (!inCal)
                    {
                        if (line.Contains("Quant Calibration Report", StringComparison.OrdinalIgnoreCase))
                            inCal = true;
                        continue;
                    }

                    var t = line.Trim();

                    // 성분명 줄 인식 (기존에 파싱된 이름과 일치)
                    var matched = file.Compounds.FirstOrDefault(c =>
                        t.Equals(c.Name, StringComparison.OrdinalIgnoreCase) ||
                        t.StartsWith(c.Name + " ", StringComparison.OrdinalIgnoreCase));
                    if (matched != null) { cur = matched; continue; }

                    // 공식 줄: slope * x ± intercept
                    var fm = FormulaRx.Match(line);
                    if (fm.Success && cur != null)
                    {
                        var slopeStr = fm.Groups[1].Value.Replace(',', '.');
                        if (double.TryParse(slopeStr, NumberStyles.Float,
                                CultureInfo.InvariantCulture, out var slope))
                            cur.SlopeA = slope;

                        if (fm.Groups[2].Success)
                        {
                            var intStr = fm.Groups[2].Value.Replace(" ", "").Replace(',', '.');
                            if (double.TryParse(intStr, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out var intercept))
                                cur.Intercept = intercept;
                        }
                        continue;
                    }

                    // R² 줄
                    var r2m = R2Rx.Match(line);
                    if (r2m.Success && cur != null)
                    {
                        if (double.TryParse(r2m.Groups[1].Value, NumberStyles.Float,
                                CultureInfo.InvariantCulture, out var r2))
                            cur.R = Math.Sqrt(Math.Max(0, r2));
                        continue;
                    }

                    // 검량선 테이블 행: Calibration {level} x {nominal_conc} {response}
                    // (파일경로와 Calibration 셀이 별도 줄로 추출될 수 있어 .D 조건 제거)
                    if (cur != null)
                    {
                        var toks = t.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                        int ci = -1;
                        for (int i = 0; i < toks.Length; i++)
                            if (toks[i].Equals("Calibration", StringComparison.OrdinalIgnoreCase))
                            { ci = i; break; }

                        if (ci >= 0 && ci + 3 < toks.Length &&
                            int.TryParse(toks[ci + 1], out var lvl))
                        {
                            var concStr = toks[ci + 3].Replace(',', '.');
                            if (double.TryParse(concStr, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out var nomConc))
                            {
                                var cp = cur.Calibration.FirstOrDefault(p => p.Level == lvl);
                                if (cp != null)
                                    cp.Conc = nomConc.ToString(CultureInfo.InvariantCulture);
                            }
                        }
                        continue;
                    }
                }
            }
        }
        catch { }

        // 칼리브레이션 리포트에서 값을 못 찾은 성분은 기존 회귀로 폴백
        foreach (var c in file.Compounds.Where(c => c.SlopeA == null))
            GcInstrumentParser.ComputeLinearRegression(c);
    }

    // ── PdfPig 줄 추출 (Y ±5pt 그룹핑) ────────────────────────────────────
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
