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
        // 배치 날짜는 AnalysisTime에서 추출되어 있으므로 별도 처리 없음
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

        // 검정곡선 표준점 수집 (Cal 행)
        if (isCal)
        {
            var lm = CalibLevelRx.Match(rawName);
            int level = lm.Success ? int.Parse(lm.Groups[1].Value) : compound.Calibration.Count + 1;

            // Final Conc = 표준 공칭농도
            compound.Calibration.Add(new GcCalibrationPoint
            {
                Level       = level,
                Enable      = "x",
                Conc        = final,
                Response    = resp,
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
