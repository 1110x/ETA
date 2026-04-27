using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SkiaSharp;
using ETA.Services.Common;
using ETA.Views.Pages.PAGE1;

namespace ETA.Services.SERVICE1;

/// <summary>
/// Show2 (TestRecordBookParsedView) 의 섹션 구조를 정식 Excel 시험기록부로 출력.
///
/// 레이아웃:
///   [Title]   "○○○ 시험기록부"  (병합 1행)
///   [Meta]    문서번호 / 분석방법 / 결과표시 / 관련근거 (4행)
///   [Cal]     검정곡선 (UvVis 2~3행 / TCIC 4행 / VOC 3행 / BOD 식종표)
///   [Sample]  시료분석결과 (스키마별 컬럼)
///   [QC]      검정곡선의 보증
///   [Sig]     결재라인 (항상 마지막 페이지 하단)
///   [BG]      페이지 헤더 워터마크 (회사 로고)
///   [Footer]  페이지 번호
/// </summary>
public static class TestRecordBookExcelExporter
{
    private const int RowsPerPage = 32;          // A4 가로 + 18~20pt 행높이 기준
    private const string FontFamily = "맑은 고딕";
    private const string LegalRef =
        "1. 공공폐수처리시설 설치 및 운영관리지침(8차개정) 「공공폐수처리시설의 수질검사」 항";

    public static string Export(TestRecordBookParsedView.Model m)
    {
        var dir = Path.Combine(AppPaths.WritableDataRoot, "PrintCache");
        Directory.CreateDirectory(dir);
        var stamp  = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var safeNm = string.Concat((m.TargetTable ?? "시험기록부")
            .Split(Path.GetInvalidFileNameChars()));
        var path = Path.Combine(dir, $"{safeNm}_{stamp}.xlsx");

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("시험기록부");

        // 페이지 설정 — A4 세로 (사용자 레퍼런스 기준), 너비 자동 맞춤
        ws.PageSetup.PaperSize        = XLPaperSize.A4Paper;
        ws.PageSetup.PageOrientation  = XLPageOrientation.Portrait;
        ws.PageSetup.Margins.Left     = 0.4;
        ws.PageSetup.Margins.Right    = 0.4;
        ws.PageSetup.Margins.Top      = 1.10;   // 머리글 로고(60pt) + 여유
        ws.PageSetup.Margins.Bottom   = 0.55;
        ws.PageSetup.Margins.Header   = 0.30;
        ws.PageSetup.Margins.Footer   = 0.25;
        ws.PageSetup.CenterHorizontally = true;
        ws.PageSetup.FitToPages(1, 0);    // 가로 1페이지에 맞춤, 세로는 자유
        ws.PageSetup.ShowGridlines    = false;

        // 푸터 — 페이지 번호
        ws.PageSetup.Footer.Right.AddText("- &P / &N -");

        ws.Style.Font.FontName = FontFamily;
        ws.Style.Font.FontSize = 10;
        ws.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        // 스키마 결정
        var meta = ResolveSampleMeta(m);
        int nCols = Math.Max(meta.headers.Length, EstimateCalCols(m));

        int row = 1;

        // 0) 타이틀
        row = WriteTitle(ws, row, nCols, m);
        row += 1;

        // 1) 메타 4행
        row = WriteMeta(ws, row, nCols, m);
        row += 1;

        // 2) 검정곡선
        int calStart = row;
        row = WriteCalibration(ws, row, nCols, m);
        row += 1;

        int sampleStartRow = row;

        // 3) 시료 표
        if (m.SampleRows.Count > 0)
        {
            row = WriteSampleSection(ws, row, nCols, m, meta, "시료분석결과", m.SampleRows, m.SampleClassByRow);
            row += 1;
        }

        // 4) 검정곡선의 보증 (QC)
        if (m.QcRows.Count > 0)
        {
            row = WriteSampleSection(ws, row, nCols, m, meta, "검정곡선의 보증", m.QcRows, m.QcClassByRow);
            row += 1;
        }

        int contentLastRow = row - 1;

        // 5) 분석 비고/메모 (memo 길이 4자 이상일 때만 출력)
        row = WriteRemarksBox(ws, row, nCols, m);

        // 6) 결재라인 — 마지막 페이지 하단으로 패딩 푸쉬
        const int rowsPerPage = 40;       // Portrait A4 기준 추정
        const int sigHeight   = 4;
        int contentEnd = row - 1;
        int curPage = Math.Max(1, (contentEnd - 1) / rowsPerPage + 1);
        int targetSigStart = curPage * rowsPerPage - sigHeight + 1;
        if (targetSigStart < contentEnd + 3)
        {
            curPage++;
            targetSigStart = curPage * rowsPerPage - sigHeight + 1;
        }
        // 빈 행으로 푸쉬
        for (int r = row; r < targetSigStart; r++) ws.Row(r).Height = 16;
        int sigStartRow = WriteSignature(ws, targetSigStart, nCols);
        int sheetLastRow = sigStartRow + sigHeight - 1;

        // 컬럼 폭
        ApplyColumnWidths(ws, nCols, meta);

        // 인쇄영역
        ws.PageSetup.PrintAreas.Add(1, 1, sheetLastRow, nCols);

        wb.SaveAs(path);

        // 페이지 머리글 워터마크 — 모든 페이지 가운데 (OpenXML 후처리)
        TryInjectHeaderLogo(path);

        return path;
    }

    // ── 스키마 + 시료표 컬럼 구성 ──────────────────────────────────────────────
    private record SampleMeta(string Schema, string[] headers, int[] widths);

    private static SampleMeta ResolveSampleMeta(TestRecordBookParsedView.Model m)
    {
        bool hasTcic        = m.IsTcic;
        bool hasIstd        = m.SampleHeaders.Contains("ISTD");
        bool hasGravimetric = m.SampleHeaders.Contains("전무게") && m.SampleHeaders.Contains("후무게");
        bool hasColiform    = m.SampleHeaders.Contains("A") && m.SampleHeaders.Contains("B");
        bool isBod          = m.IsBod;

        string[] headers; int[] widths; string schema;
        if (isBod)
        {
            schema  = "BOD";
            headers = new[] { "시료명", "시료량(mL)", "D1(15min DO)", "D2(5Day DO)", "f(x/y)", "희석배수", "Result(mg/L)", "비고", "시료번호" };
            widths  = new[] { 24, 12, 14, 14, 10, 10, 14, 12, 12 };
        }
        else if (hasColiform)
        {
            schema  = "Coliform";
            headers = new[] { "시료명", "시료량(mL)", "A", "B", "희석배수", "Result(개/mL)", "비고", "시료번호" };
            widths  = new[] { 26, 12, 10, 10, 10, 16, 12, 12 };
        }
        else if (hasGravimetric)
        {
            schema  = "Gravimetric";
            headers = new[] { "시료명", "시료량(mL)", "전무게(g)", "후무게(g)", "무게차(g)", "희석배수", "Result(mg/L)", "비고", "시료번호" };
            widths  = new[] { 24, 12, 12, 12, 12, 10, 14, 12, 12 };
        }
        else if (hasTcic)
        {
            schema  = "TCIC";
            headers = new[] { "시료명", "TCAU", "TC con(mg/L)", "ICAU", "IC con(mg/L)", "희석배수", "Result(mg/L)", "비고", "시료번호" };
            widths  = new[] { 24, 11, 13, 11, 13, 10, 14, 12, 12 };
        }
        else if (hasIstd)
        {
            schema  = "VOC";
            headers = new[] { "시료명", "Area", "ISTD Resp.", "희석배수", "Result(mg/L)", "비고", "시료번호" };
            widths  = new[] { 26, 14, 14, 10, 16, 12, 12 };
        }
        else
        {
            schema  = "UvVis";
            headers = new[] { "시료명", "시료량(mL)", "흡광도", "희석배수", "Result(mg/L)", "비고", "시료번호" };
            widths  = new[] { 28, 12, 14, 10, 16, 14, 12 };
        }
        return new SampleMeta(schema, headers, widths);
    }

    private static int EstimateCalCols(TestRecordBookParsedView.Model m)
    {
        int st = Math.Max(m.StandardKeys.Count, Math.Max(m.TcStandardKeys.Count, m.IcStandardKeys.Count));
        // 라벨(1) + ST 점수 + 기울기/절편(2)
        return st > 0 ? 1 + st + 2 : 0;
    }

    // ── 타이틀 ──────────────────────────────────────────────────────────────
    private static int WriteTitle(IXLWorksheet ws, int row, int nCols, TestRecordBookParsedView.Model m)
    {
        var pretty = TestRecordBookViewerService.PrettyName(m.TargetTable ?? "");
        string title = string.IsNullOrWhiteSpace(pretty) ? "시험기록부" : $"{pretty} 시험기록부";

        var c = ws.Cell(row, 1);
        c.Value = title;
        c.Style.Font.Bold = true;
        c.Style.Font.FontSize = 20;
        c.Style.Font.FontName = FontFamily;
        c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        c.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Range(row, 1, row, nCols).Merge();
        ws.Row(row).Height = 36;
        ws.Range(row, 1, row, nCols).Style.Border.BottomBorder = XLBorderStyleValues.Double;
        return row + 1;
    }

    // ── 메타 4행 ─────────────────────────────────────────────────────────────
    private static int WriteMeta(IXLWorksheet ws, int row, int nCols, TestRecordBookParsedView.Model m)
    {
        var pretty = TestRecordBookViewerService.PrettyName(m.TargetTable ?? "");
        string docNo = !string.IsNullOrWhiteSpace(pretty) && !string.IsNullOrWhiteSpace(m.AnalysisDate)
            ? $"{pretty} {m.AnalysisDate}"
            : (m.AnalysisDate ?? "");

        WriteMetaRow(ws, row++, nCols, "문서번호", docNo);
        WriteMetaRow(ws, row++, nCols, "분석방법",
            string.Join("   ", new[] { m.ES, m.Method ?? m.AnalysisMethod, m.Instrument }
                              .Where(s => !string.IsNullOrWhiteSpace(s))));
        string formulaText = ExtractFormulaTemplate(m);
        WriteMetaRow(ws, row++, nCols, "결과표시", formulaText);
        WriteMetaRow(ws, row++, nCols, "관련근거", LegalRef);
        return row;
    }

    private static void WriteMetaRow(IXLWorksheet ws, int row, int nCols, string label, string value)
    {
        var labelCell = ws.Cell(row, 1);
        labelCell.Value = label;
        labelCell.Style.Font.Bold = true;
        labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        labelCell.Style.Alignment.Vertical   = XLAlignmentVerticalValues.Center;
        labelCell.Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xEE, 0xEE);

        ws.Cell(row, 2).Value = value;
        ws.Cell(row, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        ws.Cell(row, 2).Style.Alignment.WrapText = true;
        if (nCols > 2) ws.Range(row, 2, row, nCols).Merge();

        var rng = ws.Range(row, 1, row, nCols);
        rng.Style.Border.TopBorder = XLBorderStyleValues.Hair;
        rng.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        ws.Column(1).Width = Math.Max(ws.Column(1).Width, 14);
        ws.Row(row).Height = 20;
    }

    private static string ExtractFormulaTemplate(TestRecordBookParsedView.Model m)
    {
        if (m.SampleRows.Count == 0) return "";
        int idx = m.SampleHeaders.FindIndex(h => h.Equals("계산식", StringComparison.OrdinalIgnoreCase));
        if (idx < 0) return "";
        var first = m.SampleRows[0];
        if (idx >= first.Count) return "";
        var raw = first[idx] ?? "";
        int eq = raw.LastIndexOf('=');
        return eq > 0 ? raw.Substring(0, eq).Trim() : raw.Trim();
    }

    // ── 검정곡선 ─────────────────────────────────────────────────────────────
    private static int WriteCalibration(IXLWorksheet ws, int row, int nCols, TestRecordBookParsedView.Model m)
    {
        if (m.IsBod) return WriteBodSeedSection(ws, row, nCols, m);
        if (m.IsTcic)
        {
            row = WriteCalSetTwoRow(ws, row, nCols, "TC", m.TcStandardKeys, m.TcStandardConc, m.TcStandardAbs, null,
                m.TcSlopeText, m.TcR2Text);
            row = WriteCalSetTwoRow(ws, row, nCols, "IC", m.IcStandardKeys, m.IcStandardConc, m.IcStandardAbs, null,
                m.IcSlopeText, m.IcR2Text);
            return row;
        }
        if (m.StandardKeys.Count > 0)
        {
            return WriteCalSetTwoRow(ws, row, nCols, "", m.StandardKeys, m.StandardConc, m.StandardAbs, m.StandardIstd,
                m.SlopeText, m.R2Text);
        }
        return row;
    }

    private static int WriteCalSetTwoRow(IXLWorksheet ws, int row, int nCols, string label,
        List<int> stKeys, List<string> conc, List<string> abs, List<string>? istd,
        string slopeText, string r2Text)
    {
        if (stKeys.Count == 0) return row;
        bool isVoc = istd != null && istd.Any(v => !string.IsNullOrWhiteSpace(v));

        int stCount  = stKeys.Count;
        int baseLeft = 1;
        int stStart  = 2;
        int stEnd    = stStart + stCount - 1;
        int slopeCol = stEnd + 1;
        int interCol = stEnd + 2;

        // 상단 헤더
        SetHeaderCell(ws, row, baseLeft, string.IsNullOrEmpty(label) ? "STANDARD" : $"STANDARD ({label})");
        for (int i = 0; i < stCount; i++)
            SetHeaderCell(ws, row, stStart + i, $"ST-{stKeys[i]}");
        if (slopeCol <= nCols) SetHeaderCell(ws, row, slopeCol, "기울기 (a)");
        if (interCol <= nCols) SetHeaderCell(ws, row, interCol, "절편 (b)");
        BorderRange(ws.Range(row, 1, row, nCols));
        ws.Row(row).Height = 20;
        row++;

        // 농도
        SetLabelCell(ws, row, baseLeft, "농도(mg/L)");
        for (int i = 0; i < stCount; i++)
            ws.Cell(row, stStart + i).Value = i < conc.Count ? conc[i] : "";
        var (a, b) = ParseSlopeText(slopeText);
        if (slopeCol <= nCols) ws.Cell(row, slopeCol).Value = a;
        if (interCol <= nCols) ws.Cell(row, interCol).Value = b;
        ApplyDataRowStyle(ws.Range(row, 1, row, nCols));
        CenterAlign(ws.Range(row, 1, row, nCols));
        ws.Row(row).Height = 20;
        row++;

        // abs / Area
        SetLabelCell(ws, row, baseLeft, isVoc ? "Area" : "abs");
        for (int i = 0; i < stCount; i++)
            ws.Cell(row, stStart + i).Value = i < abs.Count ? abs[i] : "";
        if (slopeCol <= nCols && !string.IsNullOrEmpty(r2Text))
        {
            ws.Cell(row, slopeCol).Value = r2Text;
            if (interCol <= nCols) ws.Range(row, slopeCol, row, interCol).Merge();
        }
        ApplyDataRowStyle(ws.Range(row, 1, row, nCols));
        CenterAlign(ws.Range(row, 1, row, nCols));
        ws.Row(row).Height = 20;
        row++;

        // ISTD (VOC)
        if (isVoc)
        {
            SetLabelCell(ws, row, baseLeft, "ISTD Area");
            for (int i = 0; i < stCount; i++)
                ws.Cell(row, stStart + i).Value = i < istd!.Count ? istd[i] : "";
            ApplyDataRowStyle(ws.Range(row, 1, row, nCols));
            CenterAlign(ws.Range(row, 1, row, nCols));
            ws.Row(row).Height = 20;
            row++;
        }
        return row;
    }

    private static (string a, string b) ParseSlopeText(string slopeText)
    {
        if (string.IsNullOrWhiteSpace(slopeText)) return ("", "");
        string a = "", b = "";
        var parts = slopeText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            if (p.StartsWith("a=", StringComparison.OrdinalIgnoreCase)) a = p.Substring(2);
            else if (p.StartsWith("b=", StringComparison.OrdinalIgnoreCase)) b = p.Substring(2);
        }
        return (a, b);
    }

    // ── BOD 식종 정보 ──────────────────────────────────────────────────────
    private static int WriteBodSeedSection(IXLWorksheet ws, int row, int nCols, TestRecordBookParsedView.Model m)
    {
        if (m.SeedRows.Count == 0) return row;

        for (int i = 0; i < m.SeedHeaders.Count && i + 1 <= nCols; i++)
            SetHeaderCell(ws, row, i + 1, m.SeedHeaders[i]);
        BorderRange(ws.Range(row, 1, row, Math.Max(nCols, m.SeedHeaders.Count)));
        ws.Row(row).Height = 20;
        row++;

        foreach (var sr in m.SeedRows)
        {
            for (int i = 0; i < sr.Count && i + 1 <= nCols; i++)
                ws.Cell(row, i + 1).Value = sr[i] ?? "";
            ApplyDataRowStyle(ws.Range(row, 1, row, Math.Max(nCols, m.SeedHeaders.Count)));
            CenterAlign(ws.Range(row, 1, row, Math.Max(nCols, m.SeedHeaders.Count)));
            ws.Row(row).Height = 20;
            row++;
        }
        return row;
    }

    // ── 시료/QC 표 ──────────────────────────────────────────────────────────
    private static int WriteSampleSection(IXLWorksheet ws, int row, int nCols,
        TestRecordBookParsedView.Model m, SampleMeta meta, string title,
        List<List<string>> rows, List<string> classes)
    {
        // 섹션 타이틀 — 시료분석결과/검정곡선의 보증 (장식 없는 강조 행)
        var titleCell = ws.Cell(row, 1);
        titleCell.Value = title;
        titleCell.Style.Font.Bold = true;
        titleCell.Style.Font.FontSize = 11;
        titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        titleCell.Style.Fill.BackgroundColor = XLColor.FromArgb(0xF2, 0xF2, 0xF2);
        ws.Range(row, 1, row, nCols).Merge();
        ws.Range(row, 1, row, nCols).Style.Border.TopBorder = XLBorderStyleValues.Medium;
        ws.Range(row, 1, row, nCols).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(row).Height = 22;
        row++;

        // 헤더 — 시료 테이블 폭이 nCols 보다 작으면 마지막 헤더가 우측 잔여 컬럼까지 병합
        int sampleCols = meta.headers.Length;
        bool needMergeTail = nCols > sampleCols;
        for (int i = 0; i < sampleCols; i++)
            SetHeaderCell(ws, row, i + 1, meta.headers[i]);
        if (needMergeTail) ws.Range(row, sampleCols, row, nCols).Merge();
        BorderRange(ws.Range(row, 1, row, nCols));
        ws.Row(row).Height = 22;
        row++;

        // 본문
        var srcIdx = MapSourceIndices(m.SampleHeaders, meta);
        int formulaIdx = m.SampleHeaders.FindIndex(h => h.Equals("계산식", StringComparison.OrdinalIgnoreCase));
        for (int rIdx = 0; rIdx < rows.Count; rIdx++)
        {
            var r = rows[rIdx];
            string cls = (rIdx < classes.Count) ? classes[rIdx] ?? "" : "";

            for (int i = 0; i < sampleCols; i++)
            {
                int from = srcIdx[i];
                ws.Cell(row, i + 1).Value = (from >= 0 && from < r.Count) ? (r[from] ?? "") : "";
                ws.Cell(row, i + 1).Style.Alignment.Horizontal =
                    (i == 0) ? XLAlignmentHorizontalValues.Left : XLAlignmentHorizontalValues.Center;
            }
            if (needMergeTail) ws.Range(row, sampleCols, row, nCols).Merge();

            bool isQc = !string.IsNullOrEmpty(cls) && cls != "S";
            if (isQc)
            {
                var color = XLColor.FromArgb(0xFA, 0xF6, 0xEE);
                ws.Range(row, 1, row, nCols).Style.Fill.BackgroundColor = color;
            }
            ApplyDataRowStyle(ws.Range(row, 1, row, nCols));
            ws.Row(row).Height = 20;
            row++;

            // 계산식 알트 ROW (메인 행 절반 높이, 전체 셀 병합)
            string formula = (formulaIdx >= 0 && formulaIdx < r.Count) ? (r[formulaIdx] ?? "") : "";
            if (!string.IsNullOrWhiteSpace(formula))
            {
                var fc = ws.Cell(row, 1);
                fc.Value = "  ↳  " + formula;
                fc.Style.Font.FontSize = 9;
                fc.Style.Font.Italic = true;
                fc.Style.Font.FontColor = XLColor.FromArgb(0x55, 0x55, 0x55);
                fc.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                fc.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                fc.Style.Fill.BackgroundColor = isQc
                    ? XLColor.FromArgb(0xF7, 0xF2, 0xE6)
                    : XLColor.FromArgb(0xFB, 0xFA, 0xF5);
                ws.Range(row, 1, row, nCols).Merge();
                var rng = ws.Range(row, 1, row, nCols);
                rng.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                rng.Style.Border.BottomBorderColor = XLColor.FromArgb(0xBB, 0xBB, 0xBB);
                ws.Row(row).Height = 12;
                row++;
            }
        }
        return row;
    }

    private static int[] MapSourceIndices(List<string> srcHeaders, SampleMeta meta)
    {
        var aliases = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["시료명"]              = new[] { "시료명" },
            ["시료번호"]            = new[] { "SN" },
            ["비고"]                = new[] { "비고" },
            ["시료량(mL)"]          = new[] { "시료량" },
            ["흡광도"]              = new[] { "흡광도", "AU" },
            ["Area"]                = new[] { "Area", "흡광도", "AU" },
            ["ISTD Resp."]          = new[] { "ISTD" },
            ["ISTD"]                = new[] { "ISTD" },
            ["희석배수"]            = new[] { "희석배수" },
            ["검량선으로 구한 a"]   = new[] { "농도", "계산농도" },
            ["Result(mg/L)"]        = new[] { "결과", "결과값" },
            ["Result(개/mL)"]       = new[] { "결과", "결과값" },
            ["전무게(g)"]           = new[] { "전무게" },
            ["후무게(g)"]           = new[] { "후무게" },
            ["무게차(g)"]           = new[] { "무게차" },
            ["A"]                   = new[] { "A" },
            ["B"]                   = new[] { "B" },
            ["TCAU"]                = new[] { "TCAU" },
            ["TC con(mg/L)"]        = new[] { "TCcon", "TC_농도" },
            ["ICAU"]                = new[] { "ICAU" },
            ["IC con(mg/L)"]        = new[] { "ICcon", "IC_농도" },
            ["D1(15min DO)"]        = new[] { "D1" },
            ["D2(5Day DO)"]         = new[] { "D2" },
            ["f(x/y)"]              = new[] { "F_xy", "Fxy", "f(x/y)" },
        };

        var result = new int[meta.headers.Length];
        for (int i = 0; i < meta.headers.Length; i++)
        {
            int found = -1;
            if (aliases.TryGetValue(meta.headers[i], out var keys))
            {
                foreach (var k in keys)
                {
                    int idx = srcHeaders.FindIndex(h => h.Equals(k, StringComparison.OrdinalIgnoreCase));
                    if (idx >= 0) { found = idx; break; }
                }
            }
            else
            {
                found = srcHeaders.FindIndex(h => h.Equals(meta.headers[i], StringComparison.OrdinalIgnoreCase));
            }
            result[i] = found;
        }
        return result;
    }

    // ── 비고/메모 박스 (Memo 있을 때만, 짧은 placeholder 값 거름) ───────────
    private static int WriteRemarksBox(IXLWorksheet ws, int row, int nCols, TestRecordBookParsedView.Model m)
    {
        var memo = (m.Memo ?? "").Trim();
        if (memo.Length < 4) return row;   // "1", "test" 등 의미 없는 짧은 값 제외

        var lbl = ws.Cell(row, 1);
        lbl.Value = "비   고";
        lbl.Style.Font.Bold = true;
        lbl.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        lbl.Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xEE, 0xEE);
        ws.Cell(row, 2).Value = m.Memo;
        ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        ws.Cell(row, 2).Style.Alignment.WrapText = true;
        ws.Range(row, 2, row, nCols).Merge();
        ws.Range(row, 1, row, nCols).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Range(row, 1, row, nCols).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(row).Height = 28;
        return row + 1;
    }

    // ── 결재라인 — 콘텐츠 직후 (페이지 자연 흐름) ───────────────────────────
    private static int WriteSignature(IXLWorksheet ws, int row, int nCols)
    {
        const int sigHeight = 4;
        // 좌우 2박스: 분석담당자 / 기술책임자 — 전체 폭 사용
        int rightCol = nCols;
        int leftCol  = 1;
        int midCol   = leftCol + (nCols / 2) - 1;

        // 라벨 행
        ws.Cell(row, leftCol).Value = "분 석 담 당 자";
        ws.Cell(row, leftCol).Style.Font.Bold = true;
        ws.Cell(row, leftCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell(row, leftCol).Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xEE, 0xEE);
        ws.Range(row, leftCol, row, midCol).Merge();

        ws.Cell(row, midCol + 1).Value = "기 술 책 임 자";
        ws.Cell(row, midCol + 1).Style.Font.Bold = true;
        ws.Cell(row, midCol + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell(row, midCol + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xEE, 0xEE);
        ws.Range(row, midCol + 1, row, rightCol).Merge();
        ws.Row(row).Height = 22;

        // 서명 자리 3행
        for (int r2 = row + 1; r2 < row + sigHeight; r2++)
        {
            ws.Range(r2, leftCol, r2, midCol).Merge();
            ws.Range(r2, midCol + 1, r2, rightCol).Merge();
            ws.Row(r2).Height = (r2 == row + sigHeight - 1) ? 18 : 22;

            if (r2 == row + sigHeight - 1)
            {
                ws.Cell(r2, leftCol).Value = "(서명)";
                ws.Cell(r2, leftCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(r2, leftCol).Style.Font.Italic = true;
                ws.Cell(r2, leftCol).Style.Font.FontColor = XLColor.FromArgb(0x99, 0x99, 0x99);

                ws.Cell(r2, midCol + 1).Value = "(서명)";
                ws.Cell(r2, midCol + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(r2, midCol + 1).Style.Font.Italic = true;
                ws.Cell(r2, midCol + 1).Style.Font.FontColor = XLColor.FromArgb(0x99, 0x99, 0x99);
            }
        }
        // 수평 라인만 — 위/아래 + 라벨/서명 사이만
        ws.Range(row, leftCol, row, rightCol).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Range(row, leftCol, row, rightCol).Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        ws.Range(row + sigHeight - 1, leftCol, row + sigHeight - 1, rightCol).Style.Border.BottomBorder = XLBorderStyleValues.Thin;

        return row;
    }

    // ── 컬럼 폭 ──────────────────────────────────────────────────────────
    private static void ApplyColumnWidths(IXLWorksheet ws, int nCols, SampleMeta meta)
    {
        for (int c = 1; c <= meta.widths.Length && c <= nCols; c++)
            ws.Column(c).Width = meta.widths[c - 1];
        for (int c = meta.widths.Length + 1; c <= nCols; c++)
            ws.Column(c).Width = 12;
    }

    // ── 셀 스타일 헬퍼 ─────────────────────────────────────────────────────
    private static void SetHeaderCell(IXLWorksheet ws, int row, int col, string text)
    {
        var c = ws.Cell(row, col);
        c.Value = text;
        c.Style.Font.Bold = true;
        c.Style.Font.FontSize = 10;
        c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        c.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        c.Style.Alignment.WrapText = true;
        c.Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xEE, 0xEE);
    }

    private static void SetLabelCell(IXLWorksheet ws, int row, int col, string text)
    {
        var c = ws.Cell(row, col);
        c.Value = text;
        c.Style.Font.Bold = true;
        c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        c.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        c.Style.Fill.BackgroundColor = XLColor.FromArgb(0xF6, 0xF4, 0xEE);
    }

    // 수평 라인만 — 수직 격자(네모칸) 금지
    private static void BorderRange(IXLRange rng)
    {
        rng.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        rng.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
    }

    private static void ApplyDataRowStyle(IXLRange rng)
    {
        rng.Style.Border.TopBorder = XLBorderStyleValues.Hair;
        rng.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        rng.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void CenterAlign(IXLRange rng)
    {
        rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        rng.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    // ── 페이지 머리글에 워터마크 이미지 박기 (OpenXML 후처리) ─────────────
    private static void TryInjectHeaderLogo(string xlsxPath)
    {
        try
        {
            string[] candidates =
            {
                Path.Combine(Directory.GetCurrentDirectory(), "Assets", "icons", "renewus_vertical_black.png"),
                Path.Combine(AppContext.BaseDirectory,         "Assets", "icons", "renewus_vertical_black.png"),
            };
            string? logoPath = candidates.FirstOrDefault(File.Exists);
            if (logoPath == null) return;

            string fadedPath = MakeFadedLogo(logoPath, 0.10f);

            using var doc = SpreadsheetDocument.Open(xlsxPath, true);
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var ws = wsPart.Worksheet;

            // 1) VML 드로잉 파트 — relationship ID 는 OpenXML 이 자동 생성
            var vmlPart = wsPart.AddNewPart<VmlDrawingPart>();
            string vmlRelId = wsPart.GetIdOfPart(vmlPart);

            // 2) VML 안에 이미지 파트 첨부
            var imgPart = vmlPart.AddImagePart(ImagePartType.Png);
            using (var img = File.OpenRead(fadedPath))
                imgPart.FeedData(img);
            string imgRelId = vmlPart.GetIdOfPart(imgPart);

            // 3) VML XML 작성 — 가운데(CH) 헤더 위치
            string vml = $@"<xml xmlns:v=""urn:schemas-microsoft-com:vml""
 xmlns:o=""urn:schemas-microsoft-com:office:office""
 xmlns:x=""urn:schemas-microsoft-com:office:excel"">
 <o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" data=""1""/></o:shapelayout>
 <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t""
  path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
  <v:stroke joinstyle=""miter""/>
  <v:formulas>
   <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
   <v:f eqn=""sum @0 1 0""/>
   <v:f eqn=""sum 0 0 @1""/>
   <v:f eqn=""prod @2 1 2""/>
   <v:f eqn=""prod @3 21600 pixelWidth""/>
   <v:f eqn=""prod @3 21600 pixelHeight""/>
   <v:f eqn=""sum @0 0 1""/>
   <v:f eqn=""prod @6 1 2""/>
   <v:f eqn=""prod @7 21600 pixelWidth""/>
   <v:f eqn=""sum @8 21600 0""/>
   <v:f eqn=""prod @7 21600 pixelHeight""/>
   <v:f eqn=""sum @10 21600 0""/>
  </v:formulas>
  <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
  <o:lock v:ext=""edit"" aspectratio=""t""/>
 </v:shapetype>
 <v:shape id=""CH"" o:spid=""_x0000_s1025"" type=""#_x0000_t75""
  style=""position:absolute;margin-left:0;margin-top:0;width:200pt;height:60pt;z-index:1""
  o:allowincell=""f"">
  <v:imagedata o:relid=""{imgRelId}"" o:title=""renewus""/>
  <o:lock v:ext=""edit"" rotation=""t""/>
 </v:shape>
</xml>";
            using (var s = vmlPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s, new System.Text.UTF8Encoding(false)))
                w.Write(vml);

            // 4) HeaderFooter — OddHeader 가운데에 그래픽(&G) 배치
            var hf = ws.GetFirstChild<HeaderFooter>();
            if (hf == null)
            {
                hf = new HeaderFooter();
                var pageSetup = ws.GetFirstChild<PageSetup>();
                if (pageSetup != null) ws.InsertAfter(hf, pageSetup);
                else ws.AppendChild(hf);
            }
            hf.RemoveAllChildren<OddHeader>();
            hf.PrependChild(new OddHeader { Text = "&C&G" });

            // 5) LegacyDrawingHeaderFooter — schema 순서: picture/oleObjects/controls/
            //    webPublishItems/tableParts/extLst 보다 앞이어야 함
            ws.RemoveAllChildren<LegacyDrawingHeaderFooter>();
            var legacyHf = new LegacyDrawingHeaderFooter { Id = vmlRelId };
            DocumentFormat.OpenXml.OpenXmlElement? successor =
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<Picture>() ??
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<OleObjects>() ??
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<Controls>() ??
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<WebPublishItems>() ??
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<TableParts>() ??
                (DocumentFormat.OpenXml.OpenXmlElement?)ws.GetFirstChild<WorksheetExtensionList>();
            if (successor != null) ws.InsertBefore(legacyHf, successor);
            else ws.AppendChild(legacyHf);

            ws.Save();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[HeaderLogo] {ex.Message}");
        }
    }

    private static string MakeFadedLogo(string srcPath, float alpha)
    {
        var cacheDir = Path.Combine(AppPaths.WritableDataRoot, "PrintCache");
        Directory.CreateDirectory(cacheDir);
        var dst = Path.Combine(cacheDir, $"watermark_a{(int)(alpha*100)}.png");
        if (File.Exists(dst)) return dst;

        using var input = SKBitmap.Decode(srcPath);
        if (input == null) return srcPath;
        var info = new SKImageInfo(input.Width, input.Height);
        using var surface = SKSurface.Create(info);
        surface.Canvas.Clear(SKColors.Transparent);
        using var paint = new SKPaint
        {
            ColorFilter = SKColorFilter.CreateColorMatrix(new float[]
            {
                1, 0, 0, 0, 0,
                0, 1, 0, 0, 0,
                0, 0, 1, 0, 0,
                0, 0, 0, alpha, 0,
            })
        };
        surface.Canvas.DrawBitmap(input, 0, 0, paint);
        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 95);
        using var fs = File.OpenWrite(dst);
        data.SaveTo(fs);
        return dst;
    }
}
