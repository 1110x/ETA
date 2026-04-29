using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ETA.Services.Common;
using A   = DocumentFormat.OpenXml.Drawing;
using DW  = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 생태독성 (물벼룩 급성독성) 시험기록부 Word(.docx) 출력.
/// 시각언어: 가로선 표, 8pt 본문, 9pt 섹션, renewus 워터마크.
/// 시료 1건 = 1 페이지.
/// </summary>
public static class EcotoxicityWordExporter
{
    private const string FontFamily = "맑은 고딕";

    // 페이지 폭 = A4 8.27" - 좌우 0.5"*2 = 7.27" ≈ 10400 dxa
    private const int PageDxa        = 10400;
    private const int MinColWidth    = 700;
    private const int MaxColWidth    = 6000;
    private const int CellPaddingDxa = 200;
    private const int AsciiCharDxa   = 110;
    private const int CjkCharDxa     = 175;

    public sealed record Record(
        string Date, string TestNo, string Species, string Toxicant, string SampleName,
        EcotoxicityService.EcotoxResult? TskResult,
        EcotoxicityService.EcotoxResult? ProbitResult,
        double[] Conc, int[] Org, int[] Mort,
        int CtrlOrg, int CtrlMort,
        double TestTemperature, double TestPH,
        double SampleTemperature, double SamplePH, double SampleDO,
        string Duration, string EcCalculationMethod,
        string Observations, string AnalystName);

    [ThreadStatic] private static SectionProperties? _sectPr;

    public static string Export(IList<Record> records)
    {
        Log($"Export 시작 — {records.Count}건");
        var dir = Path.Combine(AppPaths.WritableDataRoot, "PrintCache");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, $"생태독성_시험기록부_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            BuildStyles(main);
            BuildHeaderWithWatermark(main);

            for (int i = 0; i < records.Count; i++)
            {
                var rec = records[i];
                Log($"record[{i}] sample={rec.SampleName}");

                if (i == 0) body.Append(BuildTitle());
                else        body.Append(PageBreak());

                Section(body, "시험기본정보");
                body.Append(BuildBasicInfo(rec));
                body.Append(Spacer());

                Section(body, "시험조건 (ES 04704.1c 기준)");
                body.Append(BuildConditions(rec));
                body.Append(Spacer());

                Section(body, "대조군 정보");
                body.Append(BuildControl(rec));
                body.Append(Spacer());

                Section(body, "농도별 독성 시험 데이터");
                body.Append(BuildConcentration(rec));
                body.Append(Spacer());

                Section(body, "독성 분석 결과");
                body.Append(BuildResults(rec));
                body.Append(Spacer());

                Section(body, "분석법 적용 근거");
                body.Append(BuildRationale(rec));
                body.Append(Spacer());

                if (rec.TskResult != null && rec.Conc != null && rec.Conc.Length > 0)
                {
                    try
                    {
                        var png = EcotoxicityChartGenerator.Generate(
                            rec.Conc, rec.Org, rec.Mort,
                            rec.TskResult.EC50, rec.TskResult.LowerCI, rec.TskResult.UpperCI,
                            rec.TskResult.TU, rec.TskResult.Method);
                        Section(body, "용량-반응 곡선");
                        body.Append(BuildChart(main, png));
                    }
                    catch (Exception cex) { Log($"  ✗ 차트 실패: {cex.Message}"); }
                }

                if (!string.IsNullOrWhiteSpace(rec.Observations))
                {
                    Section(body, "시험 중 관찰사항");
                    body.Append(P(rec.Observations, 9));
                }
            }

            body.Append(BuildSignature());
            if (_sectPr != null) { body.Append(_sectPr); _sectPr = null; }

            main.Document.Save();
        }
        Log($"Export 완료 — {new FileInfo(path).Length} bytes");
        return path;
    }

    // ─── 스타일 / 헤더 ──────────────────────────────────────────────────────
    private static void BuildStyles(MainDocumentPart main)
    {
        var sp = main.AddNewPart<StyleDefinitionsPart>();
        sp.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                        new FontSize { Val = "16" })),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { After = "0", Before = "0", Line = "200", LineRule = LineSpacingRuleValues.Auto }))));
        sp.Styles.Save();
    }

    private static void BuildHeaderWithWatermark(MainDocumentPart main)
    {
        var hp = main.AddNewPart<HeaderPart>();
        string? relId = null;
        try
        {
            Stream? logo = null;
            try { logo = Avalonia.Platform.AssetLoader.Open(new Uri("avares://ETA/Assets/icons/renewus_vertical_black.png")); }
            catch
            {
                var p = Path.Combine(AppPaths.RootPath, "Assets", "icons", "renewus_vertical_black.png");
                if (File.Exists(p)) logo = File.OpenRead(p);
            }
            if (logo != null)
            {
                using (logo)
                {
                    var img = hp.AddImagePart(ImagePartType.Png);
                    img.FeedData(logo);
                    relId = hp.GetIdOfPart(img);
                }
            }
        }
        catch { }

        string xml = relId != null ? WatermarkImageXml(relId) : WatermarkTextXml("ETA");
        using (var s = hp.GetStream(FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(s, new System.Text.UTF8Encoding(false))) sw.Write(xml);

        _sectPr = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(hp) },
            new PageSize { Width = 11906U, Height = 16838U },
            new PageMargin { Top = 540, Right = 540, Bottom = 540, Left = 540, Header = 280, Footer = 280, Gutter = 0 });
    }

    private static string WatermarkImageXml(string relId) =>
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
       xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
       xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
       xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
  <w:p><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>
    <wp:anchor distT=""0"" distB=""0"" distL=""0"" distR=""0"" simplePos=""0""
               relativeHeight=""251658240"" behindDoc=""1"" locked=""0"" layoutInCell=""1"" allowOverlap=""1"">
      <wp:simplePos x=""0"" y=""0""/>
      <wp:positionH relativeFrom=""page""><wp:align>center</wp:align></wp:positionH>
      <wp:positionV relativeFrom=""page""><wp:align>center</wp:align></wp:positionV>
      <wp:extent cx=""2197100"" cy=""1524000""/>
      <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
      <wp:wrapNone/>
      <wp:docPr id=""1"" name=""WatermarkLogo""/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=""1""/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id=""1"" name=""renewus""/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed=""" + relId + @"""><a:alphaModFix amt=""40000""/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""2197100"" cy=""1524000""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:anchor>
  </w:drawing></w:r></w:p>
</w:hdr>";

    private static string WatermarkTextXml(string text) =>
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:v=""urn:schemas-microsoft-com:vml""
       xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <w:p><w:r><w:rPr><w:noProof/></w:rPr>
    <w:pict><v:shape type=""#_x0000_t136""
        style=""position:absolute;margin-left:0;margin-top:0;width:173pt;height:86pt;rotation:-45;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page""
        fillcolor=""#808080"" stroked=""f"">
      <v:fill opacity="".4""/>
      <v:textpath style=""font-family:&quot;맑은 고딕&quot;;font-size:1pt"" fitpath=""t"" string=""" + System.Net.WebUtility.HtmlEncode(text) + @"""/>
    </v:shape></w:pict>
  </w:r></w:p>
</w:hdr>";

    // ─── 타이틀 / 섹션 / 단락 ───────────────────────────────────────────────
    private static Paragraph BuildTitle() =>
        new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "40" }),
            new Run(new RunProperties(new Bold(), new FontSize { Val = "26" }),
                new Text("물벼룩을 이용한 급성 독성 시험기록부")),
            new Run(new RunProperties(new Italic(), new FontSize { Val = "14" }, new Color { Val = "555555" }),
                new Break(), new Text("ES 04704.1c (수질오염공정시험기준)")));

    /// <summary>1행 표 — 좁은 파란 셀(액센트바) + 텍스트 셀. 막대 길이 = 텍스트 행 높이.</summary>
    private static void Section(Body body, string title)
    {
        var t = new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" },
                new TableBorders(
                    new TopBorder    { Val = BorderValues.None },
                    new LeftBorder   { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.None },
                    new RightBorder  { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.None },
                    new InsideVerticalBorder   { Val = BorderValues.None }),
                new TableLayout { Type = TableLayoutValues.Fixed }));
        var grid = new TableGrid();
        grid.Append(new GridColumn { Width = "60" });
        grid.Append(new GridColumn { Width = "10340" });
        t.Append(grid);

        var barCp = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "60" },
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "1F4E79" },
            new TableCellMargin(
                new LeftMargin  { Width = "0", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "0", Type = TableWidthUnitValues.Dxa }));
        var barCell = new TableCell(barCp, new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0" }),
            new Run(new Text(""))));

        var txtCp = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "10340" },
            new TableCellMargin(
                new LeftMargin  { Width = "120", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "0",   Type = TableWidthUnitValues.Dxa }),
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        var txtCell = new TableCell(txtCp, new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0" }),
            new Run(new RunProperties(new Bold(), new FontSize { Val = "16" }, new Color { Val = "1F4E79" }),
                new Text(title))));

        t.Append(new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 160U, HeightType = HeightRuleValues.AtLeast }),
            barCell, txtCell));

        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "20", After = "0" })));
        body.Append(t);
        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "10" })));
    }

    private static Paragraph Spacer() =>
        new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0" }));

    private static Paragraph PageBreak() =>
        new Paragraph(new Run(new Break { Type = BreakValues.Page }));

    private static Paragraph P(string text, int pt) =>
        new Paragraph(new Run(
            new RunProperties(new FontSize { Val = (pt * 2).ToString() }),
            new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));

    // ─── 시험일자 포맷 — 채취일 + 1일 = 시작일, +duration 시간 = 종료일.
    //     예: collection=2026-04-22, duration="24 H" → "2026-04-23~24(24hr)" ───
    private static string FormatTestDateSpan(string collectionDate, string duration)
    {
        if (!DateTime.TryParse(collectionDate, out var dt))
            return collectionDate ?? "";
        int hours = ParseHours(duration);
        var start = dt.AddDays(1);
        var end   = start.AddHours(hours);
        return start.Date == end.Date
            ? $"{start:yyyy-MM-dd}({hours}hr)"
            : $"{start:yyyy-MM-dd}~{end:dd}({hours}hr)";
    }

    private static int ParseHours(string duration)
    {
        if (string.IsNullOrWhiteSpace(duration)) return 24;
        var m = System.Text.RegularExpressions.Regex.Match(duration, @"(\d+)");
        return m.Success && int.TryParse(m.Groups[1].Value, out var n) ? n : 24;
    }

    // ─── 시험기본정보 ───────────────────────────────────────────────────────
    private static Table BuildBasicInfo(Record r)
    {
        var row = new[] {
            ("시험일자", FormatTestDateSpan(r.Date, r.Duration)),
            ("시험번호", r.TestNo),
            ("시료명",   r.SampleName),
        };
        return BuildPackedTable(new[] { row });
    }

    // ─── 시험조건 ────────────────────────────────────────────────────────────
    private static Table BuildConditions(Record r)
    {
        // 실제 계산된 결과의 Method 를 우선 (EcCalculationMethod 는 폼 기본값 "Probit" 잔존 가능)
        string method = !string.IsNullOrWhiteSpace(r.TskResult?.Method) ? r.TskResult!.Method
                      : !string.IsNullOrWhiteSpace(r.EcCalculationMethod) ? r.EcCalculationMethod
                      : "—";
        var row1 = new[] {
            ("시험온도(°C)", $"{Fmt(r.TestTemperature)} (기준 20±2)"),
            ("시험 pH",      $"{Fmt(r.TestPH)} (기준 7.6~8.0)"),
            ("EC50 계산방법", method),
        };
        var row2 = new[] {
            ("시료온도(°C)",   Fmt(r.SampleTemperature)),
            ("시료 pH",        Fmt(r.SamplePH)),
            ("용존산소(mg/L)", Fmt(r.SampleDO)),
        };
        var t = BuildPackedTable(new[] { row1, row2 });
        if (!string.IsNullOrWhiteSpace(r.AnalystName))
        {
            var single = new[] { ("분석자", r.AnalystName) };
            int[] w = PackedColWidths(new[] { single });
            AddPackedRow(t, single, w);
        }
        return t;
    }

    // ─── 대조군 ─────────────────────────────────────────────────────────────
    private static Table BuildControl(Record r)
    {
        double rate = r.CtrlOrg > 0 ? (r.CtrlMort * 100.0 / r.CtrlOrg) : 0;
        var cells = new[] {
            ("대조생물수",  r.CtrlOrg.ToString()),
            ("대조 치사수", r.CtrlMort.ToString()),
            ("치사율",      $"{rate:F1}%"),
        };
        return BuildPackedTable(new[] { cells }, Align.Center);
    }

    // ─── 농도별 시험 데이터 ─────────────────────────────────────────────────
    private static Table BuildConcentration(Record r)
    {
        var headers = new[] { "농도(%)", "생물수", "유영저해 및 치사", "치사율(%)" };
        var rows = new List<string[]> { headers };
        for (int i = 0; i < r.Conc.Length; i++)
        {
            int org  = i < r.Org.Length  ? r.Org[i]  : 0;
            int mort = i < r.Mort.Length ? r.Mort[i] : 0;
            double rate = org > 0 ? (mort * 100.0 / org) : 0;
            rows.Add(new[] { Fmt(r.Conc[i]), org.ToString(), mort.ToString(), rate.ToString("F1") });
        }

        int[] widths = ColWidths(rows.ToArray());
        var t = NewTable();
        AddGrid(t, widths);

        // 헤더
        var hdr = new TableRow();
        for (int i = 0; i < headers.Length; i++)
            hdr.Append(Cell(headers[i], widths[i], isLabel: true, align: Align.Center));
        t.Append(hdr);

        // 데이터 — 가운데 정렬
        for (int rIdx = 1; rIdx < rows.Count; rIdx++)
        {
            var tr = new TableRow();
            for (int c = 0; c < headers.Length; c++)
                tr.Append(Cell(rows[rIdx][c], widths[c], align: Align.Center));
            t.Append(tr);
        }
        return t;
    }

    // ─── 독성 분석 결과 ─────────────────────────────────────────────────────
    private static Table BuildResults(Record r)
    {
        // 헤더 1 + 데이터 1~2 행. 95% CI 합치고 Trim% 가 있을 때만 컬럼 추가.
        bool showTrim = (r.TskResult?.TrimPercent ?? -1) >= 0
                     || (r.ProbitResult?.TrimPercent ?? -1) >= 0;
        var headers = showTrim
            ? new[] { "분석법", "EC50", "TU", "95% CI", "Trim%" }
            : new[] { "분석법", "EC50", "TU", "95% CI" };

        var data = new List<string[]>();
        if (r.TskResult    != null) data.Add(ResultRow("TSK",    r.TskResult,    showTrim));
        if (r.ProbitResult != null) data.Add(ResultRow("Probit", r.ProbitResult, showTrim));
        if (data.Count == 0)
        {
            var row = new string[headers.Length];
            row[0] = "—"; row[1] = "분석 결과 없음";
            for (int i = 2; i < row.Length; i++) row[i] = "";
            data.Add(row);
        }

        var allRows = new List<string[]> { headers };
        allRows.AddRange(data);
        int[] widths = ColWidths(allRows.ToArray());

        var t = NewTable();
        AddGrid(t, widths);

        var hdr = new TableRow();
        for (int i = 0; i < headers.Length; i++)
            hdr.Append(Cell(headers[i], widths[i], isLabel: true, align: Align.Center));
        t.Append(hdr);

        foreach (var d in data)
        {
            var tr = new TableRow();
            for (int i = 0; i < headers.Length; i++)
            {
                bool firstCol = i == 0;
                tr.Append(Cell(d[i], widths[i], isLabel: firstCol, align: Align.Center));
            }
            t.Append(tr);
        }
        return t;
    }

    private static string[] ResultRow(string method, EcotoxicityService.EcotoxResult res, bool showTrim)
    {
        string ci = $"[{Fmt(res.LowerCI)} ~ {Fmt(res.UpperCI)}]";
        string trim = res.TrimPercent >= 0 ? $"{res.TrimPercent:F0}%" : "—";
        return showTrim
            ? new[] { method, Fmt(res.EC50), Fmt(res.TU), ci, trim }
            : new[] { method, Fmt(res.EC50), Fmt(res.TU), ci };
    }

    // ─── 분석법 적용 근거 ───────────────────────────────────────────────────
    private static Table BuildRationale(Record r)
    {
        int total = r.Conc?.Length ?? 0;
        int partial = 0, allDead = 0, allLive = 0;
        for (int i = 0; i < total; i++)
        {
            int org  = i < r.Org.Length  ? r.Org[i]  : 0;
            int mort = i < r.Mort.Length ? r.Mort[i] : 0;
            if (org <= 0) continue;
            double pct = mort * 100.0 / org;
            if      (pct <= 0)   allLive++;
            else if (pct >= 100) allDead++;
            else                 partial++;
        }

        bool hasTsk    = r.TskResult    != null;
        bool hasProbit = r.ProbitResult != null;
        // 실제 계산된 결과의 Method 우선 (TSK/Probit). 폴백: EcCalculationMethod, 그 다음 가용 결과 추정.
        string used = !string.IsNullOrWhiteSpace(r.TskResult?.Method)
            ? r.TskResult!.Method
            : !string.IsNullOrWhiteSpace(r.EcCalculationMethod)
                ? r.EcCalculationMethod
                : (hasProbit ? "Probit" : (hasTsk ? "TSK" : "—"));

        bool isProbit = used.Contains("Probit", StringComparison.OrdinalIgnoreCase);
        bool isTsk    = used.Contains("TSK", StringComparison.OrdinalIgnoreCase)
                     || used.Contains("Spearman", StringComparison.OrdinalIgnoreCase);
        string reason = isProbit
            ? (partial >= 2 ? "부분치사 농도점 2개 이상 → Probit 회귀 적용." : "사용자 지정으로 Probit 회귀 적용.")
            : isTsk
                ? (partial >= 2 ? "비매개적 Spearman-Kärber(TSK) 적용." : "부분치사 농도점 부족 → 비매개적 Spearman-Kärber(TSK) 적용.")
                : "데이터 부족으로 EC50 산출 불가.";
        if (isTsk && r.TskResult != null && r.TskResult.TrimPercent > 0)
            reason += $" Trim {r.TskResult.TrimPercent:F0}%.";

        var row1 = new[] { ("측정 농도점", $"{total}개 (부분치사 {partial}, 100% 사망 {allDead}, 사망 0% {allLive})") };
        var row2 = new[] { ("적용 분석법", $"{used} — {reason}") };
        return BuildPackedTable(new[] { row1, row2 });
    }

    // ─── 차트 PNG 임베드 ────────────────────────────────────────────────────
    private static Paragraph BuildChart(MainDocumentPart main, byte[] png)
    {
        var imgPart = main.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(png)) imgPart.FeedData(ms);
        string relId = main.GetIdOfPart(imgPart);

        long cx = 6000000;                      // ≈ 472pt 폭 (≈ 6.56" — A4 본문 폭 거의 전체)
        long cy = (long)(cx * 280.0 / 640.0);   // 비율 압축 (640:280)

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = 100U, Name = "Chart" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 100U, Name = "Chart" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };

        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0", Line = "200", LineRule = LineSpacingRuleValues.Auto }),
            new Run(new RunProperties(new NoProof()), new Drawing(inline)));
    }

    // ─── 서명 (페이지 하단 고정) ────────────────────────────────────────────
    private static Table BuildSignature()
    {
        var t = NewTable();
        var tp = t.GetFirstChild<TableProperties>();
        if (tp != null)
            tp.InsertAt(new TablePositionProperties
            {
                VerticalAnchor = VerticalAnchorValues.Margin,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            }, 0);
        AddGrid(t, 3467, 3467, 3466);
        var hdr = new TableRow(new TableRowProperties(new TableRowHeight { Val = 280U, HeightType = HeightRuleValues.AtLeast }));
        hdr.Append(Cell("작성자", 3467, isLabel: true, align: Align.Center));
        hdr.Append(Cell("검토자", 3467, isLabel: true, align: Align.Center));
        hdr.Append(Cell("승인자", 3466, isLabel: true, align: Align.Center));
        t.Append(hdr);
        var sig = new TableRow(new TableRowProperties(new TableRowHeight { Val = 500U, HeightType = HeightRuleValues.AtLeast }));
        sig.Append(Cell("", 3467, align: Align.Center));
        sig.Append(Cell("", 3467, align: Align.Center));
        sig.Append(Cell("", 3466, align: Align.Center));
        t.Append(sig);
        return t;
    }

    // ─── 공통 빌드 유틸 ─────────────────────────────────────────────────────

    /// <summary>(라벨, 값) 페어들로 구성된 가로선 표.
    /// 셀 안에서 라벨은 굵게, 값은 일반 — 한 셀 안에서 시각 구분.
    /// 기본 좌측 정렬, 가운데 정렬은 align 파라미터로.</summary>
    private static Table BuildPackedTable((string label, string value)[][] rows, Align align = Align.Left)
    {
        int[] widths = PackedColWidths(rows);
        var t = NewTable();
        AddGrid(t, widths);
        foreach (var r in rows) AddPackedRow(t, r, widths, align);
        return t;
    }

    private static void AddPackedRow(Table t, (string label, string value)[] cells, int[] widths, Align align = Align.Left)
    {
        var tr = new TableRow();
        for (int i = 0; i < cells.Length; i++)
        {
            int w = i < widths.Length ? widths[i] : MinColWidth;
            tr.Append(PackedCell(cells[i].label, cells[i].value, w, align));
        }
        t.Append(tr);
    }

    /// <summary>(라벨, 값) row 들의 컬럼별 max 폭 측정 → 페이지 폭 비례 분배.</summary>
    private static int[] PackedColWidths((string label, string value)[][] rows)
    {
        var asText = rows.Select(r => r.Select(c => $"{c.label}: {c.value}").ToArray()).ToArray();
        return ColWidths(asText);
    }

    /// <summary>한 셀 안에 굵은 라벨 + 일반 값. align 으로 정렬 지정.</summary>
    private static TableCell PackedCell(string label, string value, int width, Align align = Align.Left)
    {
        var cp = new TableCellProperties();
        cp.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() });
        cp.Append(new TableCellBorders(new LeftBorder { Val = BorderValues.None }, new RightBorder { Val = BorderValues.None }));
        cp.Append(new TableCellMargin(
            new LeftMargin  { Width = "120", Type = TableWidthUnitValues.Dxa },
            new RightMargin { Width = "120", Type = TableWidthUnitValues.Dxa }));
        cp.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        var pp = new ParagraphProperties();
        pp.Append(new Justification { Val = align switch
        {
            Align.Right  => JustificationValues.Right,
            Align.Center => JustificationValues.Center,
            _            => JustificationValues.Left,
        } });

        var labelRun = new Run(
            new RunProperties(new Bold(), new FontSize { Val = "16" }),
            new Text(label + " ") { Space = SpaceProcessingModeValues.Preserve });
        var valueRun = new Run(
            new RunProperties(new FontSize { Val = "16" }),
            new Text(value ?? "") { Space = SpaceProcessingModeValues.Preserve });

        return new TableCell(cp, new Paragraph(pp, labelRun, valueRun));
    }

    /// <summary>각 컬럼의 max 글자 폭 측정 + 페이지 폭에 비례 분배.</summary>
    private static int[] ColWidths(string[][] rows)
    {
        if (rows.Length == 0 || rows[0].Length == 0) return Array.Empty<int>();
        int n = rows.Max(r => r.Length);

        var measured = new int[n];
        for (int c = 0; c < n; c++)
        {
            int w = 0;
            foreach (var r in rows)
                if (c < r.Length) w = Math.Max(w, MeasureCharWidth(r[c] ?? ""));
            measured[c] = Math.Clamp(w + CellPaddingDxa, MinColWidth, MaxColWidth);
        }

        int sum = measured.Sum();
        var widths = new int[n];
        if (sum == 0)
        {
            for (int i = 0; i < n; i++) widths[i] = PageDxa / n;
            return widths;
        }
        double ratio = (double)PageDxa / sum;
        int used = 0;
        for (int i = 0; i < n - 1; i++)
        {
            widths[i] = Math.Max(MinColWidth, (int)Math.Round(measured[i] * ratio));
            used += widths[i];
        }
        widths[n - 1] = Math.Max(MinColWidth, PageDxa - used);
        return widths;
    }

    private static int MeasureCharWidth(string s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        if (s.Contains('\n') || s.Contains('\r'))
            return s.Replace("\r\n", "\n").Split('\n').Max(MeasureCharWidth);
        int w = 0;
        foreach (var ch in s)
        {
            if ((ch >= 0xAC00 && ch <= 0xD7A3) || (ch >= 0x4E00 && ch <= 0x9FFF)
             || (ch >= 0x3040 && ch <= 0x30FF) || (ch >= 0xFF00 && ch <= 0xFFEF))
                w += CjkCharDxa;
            else
                w += AsciiCharDxa;
        }
        return w;
    }

    private static Table NewTable() =>
        new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new LeftBorder   { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new RightBorder  { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "BBBBBB" },
                    new InsideVerticalBorder   { Val = BorderValues.None }),
                new TableLayout { Type = TableLayoutValues.Fixed }));

    private static void AddGrid(Table t, params int[] widths)
    {
        var grid = new TableGrid();
        foreach (var w in widths) grid.Append(new GridColumn { Width = w.ToString() });
        t.Append(grid);
    }

    private enum Align { Left, Center, Right }

    /// <summary>표준 셀 — 가운데 수직정렬, 좌/우 테두리 없음, 폰트 8pt.
    /// 정렬은 Align 으로 단일 지정. 라벨인 경우 굵게.</summary>
    private static TableCell Cell(string text, int width, bool isLabel = false,
        Align align = Align.Left, string? shadingFill = null)
    {
        var cp = new TableCellProperties();
        cp.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() });
        cp.Append(new TableCellBorders(new LeftBorder { Val = BorderValues.None }, new RightBorder { Val = BorderValues.None }));
        if (!string.IsNullOrEmpty(shadingFill))
            cp.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = shadingFill });
        cp.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        var pp = new ParagraphProperties();
        pp.Append(new Justification { Val = align switch
        {
            Align.Right  => JustificationValues.Right,
            Align.Center => JustificationValues.Center,
            _            => JustificationValues.Left,
        } });

        var rp = new RunProperties();
        if (isLabel) rp.Append(new Bold());
        rp.Append(new FontSize { Val = "16" });

        return new TableCell(cp, new Paragraph(pp,
            new Run(rp, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static string Fmt(double v)
    {
        if (double.IsNaN(v) || double.IsInfinity(v)) return "—";
        if (v == 0) return "0";
        if (Math.Abs(v) >= 100) return v.ToString("F1");
        if (Math.Abs(v) >= 1)   return v.ToString("F2");
        return v.ToString("F3");
    }

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [EcotoxExporter] {msg}";
        try { File.AppendAllText(Path.Combine(AppPaths.LogsDir, "EcotoxExport.log"), line + Environment.NewLine); }
        catch { }
        System.Diagnostics.Debug.WriteLine(line);
    }
}
