using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 생태독성 (물벼룩 급성독성) 시험기록부 Word(.docx) 출력.
/// TestRecordBookWordExporter 와 동일한 시각언어 (가로선 표 / 8pt 본문 / 9pt 섹션 / renewus 워터마크).
/// </summary>
public static class EcotoxicityWordExporter
{
    private const string FontFamily = "맑은 고딕";

    public sealed record Record(
        string Date,
        string TestNo,
        string Species,
        string Toxicant,
        string SampleName,
        EcotoxicityService.EcotoxResult? TskResult,
        EcotoxicityService.EcotoxResult? ProbitResult,
        double[] Conc,
        int[] Org,
        int[] Mort,
        int CtrlOrg,
        int CtrlMort,
        double TestTemperature,
        double TestPH,
        double SampleTemperature,
        double SamplePH,
        double SampleDO,
        string Duration,
        string EcCalculationMethod,
        string Observations,
        string AnalystName);

    public static string Export(IList<Record> records)
    {
        var dir = Path.Combine(AppPaths.WritableDataRoot, "PrintCache");
        Directory.CreateDirectory(dir);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var path = Path.Combine(dir, $"생태독성_시험기록부_{stamp}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            ApplyDocDefaults(main);
            ApplyPageSetup(main, body);

            // ── 타이틀 ───────────────────────────────────────────────────────
            body.Append(BuildTitle());

            for (int idx = 0; idx < records.Count; idx++)
            {
                var rec = records[idx];

                // 시험 식별 (시험번호) — 다건일 때 구분
                if (records.Count > 1)
                    body.Append(BuildRecordHeader(idx + 1, rec));

                // 1) 시험기본정보
                body.Append(BuildSectionTitle("시험기본정보"));
                body.Append(BuildBasicInfoTable(rec));
                body.Append(SmallSpacer());

                // 2) 시험조건
                body.Append(BuildSectionTitle("시험조건 (ES 04704.1c 기준)"));
                body.Append(BuildConditionTable(rec));
                body.Append(SmallSpacer());

                // 3) 대조군 정보
                body.Append(BuildSectionTitle("대조군 정보"));
                body.Append(BuildControlTable(rec));
                body.Append(SmallSpacer());

                // 4) 농도별 시험 데이터
                body.Append(BuildSectionTitle("농도별 독성 시험 데이터"));
                body.Append(BuildConcentrationTable(rec));
                body.Append(SmallSpacer());

                // 5) 분석 결과 (TSK / Probit)
                body.Append(BuildSectionTitle("독성 분석 결과"));
                body.Append(BuildResultsTable(rec));
                body.Append(SmallSpacer());

                // 6) 용량-반응 곡선 PNG (TSK 결과가 있을 때만)
                if (rec.TskResult != null && rec.Conc != null && rec.Conc.Length > 0)
                {
                    try
                    {
                        var png = EcotoxicityChartGenerator.Generate(
                            rec.Conc, rec.Org, rec.Mort,
                            rec.TskResult.EC50, rec.TskResult.LowerCI, rec.TskResult.UpperCI,
                            rec.TskResult.TU, rec.TskResult.Method);
                        body.Append(BuildSectionTitle("용량-반응 곡선"));
                        body.Append(BuildChartParagraph(main, png));
                        body.Append(SmallSpacer());
                    }
                    catch { /* 차트 실패 시 본문만 */ }
                }

                // 7) 비고 / 분석자
                if (!string.IsNullOrWhiteSpace(rec.Observations))
                {
                    body.Append(BuildSectionTitle("시험 중 관찰사항"));
                    body.Append(MakeParagraph(rec.Observations, fontSize: 9));
                    body.Append(SmallSpacer());
                }

                // 다건일 때 페이지 분리 (마지막 제외)
                if (records.Count > 1 && idx < records.Count - 1)
                    body.Append(PageBreakParagraph());
            }

            // 8) 서명 — 마지막 페이지 하단
            body.Append(BuildSignatureTable());

            main.Document.Save();
        }
        return path;
    }

    // ─── 페이지 + 워터마크 ─────────────────────────────────────────────────
    private static void ApplyPageSetup(MainDocumentPart main, Body body)
    {
        var headerPart = main.AddNewPart<HeaderPart>();

        string? logoRelId = null;
        try
        {
            Stream? logoStream = null;
            try
            {
                var uri = new Uri("avares://ETA/Assets/icons/renewus_vertical_black.png");
                logoStream = Avalonia.Platform.AssetLoader.Open(uri);
            }
            catch
            {
                var logoPath = Path.Combine(AppPaths.RootPath, "Assets", "icons", "renewus_vertical_black.png");
                if (File.Exists(logoPath)) logoStream = File.OpenRead(logoPath);
            }
            if (logoStream != null)
            {
                using (logoStream)
                {
                    var imagePart = headerPart.AddImagePart(ImagePartType.Png);
                    imagePart.FeedData(logoStream);
                    logoRelId = headerPart.GetIdOfPart(imagePart);
                }
            }
        }
        catch { }

        string headerXml = !string.IsNullOrEmpty(logoRelId)
            ? BuildImageWatermarkHeaderXml(logoRelId!)
            : BuildTextWatermarkHeaderXml("ETA");

        using (var s = headerPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
        using (var sw = new System.IO.StreamWriter(s, new System.Text.UTF8Encoding(false)))
            sw.Write(headerXml);
        var headerRef = new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) };

        var sectionProps = new SectionProperties(
            headerRef,
            new PageSize { Width = 11906U, Height = 16838U },
            new PageMargin
            {
                Top = 720, Right = 720, Bottom = 720, Left = 720,
                Header = 360, Footer = 360, Gutter = 0,
            });
        body.Append(sectionProps);
    }

    private static string BuildImageWatermarkHeaderXml(string relId) =>
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
       xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
       xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
       xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""
       xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"">
  <w:p>
    <w:r>
      <w:rPr><w:noProof/></w:rPr>
      <w:drawing>
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
          <a:graphic>
            <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id=""1"" name=""renewus""/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed=""" + relId + @""">
                    <a:alphaModFix amt=""40000""/>
                  </a:blip>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""2197100"" cy=""1524000""/></a:xfrm>
                  <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>";

    private static string BuildTextWatermarkHeaderXml(string text) =>
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:v=""urn:schemas-microsoft-com:vml""
       xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <w:p>
    <w:r>
      <w:rPr><w:noProof/></w:rPr>
      <w:pict>
        <v:shape id=""WMText"" type=""#_x0000_t136""
                 style=""position:absolute;margin-left:0;margin-top:0;width:173pt;height:86pt;rotation:-45;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page""
                 fillcolor=""#808080"" stroked=""f"">
          <v:fill opacity="".4""/>
          <v:textpath style=""font-family:&quot;맑은 고딕&quot;;font-size:1pt;v-text-kern:t""
                      fitpath=""t"" string=""" + System.Net.WebUtility.HtmlEncode(text) + @"""/>
        </v:shape>
      </w:pict>
    </w:r>
  </w:p>
</w:hdr>";

    private static void ApplyDocDefaults(MainDocumentPart main)
    {
        var stylePart = main.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                        new FontSize { Val = "16" } // 8pt
                    )
                ),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { After = "0", Before = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                    )
                )
            )
        );
        stylePart.Styles = styles;
        stylePart.Styles.Save();
    }

    // ─── 타이틀 / 섹션 ────────────────────────────────────────────────────
    private static Paragraph BuildTitle() =>
        new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "60", Before = "120" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "36" }, // 18pt
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }
                ),
                new Text("물벼룩을 이용한 급성 독성 시험 기록부")
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "18" }, // 9pt
                    new Italic(),
                    new Color { Val = "555555" }
                ),
                new Break(),
                new Text("ES 04704.1c (수질오염공정시험기준)")
            )
        );

    private static Paragraph BuildRecordHeader(int seq, Record rec) =>
        new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "120", After = "60" }),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "20" }, // 10pt
                    new Color { Val = "1F3A5F" }
                ),
                new Text($"[{seq}] 시험번호: {rec.TestNo}    시험일자: {rec.Date}") { Space = SpaceProcessingModeValues.Preserve }
            )
        );

    private static Paragraph BuildSectionTitle(string title) =>
        new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "160", After = "60" }),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "18" }, // 9pt
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                    new Color { Val = "1F3A5F" }
                ),
                new Text("■ " + title)
            )
        );

    private static Paragraph SmallSpacer() =>
        new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "40", After = "40" }));

    private static Paragraph PageBreakParagraph() =>
        new Paragraph(new Run(new Break { Type = BreakValues.Page }));

    // ─── 표 빌더 ───────────────────────────────────────────────────────────
    private static Table BuildBasicInfoTable(Record r)
    {
        var t = NewBorderedTable();
        AddColumnGrid(t, 1700, 3500, 1700, 3500);
        t.Append(MakeKvRow(("시험일자", r.Date), ("시험번호", r.TestNo)));
        t.Append(MakeKvRow(("시험생물", string.IsNullOrWhiteSpace(r.Species) ? "Daphnia magna Straus" : r.Species),
                           ("시험기간", r.Duration)));
        t.Append(MakeKvRow(("시료명", r.SampleName), ("오염물질", r.Toxicant)));
        return t;
    }

    private static Table BuildConditionTable(Record r)
    {
        var t = NewBorderedTable();
        AddColumnGrid(t, 1700, 1700, 1700, 1700, 1700, 1900);
        // 시험온도 / 기준 / pH / 기준
        t.Append(MakeRow(
            ("시험온도(°C)", true), (FmtNum(r.TestTemperature), false),
            ("기준 20±2",  true), ("시험 pH",     true),
            (FmtNum(r.TestPH), false), ("기준 7.6~8.0", true)));
        t.Append(MakeRow(
            ("시료온도(°C)", true), (FmtNum(r.SampleTemperature), false),
            ("시료 pH",     true), (FmtNum(r.SamplePH), false),
            ("용존산소(mg/L)", true), (FmtNum(r.SampleDO), false)));
        t.Append(MakeKvRowWide("EC50 계산방법", string.IsNullOrWhiteSpace(r.EcCalculationMethod) ? "—" : r.EcCalculationMethod));
        if (!string.IsNullOrWhiteSpace(r.AnalystName))
            t.Append(MakeKvRowWide("분석자", r.AnalystName));
        return t;
    }

    private static Table BuildControlTable(Record r)
    {
        var t = NewBorderedTable();
        AddColumnGrid(t, 1700, 1700, 1700, 1700, 1700, 1900);
        double rate = r.CtrlOrg > 0 ? (r.CtrlMort * 100.0 / r.CtrlOrg) : 0;
        bool over = rate > 15;
        var row = new TableRow();
        row.Append(MakeCell("대조생물수", isLabel: true, alignCenter: true, width: 1700));
        row.Append(MakeCell(r.CtrlOrg.ToString(), alignRight: true, width: 1700));
        row.Append(MakeCell("대조 치사수", isLabel: true, alignCenter: true, width: 1700));
        row.Append(MakeCell(r.CtrlMort.ToString(), alignRight: true, width: 1700));
        row.Append(MakeCell("치사율(%)",  isLabel: true, alignCenter: true, width: 1700));
        // 치사율 15% 초과면 셀 강조 (옅은 빨강)
        var rateCell = MakeCell(rate.ToString("F1"), alignRight: true, width: 1900,
            shadingFill: over ? "FBE2E2" : null);
        row.Append(rateCell);
        t.Append(row);
        return t;
    }

    private static Table BuildConcentrationTable(Record r)
    {
        var t = NewBorderedTable();
        AddColumnGrid(t, 2000, 1400, 2000, 2000, 3000);
        var hdr = new TableRow();
        hdr.Append(MakeCell("농도",     isLabel: true, alignCenter: true, width: 2000));
        hdr.Append(MakeCell("단위",     isLabel: true, alignCenter: true, width: 1400));
        hdr.Append(MakeCell("생물수",    isLabel: true, alignCenter: true, width: 2000));
        hdr.Append(MakeCell("사망수",    isLabel: true, alignCenter: true, width: 2000));
        hdr.Append(MakeCell("치사율(%)", isLabel: true, alignCenter: true, width: 3000));
        t.Append(hdr);

        for (int i = 0; i < r.Conc.Length; i++)
        {
            var conc = r.Conc[i];
            var org = i < r.Org.Length ? r.Org[i] : 0;
            var mort = i < r.Mort.Length ? r.Mort[i] : 0;
            var rate = org > 0 ? (mort * 100.0 / org) : 0;

            var row = new TableRow();
            row.Append(MakeCell(FmtNum(conc), alignRight: true, width: 2000));
            row.Append(MakeCell("%",            alignCenter: true, width: 1400));
            row.Append(MakeCell(org.ToString(), alignRight: true, width: 2000));
            row.Append(MakeCell(mort.ToString(),alignRight: true, width: 2000));
            row.Append(MakeCell(rate.ToString("F1"), alignRight: true, width: 3000));
            t.Append(row);
        }
        return t;
    }

    private static Table BuildResultsTable(Record r)
    {
        var t = NewBorderedTable();
        AddColumnGrid(t, 2000, 2200, 2200, 2000, 2000);
        // 헤더
        var hdr = new TableRow();
        hdr.Append(MakeCell("분석법",      isLabel: true, alignCenter: true, width: 2000));
        hdr.Append(MakeCell("EC50",        isLabel: true, alignCenter: true, width: 2200));
        hdr.Append(MakeCell("TU",          isLabel: true, alignCenter: true, width: 2200));
        hdr.Append(MakeCell("95% CI 하한", isLabel: true, alignCenter: true, width: 2000));
        hdr.Append(MakeCell("95% CI 상한", isLabel: true, alignCenter: true, width: 2000));
        t.Append(hdr);

        if (r.TskResult != null)
        {
            var row = new TableRow();
            row.Append(MakeCell("TSK", isLabel: true, alignCenter: true, width: 2000));
            row.Append(MakeCell(FmtNum(r.TskResult.EC50),   alignRight: true, width: 2200));
            row.Append(MakeCell(FmtNum(r.TskResult.TU),     alignRight: true, width: 2200));
            row.Append(MakeCell(FmtNum(r.TskResult.LowerCI),alignRight: true, width: 2000));
            row.Append(MakeCell(FmtNum(r.TskResult.UpperCI),alignRight: true, width: 2000));
            t.Append(row);

            if (r.TskResult.TrimPercent >= 0)
                t.Append(MakeKvRowWide("Trim %", r.TskResult.TrimPercent.ToString("F1")));
            if (!string.IsNullOrEmpty(r.TskResult.Warning))
                t.Append(MakeKvRowWide("⚠ TSK 주의", r.TskResult.Warning, valueShading: "FFF6D6"));

            // TU 기준 독성 분류 한 줄
            string cls = r.TskResult.TU > 16 ? "높은 독성"
                       : r.TskResult.TU > 8  ? "중간 독성"
                       : r.TskResult.TU > 1  ? "낮은 독성"
                                              : "매우 낮은 독성";
            string clsShade = r.TskResult.TU > 16 ? "FBE2E2"
                            : r.TskResult.TU > 8  ? "FFF6D6"
                                                  : null!;
            t.Append(MakeKvRowWide("독성 분류 (TU 기준)", cls, valueShading: clsShade));
        }

        if (r.ProbitResult != null)
        {
            var row = new TableRow();
            row.Append(MakeCell("Probit", isLabel: true, alignCenter: true, width: 2000));
            row.Append(MakeCell(FmtNum(r.ProbitResult.EC50),   alignRight: true, width: 2200));
            row.Append(MakeCell(FmtNum(r.ProbitResult.TU),     alignRight: true, width: 2200));
            row.Append(MakeCell(FmtNum(r.ProbitResult.LowerCI),alignRight: true, width: 2000));
            row.Append(MakeCell(FmtNum(r.ProbitResult.UpperCI),alignRight: true, width: 2000));
            t.Append(row);

            if (!string.IsNullOrEmpty(r.ProbitResult.Warning))
                t.Append(MakeKvRowWide("⚠ Probit 주의", r.ProbitResult.Warning, valueShading: "FFF6D6"));
        }

        if (r.TskResult == null && r.ProbitResult == null)
            t.Append(MakeKvRowWide("결과", "분석 결과 없음"));
        return t;
    }

    private static Paragraph BuildChartParagraph(MainDocumentPart main, byte[] png)
    {
        var imgPart = main.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(png)) imgPart.FeedData(ms);
        string relId = main.GetIdOfPart(imgPart);

        // 640x400 → EMU (1 px ≈ 9525 EMU). 페이지 폭 맞춤(약 425pt = 5398313 EMU 폭, 비율 유지 → 높이 3373945)
        // 이미지 비율 640:400 = 1.6:1 → 폭 425pt 이면 height 265pt
        long cx = 5400000; // ≈ 425pt = 6"
        long cy = (long)(cx * 400.0 / 640.0); // 비율 유지

        string xml =
$@"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
        xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
        xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
        xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
        xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
  <w:pPr><w:jc w:val=""center""/></w:pPr>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <w:drawing>
      <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
        <wp:extent cx=""{cx}"" cy=""{cy}""/>
        <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
        <wp:docPr id=""100"" name=""Chart""/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=""1""/></wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
            <pic:pic>
              <pic:nvPicPr><pic:cNvPr id=""100"" name=""Chart""/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed=""{relId}""/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{cx}"" cy=""{cy}""/></a:xfrm>
                <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>";
        return new Paragraph(new DocumentFormat.OpenXml.OpenXmlUnknownElement(xml));
    }

    // ─── 서명 표 (페이지 하단 anchor) ───────────────────────────────────────
    private static Table BuildSignatureTable()
    {
        var t = NewBorderedTable();
        var tp = t.GetFirstChild<TableProperties>();
        if (tp != null)
            tp.Append(new TablePositionProperties
            {
                LeftFromText = 0, RightFromText = 0,
                TopFromText = 0, BottomFromText = 0,
                VerticalAnchor = VerticalAnchorValues.Margin,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            });

        AddColumnGrid(t, 3467, 3467, 3466);
        var hdr = new TableRow(new TableRowProperties(new TableRowHeight { Val = 280U, HeightType = HeightRuleValues.AtLeast }));
        hdr.Append(MakeCell("작성자", width: 3467, isLabel: true, alignCenter: true));
        hdr.Append(MakeCell("검토자", width: 3467, isLabel: true, alignCenter: true));
        hdr.Append(MakeCell("승인자", width: 3466, isLabel: true, alignCenter: true));
        t.Append(hdr);

        var sig = new TableRow(new TableRowProperties(new TableRowHeight { Val = 500U, HeightType = HeightRuleValues.AtLeast }));
        sig.Append(MakeCell("", width: 3467, alignCenter: true));
        sig.Append(MakeCell("", width: 3467, alignCenter: true));
        sig.Append(MakeCell("", width: 3466, alignCenter: true));
        t.Append(sig);
        return t;
    }

    // ─── 헬퍼 ──────────────────────────────────────────────────────────────
    private static TableRow MakeRow(params (string text, bool isLabel)[] cells)
    {
        var row = new TableRow();
        foreach (var (text, isLabel) in cells)
        {
            if (isLabel) row.Append(MakeCell(text, isLabel: true, alignCenter: true));
            else         row.Append(MakeCell(text, alignRight: true));
        }
        return row;
    }

    private static TableRow MakeKvRow((string k, string v) a, (string k, string v) b)
    {
        var row = new TableRow();
        row.Append(MakeCell(a.k, isLabel: true, alignCenter: true, width: 1700));
        row.Append(MakeCell(string.IsNullOrWhiteSpace(a.v) ? "—" : a.v, alignLeft: true, width: 3500));
        row.Append(MakeCell(b.k, isLabel: true, alignCenter: true, width: 1700));
        row.Append(MakeCell(string.IsNullOrWhiteSpace(b.v) ? "—" : b.v, alignLeft: true, width: 3500));
        return row;
    }

    private static TableRow MakeKvRowWide(string k, string v, string? valueShading = null)
    {
        var row = new TableRow();
        row.Append(MakeCell(k, isLabel: true, alignCenter: true, width: 2200));
        row.Append(MakeCell(string.IsNullOrWhiteSpace(v) ? "—" : v, alignLeft: true, width: 8200, shadingFill: valueShading));
        return row;
    }

    private static string FmtNum(double v)
    {
        if (double.IsNaN(v) || double.IsInfinity(v)) return "—";
        if (v == 0) return "0";
        if (Math.Abs(v) >= 100) return v.ToString("F1");
        if (Math.Abs(v) >= 1)   return v.ToString("F2");
        return v.ToString("F3");
    }

    private static TableCell MakeCell(string text, int? width = null,
        bool isLabel = false, bool alignRight = false, bool alignCenter = false,
        bool alignLeft = false, string? shadingFill = null)
    {
        var cellProps = new TableCellProperties();
        if (width.HasValue)
            cellProps.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.Value.ToString() });
        cellProps.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        if (!string.IsNullOrEmpty(shadingFill))
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = shadingFill });
        cellProps.Append(new TableCellBorders(
            new LeftBorder  { Val = BorderValues.None },
            new RightBorder { Val = BorderValues.None }));

        var paraProps = new ParagraphProperties();
        if (alignRight)  paraProps.Append(new Justification { Val = JustificationValues.Right });
        else if (alignLeft)   paraProps.Append(new Justification { Val = JustificationValues.Left });
        else if (alignCenter || isLabel) paraProps.Append(new Justification { Val = JustificationValues.Center });

        var runProps = new RunProperties();
        if (isLabel) runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = "16" });

        var run = new Run(runProps, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        return new TableCell(cellProps, new Paragraph(paraProps, run));
    }

    private static Paragraph MakeParagraph(string text, int fontSize = 9) =>
        new Paragraph(
            new Run(
                new RunProperties(
                    new FontSize { Val = (fontSize * 2).ToString() },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }
                ),
                new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }
            )
        );

    private static Table NewBorderedTable() =>
        new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new LeftBorder   { Val = BorderValues.None },
                    new RightBorder  { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "BBBBBB" },
                    new InsideVerticalBorder   { Val = BorderValues.None }
                ),
                new TableLayout { Type = TableLayoutValues.Fixed }
            )
        );

    private static void AddColumnGrid(Table table, params int[] widths)
    {
        var grid = new TableGrid();
        foreach (var w in widths)
            grid.Append(new GridColumn { Width = w.ToString() });
        table.Append(grid);
    }
}
