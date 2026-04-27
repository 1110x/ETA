using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ETA.Services.Common;
using ETA.Views.Pages.PAGE1;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 시험기록부 Word(.docx) 출력 — Excel 출력과 같은 섹션 구조
/// (타이틀 → 메타 → 검정곡선 → 시료/QC → 비고 → 서명).
/// 프로토타입 단계 — TOC(NPOC/TCIC) 검증 우선.
/// </summary>
public static class TestRecordBookWordExporter
{
    private const string FontFamily = "맑은 고딕";

    public static string Export(TestRecordBookParsedView.Model m)
    {
        var dir = Path.Combine(AppPaths.WritableDataRoot, "PrintCache");
        Directory.CreateDirectory(dir);
        var stamp  = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var safeNm = string.Concat((m.TargetTable ?? "시험기록부")
            .Split(Path.GetInvalidFileNameChars()));
        var path = Path.Combine(dir, $"{safeNm}_{stamp}.docx");

        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            ApplyDocDefaults(main);
            ApplyPageSetup(main, body, watermarkText: "ETA");

            // 0) 타이틀
            body.Append(BuildTitle(m));

            // 1) 기록부정보 — 문서번호 / 분석방법 / 결과표시
            body.Append(BuildSectionTitle("기록부정보"));
            body.Append(BuildMetaTable(m));
            body.Append(SmallSpacer());

            // 1-1) 분석조건 (설정→분석조건) — 비어있으면 "해당없음"
            body.Append(BuildSectionTitle("분석조건"));
            body.Append(BuildAnalysisConditionsTable(m));
            body.Append(SmallSpacer());

            // 2) 검정곡선 (TCIC 면 TC + IC 두 세트). 기울기·R² 는 표 아래 별도 라인.
            int decimals = ResolveDecimals(m);
            if (m.IsTcic)
            {
                if (m.TcStandardKeys.Count > 0)
                {
                    body.Append(BuildSectionTitle("검정곡선 (TC)"));
                    body.Append(BuildCalibrationTable(m.TcStandardKeys, m.TcStandardConc, m.TcStandardAbs, decimals));
                    body.Append(BuildCalibrationFooter(m.TcSlopeText, m.TcR2Text));
                    body.Append(SmallSpacer());
                }
                if (m.IcStandardKeys.Count > 0)
                {
                    body.Append(BuildSectionTitle("검정곡선 (IC)"));
                    body.Append(BuildCalibrationTable(m.IcStandardKeys, m.IcStandardConc, m.IcStandardAbs, decimals));
                    body.Append(BuildCalibrationFooter(m.IcSlopeText, m.IcR2Text));
                    body.Append(SmallSpacer());
                }
            }
            else if (m.StandardKeys.Count > 0)
            {
                body.Append(BuildSectionTitle("검정곡선"));
                body.Append(BuildCalibrationTable(m.StandardKeys, m.StandardConc, m.StandardAbs, decimals));
                body.Append(BuildCalibrationFooter(m.SlopeText, m.R2Text));
                body.Append(SmallSpacer());
            }

            // 3) 검정곡선의 보증 (QC) — 비어있어도 "자료없음" 표시
            body.Append(BuildSectionTitle("검정곡선의 보증"));
            if (m.QcRows.Count > 0)
                body.Append(BuildSampleTable(m.SampleHeaders, m.QcRows));
            else
                body.Append(BuildEmptyPlaceholder("자료없음"));
            body.Append(SmallSpacer());

            // 4) 시료분석결과
            if (m.SampleRows.Count > 0)
            {
                body.Append(BuildSectionTitle("시료분석결과"));
                body.Append(BuildSampleTable(m.SampleHeaders, m.SampleRows));
                body.Append(SmallSpacer());
            }

            // 5) 비고 / 메모
            if (!string.IsNullOrWhiteSpace(m.Memo))
            {
                body.Append(BuildSectionTitle("비고"));
                body.Append(MakeParagraph(m.Memo, fontSize: 10));
                body.Append(SmallSpacer());
            }

            // 6) 서명
            body.Append(BuildSignatureTable());

            // SectionProperties 는 body 의 마지막 자식이어야 함
            if (_pendingSectionProps != null)
            {
                body.Append(_pendingSectionProps);
                _pendingSectionProps = null;
            }

            main.Document.Save();
        }
        return path;
    }

    // ─── 페이지 + 헤더 워터마크 ──────────────────────────────────────────────
    private static void ApplyPageSetup(MainDocumentPart main, Body body, string watermarkText)
    {
        var headerPart = main.AddNewPart<HeaderPart>();

        // 회사 로고(renewus PNG)를 헤더에 임베드 — 이미지 워터마크.
        // Avalonia 임베디드 리소스(avares://) 우선, 디스크 파일 폴백.
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
        catch { /* 로고 없으면 텍스트 워터마크로 폴백 */ }

        string esc = System.Net.WebUtility.HtmlEncode(watermarkText ?? "ETA");
        string headerXml = !string.IsNullOrEmpty(logoRelId)
            ? BuildImageWatermarkHeaderXml(logoRelId)
            : BuildTextWatermarkHeaderXml(esc);

        using (var s = headerPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
        using (var sw = new System.IO.StreamWriter(s, new System.Text.UTF8Encoding(false)))
            sw.Write(headerXml);
        var headerRef = new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) };

        // SectionProperties 는 body 의 LAST child 여야 함 (스키마).
        // 만들어 두기만 하고, Export() 끝에서 body.Append.
        _pendingSectionProps = new SectionProperties(
            headerRef,
            new PageSize { Width = 11906U, Height = 16838U },     // A4 portrait
            new PageMargin
            {
                Top    = 720,    // 0.5"
                Right  = 720,
                Bottom = 720,
                Left   = 720,
                Header = 360,
                Footer = 360,
                Gutter = 0,
            });
    }

    [ThreadStatic] private static SectionProperties? _pendingSectionProps;

    /// <summary>로고 PNG 임베드 워터마크 — behindDoc=1, alphaModFix 로 40% 불투명.</summary>
    private static string BuildImageWatermarkHeaderXml(string relId)
    {
        // renewus 로고 비율 771:536 ≈ 1.44 (가로:세로). width 173pt 면 height ≈ 120pt
        // EMU: 1pt = 12700, 173pt = 2197100, 120pt = 1524000
        return
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
    }

    /// <summary>로고 파일이 없을 때 폴백 — WordArt 텍스트 워터마크.</summary>
    private static string BuildTextWatermarkHeaderXml(string esc)
    {
        return
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
       xmlns:v=""urn:schemas-microsoft-com:vml""
       xmlns:o=""urn:schemas-microsoft-com:office:office""
       xmlns:w10=""urn:schemas-microsoft-com:office:word""
       xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""
       mc:Ignorable=""w14 wp14 v"">
  <w:p>
    <w:r>
      <w:rPr><w:noProof/></w:rPr>
      <w:pict>
        <v:shape id=""WMText"" o:spid=""_x0000_s1026"" type=""#_x0000_t136""
                 style=""position:absolute;margin-left:0;margin-top:0;width:173pt;height:86pt;rotation:-45;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page""
                 fillcolor=""#808080"" stroked=""f"">
          <v:fill opacity="".4""/>
          <v:textpath style=""font-family:&quot;맑은 고딕&quot;;font-size:1pt;v-text-kern:t""
                      fitpath=""t"" string=""" + esc + @"""/>
        </v:shape>
      </w:pict>
    </w:r>
  </w:p>
</w:hdr>";
    }

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

    // ─── 타이틀 ──────────────────────────────────────────────────────────────
    private static Paragraph BuildTitle(TestRecordBookParsedView.Model m)
    {
        var pretty = TestRecordBookViewerService.PrettyName(m.TargetTable ?? "");
        var title = !string.IsNullOrWhiteSpace(pretty)
            ? $"{pretty} 시험기록부"
            : "시험기록부";

        var p = new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "120", Before = "120" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "36" }, // 18pt
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }
                ),
                new Text(title) { Space = SpaceProcessingModeValues.Preserve }
            )
        );

        if (!string.IsNullOrWhiteSpace(m.AnalysisDate))
        {
            var sub = new Run(
                new RunProperties(
                    new FontSize { Val = "20" },
                    new Color { Val = "555555" }
                ),
                new Break(),
                new Text($"분석일: {m.AnalysisDate}") { Space = SpaceProcessingModeValues.Preserve }
            );
            p.Append(sub);
        }
        return p;
    }

    // ─── 섹션 타이틀 — 9pt ────────────────────────────────────────────────
    private static Paragraph BuildSectionTitle(string title)
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "160", After = "60" }
            ),
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
    }

    private static Paragraph SmallSpacer() =>
        new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "60" }));

    // ─── 메타 표 ─────────────────────────────────────────────────────────────
    private static Table BuildMetaTable(TestRecordBookParsedView.Model m)
    {
        var pretty = TestRecordBookViewerService.PrettyName(m.TargetTable ?? "");
        string docNo = !string.IsNullOrWhiteSpace(pretty) && !string.IsNullOrWhiteSpace(m.AnalysisDate)
            ? $"{pretty} {m.AnalysisDate}"
            : (m.AnalysisDate ?? "");
        var methodLine = string.Join("   ", new[] { m.ES, m.Method ?? m.AnalysisMethod, m.Instrument }
                            .Where(s => !string.IsNullOrWhiteSpace(s)));
        var formulaText = !string.IsNullOrWhiteSpace(m.ResultFormula) ? m.ResultFormula : "";

        var table = NewBorderedTable();
        AddColumnGrid(table, 1600, 8800);  // 라벨 1600 + 값 8800 = 10400 dxa (≈ 7.2")
        table.Append(MakeMetaRow("문서번호",  docNo));
        table.Append(MakeMetaRow("분석방법",  methodLine));
        table.Append(MakeMetaRow("결과표시",  formulaText));
        return table;
    }

    private static TableRow MakeMetaRow(string label, string value)
    {
        var row = new TableRow();
        row.Append(MakeCell(label, width: 1600, isLabel: true, alignCenter: true));
        row.Append(MakeCell(value, width: 8800, alignLeft: true));
        return row;
    }

    // ─── 검정곡선 표 ─────────────────────────────────────────────────────────
    private static Table BuildCalibrationTable(
        List<int> keys, List<string> conc, List<string> abs, int decimals)
    {
        var table = NewBorderedTable();

        // 컬럼 그리드: STANDARD 라벨(1400) + N개 검정점 균등 분배
        const int totalWidth = 10400;
        const int labelW = 1400;
        int dataCols = Math.Max(keys.Count, 1);
        int dataW = (totalWidth - labelW) / dataCols;
        var widths = new List<int> { labelW };
        for (int i = 0; i < dataCols; i++) widths.Add(dataW);
        AddColumnGrid(table, widths.ToArray());

        // 헤더 행: "STANDARD" | k1 | k2 | ...
        var hdr = new TableRow();
        hdr.Append(MakeCell("STANDARD", isLabel: true, alignCenter: true));
        foreach (var k in keys)
            hdr.Append(MakeCell(k.ToString(), isLabel: true, alignCenter: true));
        table.Append(hdr);

        // 농도 행 — 사용자 지정 소수점 자릿수로 포매팅
        var concRow = new TableRow();
        concRow.Append(MakeCell("농도(mg/L)", isLabel: true, alignCenter: true));
        for (int i = 0; i < keys.Count; i++)
        {
            string v = i < conc.Count ? FormatNumber(conc[i], decimals) : "";
            concRow.Append(MakeCell(v, alignRight: true));
        }
        table.Append(concRow);

        // 흡광도 행
        var absRow = new TableRow();
        absRow.Append(MakeCell("흡광도", isLabel: true, alignCenter: true));
        for (int i = 0; i < keys.Count; i++)
            absRow.Append(MakeCell(i < abs.Count ? abs[i] : "", alignRight: true));
        table.Append(absRow);

        return table;
    }

    /// <summary>검정곡선 아래 한 줄로 기울기/R² 표시. slopeText/r2Text 가 이미 "a=..", "R²=.." 라벨을
    /// 포함하는 경우 중복 표기 제거.</summary>
    private static Paragraph BuildCalibrationFooter(string slopeText, string r2Text)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(slopeText))
        {
            var s = slopeText.Trim();
            // 입력이 "a=0.000679 b=-0.575" 처럼 이미 라벨 가짐 — 그대로 사용
            parts.Add(s.StartsWith("a", StringComparison.OrdinalIgnoreCase) ? $"기울기 {s}" : $"기울기 a = {s}");
        }
        if (!string.IsNullOrWhiteSpace(r2Text))
        {
            var s = r2Text.Trim();
            parts.Add(s.StartsWith("R", StringComparison.OrdinalIgnoreCase) ? s : $"R² = {s}");
        }
        var line = string.Join("    ", parts);
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "60", After = "0" },
                new Justification { Val = JustificationValues.Right }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "16" }, // 8pt
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                    new Color { Val = "555555" }
                ),
                new Text(line) { Space = SpaceProcessingModeValues.Preserve }
            )
        );
    }

    private static int ResolveDecimals(TestRecordBookParsedView.Model m)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(m.Analyte))
                return TestRecordBookViewerService.GetDecimalPlaces(m.Analyte);
        }
        catch { }
        return 3; // 기본 3자리
    }

    private static string FormatNumber(string raw, int decimals)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        if (double.TryParse(raw, out var d))
            return d.ToString("F" + decimals);
        return raw;
    }

    // ─── 시료/QC 표 (계산식 alt-row 포함) ────────────────────────────────────
    private static Table BuildSampleTable(List<string> headers, List<List<string>> rows)
    {
        // 계산식 컬럼은 alt-row 로 빠짐 (Excel/Show2 와 동일 시각언어)
        int formulaIdx = headers.IndexOf("계산식");
        var displayHeaders = headers.Where(h => h != "계산식").ToList();

        // 표시용 행 (계산식 제거)
        var displayRows = rows.ConvertAll(r =>
        {
            if (formulaIdx < 0 || formulaIdx >= r.Count) return r;
            var copy = new List<string>(r); copy.RemoveAt(formulaIdx); return copy;
        });

        // 컬럼 너비 — 헤더/데이터의 실제 글자 폭 기반 자동 분배 (스키마 가변)
        int[] widths = AutoColumnWidths(displayHeaders, displayRows);

        // 컬럼별 정렬 자동 판정 (헤더 + 데이터 기반)
        var aligns = new (bool right, bool center)[displayHeaders.Count];
        for (int i = 0; i < displayHeaders.Count; i++)
            aligns[i] = ResolveAlignment(displayHeaders[i], displayRows, i);

        var table = NewBorderedTable();
        AddColumnGrid(table, widths);

        // 헤더
        var hdr = new TableRow();
        for (int i = 0; i < displayHeaders.Count; i++)
            hdr.Append(MakeCell(displayHeaders[i], width: widths[i], isLabel: true, alignCenter: true));
        table.Append(hdr);

        // 데이터 행 + 계산식 alt-row
        foreach (var r in rows)
        {
            var dataRow = new TableRow();
            int colIdx = 0;
            for (int srcCol = 0; srcCol < headers.Count; srcCol++)
            {
                if (srcCol == formulaIdx) continue;
                var val = srcCol < r.Count ? (r[srcCol] ?? "") : "";
                var (right, center) = aligns[colIdx];
                // 숫자 컬럼: decimal tab 정렬로 소수점 일치 + 줄바꿈 차단
                bool useDec = right && IsDecimalLike(val);
                dataRow.Append(MakeCell(val, width: widths[colIdx],
                    alignRight: right && !useDec,
                    alignCenter: center,
                    alignLeft: !right && !center,
                    noWrap: right,
                    decimalAlign: useDec));
                colIdx++;
            }
            table.Append(dataRow);

            // 계산식 alt-row (모든 컬럼 병합) — 라이트 회색 배경, 이탤릭, 8pt
            string formula = (formulaIdx >= 0 && formulaIdx < r.Count) ? (r[formulaIdx] ?? "") : "";
            if (!string.IsNullOrWhiteSpace(formula))
            {
                var altRow = new TableRow();
                var mergedCell = new TableCell(
                    new TableCellProperties(
                        new GridSpan { Val = displayHeaders.Count },
                        // 옅은 베이지 — 시료별 식 alt-row 시각 구분 (Excel 출력과 동일)
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FBFAF5" },
                        new TableCellBorders(
                            new TopBorder    { Val = BorderValues.None },
                            new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "DDDDDD" },
                            new LeftBorder   { Val = BorderValues.None },
                            new RightBorder  { Val = BorderValues.None }
                        )
                    ),
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Left },
                            new Indentation { Left = "200" }
                        ),
                        new Run(
                            new RunProperties(
                                new Italic(),
                                new FontSize { Val = "16" }, // 8pt
                                new Color { Val = "555555" }
                            ),
                            new Text("↳  " + formula) { Space = SpaceProcessingModeValues.Preserve }
                        )
                    )
                );
                altRow.Append(mergedCell);
                table.Append(altRow);
            }
        }
        return table;
    }

    // ─── 데이터 기반 컬럼 너비 자동 산정 ─────────────────────────────────────
    // 스키마마다 다른 시험기록부에 동일하게 적용 — 헤더/데이터의 실제 글자 길이를
    // 측정해 비율 분배. 한글/CJK 는 ASCII 의 약 1.6 배 폭으로 가산.
    private const int PageContentWidth = 10400;  // dxa, A4 8.27" - 좌우 0.5"*2 마진 ≈ 7.27"
    private const int MinColWidth      = 700;    // 최소 컬럼 폭 (너무 좁아서 잘리지 않도록)
    private const int MaxColWidth      = 4500;   // 한 컬럼 독식 방지
    private const int CellPaddingDxa   = 240;    // 셀 좌우 inset 보정
    private const int AsciiCharDxa     = 110;    // 8pt ASCII 한 글자 ≈ 110 dxa
    private const int CjkCharDxa       = 175;    // 8pt 한글/CJK 한 글자 ≈ 175 dxa

    private static int[] AutoColumnWidths(List<string> headers, List<List<string>> rows)
    {
        int n = headers.Count;
        if (n == 0) return Array.Empty<int>();

        // 1) 헤더 + 데이터 셀에서 가장 넓은 글자 폭 추정
        var measured = new int[n];
        for (int c = 0; c < n; c++)
        {
            int w = MeasureCharWidth(headers[c]);
            foreach (var r in rows)
                if (c < r.Count) w = Math.Max(w, MeasureCharWidth(r[c] ?? ""));
            measured[c] = Math.Clamp(w + CellPaddingDxa, MinColWidth, MaxColWidth);
        }

        // 2) 합계가 페이지 폭에 맞게 비례 축소/확대
        int sum = measured.Sum();
        var widths = new int[n];
        if (sum == PageContentWidth)
            return measured;
        if (sum == 0)
        {
            for (int i = 0; i < n; i++) widths[i] = PageContentWidth / n;
            return widths;
        }
        double ratio = (double)PageContentWidth / sum;
        int used = 0;
        for (int i = 0; i < n - 1; i++)
        {
            widths[i] = Math.Max(MinColWidth, (int)Math.Round(measured[i] * ratio));
            used += widths[i];
        }
        widths[n - 1] = PageContentWidth - used;
        if (widths[n - 1] < MinColWidth) widths[n - 1] = MinColWidth;
        return widths;
    }

    private static int MeasureCharWidth(string s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        // 줄바꿈 포함 셀(매칭된 시료명 등) — 가장 긴 라인 기준으로 측정
        if (s.Contains('\n') || s.Contains('\r'))
        {
            var lines = s.Replace("\r\n", "\n").Split('\n');
            int maxW = 0;
            foreach (var line in lines)
                maxW = Math.Max(maxW, MeasureCharWidth(line));
            return maxW;
        }
        int w = 0;
        foreach (var ch in s)
        {
            // 한글/한자/가나 = CJK 너비, 그 외 ASCII 너비
            if ((ch >= 0xAC00 && ch <= 0xD7A3)         // 한글 음절
             || (ch >= 0x4E00 && ch <= 0x9FFF)         // CJK 통합 한자
             || (ch >= 0x3040 && ch <= 0x30FF)         // 가나
             || (ch >= 0xFF00 && ch <= 0xFFEF))        // 전각
                w += CjkCharDxa;
            else
                w += AsciiCharDxa;
        }
        return w;
    }

    /// <summary>"5.32" 처럼 소수점/숫자로만 구성됐는지 — decimal tab 적용 가능 여부.</summary>
    private static bool IsDecimalLike(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        return double.TryParse(s.Trim(), out _);
    }

    /// <summary>헤더/값으로 정렬 자동 판정. 값이 숫자/논숫자 비율로 결정.</summary>
    private static (bool right, bool center) ResolveAlignment(string header, List<List<string>> rows, int colIdx)
    {
        // 라벨성 헤더(SN/시료번호/시료구분/비고)는 가운데
        if (header is "SN" or "시료번호" or "시료구분" or "비고") return (false, true);
        // 시료명/계산식은 좌측
        if (header is "시료명" or "계산식") return (false, false);
        // 데이터 기준 — 셀 50% 이상이 숫자 파싱되면 우측, 아니면 좌측
        int total = 0, numeric = 0;
        foreach (var r in rows)
        {
            if (colIdx >= r.Count) continue;
            var v = r[colIdx]?.Trim() ?? "";
            if (v.Length == 0 || v == "—" || v == "ND") continue;
            total++;
            if (double.TryParse(v, out _)) numeric++;
        }
        if (total == 0) return (false, false);
        return (numeric * 2 >= total, false);
    }

    /// <summary>분석조건 Key/Value 표 — 비어있으면 "해당없음" 한 줄.</summary>
    private static Table BuildAnalysisConditionsTable(TestRecordBookParsedView.Model m)
    {
        var table = NewBorderedTable();
        AddColumnGrid(table, 2200, 8200);
        if (m.AnalysisConditions.Count == 0)
        {
            var row = new TableRow();
            row.Append(MakeCell("해당없음", width: 10400, alignCenter: true));
            // 단일 셀로 합치기 위해 GridSpan 적용
            var cell = (TableCell)row.LastChild!;
            var props = cell.GetFirstChild<TableCellProperties>() ?? new TableCellProperties();
            props.Append(new GridSpan { Val = 2 });
            if (cell.GetFirstChild<TableCellProperties>() == null)
                cell.InsertAt(props, 0);
            table.Append(row);
            return table;
        }
        foreach (var (k, v) in m.AnalysisConditions)
        {
            var row = new TableRow();
            row.Append(MakeCell(k, width: 2200, isLabel: true, alignCenter: true));
            row.Append(MakeCell(string.IsNullOrWhiteSpace(v) ? "—" : v, width: 8200, alignLeft: true));
            table.Append(row);
        }
        return table;
    }

    /// <summary>비어있는 섹션용 placeholder (얇은 한 줄).</summary>
    private static Table BuildEmptyPlaceholder(string text)
    {
        var table = NewBorderedTable();
        AddColumnGrid(table, 10400);
        var row = new TableRow();
        row.Append(MakeCell(text, width: 10400, alignCenter: true));
        table.Append(row);
        return table;
    }

    // ─── 서명 표 (페이지 하단 anchor + 서명칸 높이 절반) ────────────────────
    private static Table BuildSignatureTable()
    {
        var table = NewBorderedTable();

        // 스키마: tblpPr 는 tblBorders 보다 앞에 와야 함 → tblW 다음에 삽입
        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null)
        {
            var pos = new TablePositionProperties
            {
                LeftFromText  = 0, RightFromText = 0,
                TopFromText   = 0, BottomFromText = 0,
                VerticalAnchor   = VerticalAnchorValues.Margin,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            };
            // tblpPr 는 tblPr 의 가장 앞 (스키마: tblpPr > tblW > tblBorders ...)
            tp.InsertAt(pos, 0);
        }

        AddColumnGrid(table, 3467, 3467, 3466);
        var hdr = new TableRow(new TableRowProperties(new TableRowHeight { Val = 280U, HeightType = HeightRuleValues.AtLeast }));
        hdr.Append(MakeCell("작성자", width: 3467, isLabel: true, alignCenter: true));
        hdr.Append(MakeCell("검토자", width: 3467, isLabel: true, alignCenter: true));
        hdr.Append(MakeCell("승인자", width: 3466, isLabel: true, alignCenter: true));
        table.Append(hdr);

        // 서명 칸 — 이전 1000 의 절반 (500 twips ≈ 0.35")
        var sig = new TableRow(new TableRowProperties(new TableRowHeight { Val = 500U, HeightType = HeightRuleValues.AtLeast }));
        sig.Append(MakeCell("", width: 3467, alignCenter: true));
        sig.Append(MakeCell("", width: 3467, alignCenter: true));
        sig.Append(MakeCell("", width: 3466, alignCenter: true));
        table.Append(sig);
        return table;
    }

    // ─── 공통 헬퍼 ───────────────────────────────────────────────────────────
    /// <summary>가로선만 (위/아래/내부 가로) — 세로선 없음, 정렬에 의존.</summary>
    private static Table NewBorderedTable()
    {
        return new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100%
                // 스키마 자식 순서: top → left → bottom → right → insideH → insideV
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new LeftBorder   { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.Single, Size = 6U, Color = "555555" },
                    new RightBorder  { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "BBBBBB" },
                    new InsideVerticalBorder   { Val = BorderValues.None }
                ),
                new TableLayout { Type = TableLayoutValues.Fixed }
            )
        );
    }

    /// <summary>고정 너비 컬럼 그리드 추가 — Fixed layout 에 필수.</summary>
    private static void AddColumnGrid(Table table, params int[] widthsDxa)
    {
        var grid = new TableGrid();
        foreach (var w in widthsDxa) grid.Append(new GridColumn { Width = w.ToString() });
        // TableProperties 다음 위치에 삽입
        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null) tp.InsertAfterSelf(grid);
        else table.InsertAt(grid, 0);
    }

    private static TableCell MakeCell(string text, int? width = null,
        bool isLabel = false, bool alignRight = false, bool alignCenter = false,
        bool alignLeft = false, bool noWrap = false, bool decimalAlign = false)
    {
        // 스키마 자식 순서: tcW → tcBorders → noWrap → vAlign
        var cellProps = new TableCellProperties();
        if (width.HasValue)
            cellProps.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.Value.ToString() });
        // 셀 좌/우 테두리 명시 제거 — 표 자체의 수평선만 보이도록
        cellProps.Append(new TableCellBorders(
            new LeftBorder  { Val = BorderValues.None },
            new RightBorder { Val = BorderValues.None }));
        if (noWrap) cellProps.Append(new NoWrap());
        cellProps.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        var paraProps = new ParagraphProperties();

        bool useDecimalTab = decimalAlign && width.HasValue && !string.IsNullOrWhiteSpace(text);
        if (useDecimalTab)
        {
            // 셀 내부 소수점 정렬 — tab 위치는 셀 폭에서 우측 패딩(약 120dxa) 뺀 지점
            int tabPos = Math.Max(200, width!.Value - 140);
            paraProps.Append(new Tabs(new TabStop { Val = TabStopValues.Decimal, Position = tabPos }));
        }
        else if (alignRight)
            paraProps.Append(new Justification { Val = JustificationValues.Right });
        else if (alignLeft)
            paraProps.Append(new Justification { Val = JustificationValues.Left });
        else if (alignCenter || isLabel)
            paraProps.Append(new Justification { Val = JustificationValues.Center });

        var runProps = new RunProperties();
        if (isLabel) runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = "16" }); // 8pt

        if (useDecimalTab)
        {
            // <w:tab/> + <w:t>...</w:t> — decimal tab 가 \t 위치를 기준으로 정렬
            var run = new Run(runProps,
                new TabChar(),
                new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
            return new TableCell(cellProps, new Paragraph(paraProps, run));
        }

        if (!string.IsNullOrEmpty(text) && (text.Contains('\n') || text.Contains('\r')))
        {
            // 줄바꿈 포함 — "원본\n↳ 매칭명" 같은 다중 라인 셀 (시료명 매칭 표시)
            // 매칭 라인 (↳ 시작) 은 회색 + 이탤릭 으로 시각 구분.
            var parts = text.Replace("\r\n", "\n").Split('\n');
            var para = new Paragraph(paraProps);
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) para.Append(new Run(new Break()));
                var line = parts[i];
                bool isMatched = line.TrimStart().StartsWith("↳");
                var lineProps = new RunProperties();
                if (isLabel) lineProps.Append(new Bold());
                lineProps.Append(new FontSize { Val = "16" });
                if (isMatched)
                {
                    lineProps.Append(new Italic());
                    lineProps.Append(new Color { Val = "777777" });
                }
                para.Append(new Run(lineProps,
                    new Text(line) { Space = SpaceProcessingModeValues.Preserve }));
            }
            return new TableCell(cellProps, para);
        }

        var simpleRun = new Run(runProps, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        return new TableCell(cellProps, new Paragraph(paraProps, simpleRun));
    }

    private static Paragraph MakeParagraph(string text, int fontSize = 10)
    {
        return new Paragraph(
            new Run(
                new RunProperties(
                    new FontSize { Val = (fontSize * 2).ToString() },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }
                ),
                new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }
            )
        );
    }
}
