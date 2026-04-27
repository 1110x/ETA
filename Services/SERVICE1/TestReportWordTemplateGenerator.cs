using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 시험성적서 Word(.docx) 템플릿 생성기.
///
/// 시각언어: 시험기록부 Word(`TestRecordBookWordExporter`) 와 동일 — 가로선 표,
/// 맑은 고딕 8pt 본문, 9~10pt 섹션 타이틀, A4 portrait, 좌우 0.5", 중앙 워터마크
/// (renewus 로고 40% 불투명).
///
/// 레이아웃 (원본 시험성적서 Excel 구조 준수):
///   [상단 우측]   성적서번호       (작은 라벨 우측정렬)
///   [중앙]        시  험  성  적  서  (24pt bold)
///   [메타 표]     4열 페어 레이아웃 (라벨|값|라벨|값) × 5행
///   [섹션]        분석결과 — 7열 표 (번호/구분/항목/ES/결과/단위/기준)
///                              헤더행 + 시료행 1개(치환자) — 시료 항목 수만큼 복제
///   [하단 anchor 표] (마지막 페이지 가장 밑 — tblpPr 로 PageMargin 하단 정렬)
///     1행: "품질책임 수질분야 환경측정분석사  {{품질책임자}}  (서명)"
///          (마지막 "(서명)" 만 이탤릭/회색)
///     2행: {{비고문구}} — ▩ ES 안내 문구
///
/// 치환자 (TestReportWordPrintService.FillAndSave 가 채움):
///   헤더: {{성적서번호}} {{회사명}} {{대표자}} {{시료명}} {{채취일자}} {{채취자}}
///         {{입회자}} {{분석종료일}} {{용도}} {{비고문구}} {{품질책임자}}
///   시료행: {{번호}} {{구분}} {{항목}} {{ES}} {{결과}} {{단위}} {{기준}}
/// </summary>
public static class TestReportWordTemplateGenerator
{
    private const string FontFamily = "맑은 고딕";

    /// <summary>레거시 API — 기본 양식(다중 페이지) 생성. 신규 코드는 GenerateSinglePage / GenerateMultiPage 사용 권장.</summary>
    public static void Generate(string savePath, string? logoPath = null)
        => GenerateMultiPage(savePath, logoPath);

    /// <summary>단일 페이지 양식 — 항목 수가 적어 한 장에 다 들어가는 케이스.
    /// 섹션 분리 없음. 결재라인이 결과 표 바로 아래 인라인. 페이지 하단 footer 에 회사명.</summary>
    public static void GenerateSinglePage(string savePath, string? logoPath = null)
        => GenerateInternal(savePath, logoPath, multiPage: false);

    /// <summary>다중 페이지 양식 — 항목 수가 많아 여러 장이 필요한 케이스.
    /// 섹션 1(결과 페이지, 좁은 마진) + 섹션 2(결재라인 전용 페이지, 넓은 마진) 로 분리.</summary>
    public static void GenerateMultiPage(string savePath, string? logoPath = null)
        => GenerateInternal(savePath, logoPath, multiPage: true);

    private static void GenerateInternal(string savePath, string? logoPath, bool multiPage)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(savePath)!);
        if (File.Exists(savePath)) File.Delete(savePath);

        using var doc = WordprocessingDocument.Create(savePath, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body());
        var body = main.Document.Body!;

        ApplyDocDefaults(main);
        ApplyPageSetup(main, body, watermarkText: "ETA", logoPath: logoPath, multiPage: multiPage);

        // 0) 우상단 성적서번호 (작은 라벨)
        body.Append(BuildReportNoLine());

        // 1) 타이틀
        body.Append(BuildTitle());

        // 2) 메타정보 — 4열 페어 (라벨|값|라벨|값)
        body.Append(BuildHeaderTable());
        body.Append(SmallSpacer());

        // 3) 분석결과 표 (시료 항목 수만큼 복제될 행 1개 포함)
        AppendSectionTitle(body, "분석결과");
        body.Append(BuildResultTable());

        // 4) 결재라인 (서명 + 비고문구) — 결과 표 바로 아래 인라인 흐름.
        // 다중페이지 양식도 결과가 마지막 페이지 끝까지 차면 결재라인이 그 아래(자연 분기 또는 같은 페이지) 에 위치.
        body.Append(BuildFooterAnchored());

        if (_pendingSectionProps != null)
        {
            body.Append(_pendingSectionProps);
            _pendingSectionProps = null;
        }

        main.Document.Save();
    }

    // ─── 페이지 설정 + 중앙 워터마크 ──────────────────────────────────────────
    [ThreadStatic] private static SectionProperties? _pendingSectionProps;

    private static void ApplyPageSetup(MainDocumentPart main, Body body, string watermarkText, string? logoPath, bool multiPage)
    {
        // 페이지 footer — 회사명 + 조건부 페이지번호 (2장 이상일 때만 "n/N")
        // IF 필드로 NUMPAGES > 1 검사, raw XML 로 작성 (OpenXML SDK 의 중첩 필드는 verbose)
        var footerPart = main.AddNewPart<FooterPart>();
        using (var fs = footerPart.GetStream(FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(fs, new System.Text.UTF8Encoding(false)))
            sw.Write(BuildFooterXml());

        var headerPart = main.AddNewPart<HeaderPart>();

        // renewus 로고 PNG 를 헤더에 임베드 — 페이지 중앙에 40% 불투명 워터마크.
        // logoPath 미지정 시 ETA 기본 위치(<repo>/Assets/icons/renewus_vertical_black.png).
        string? logoRelId = null;
        try
        {
            string resolved = logoPath ?? Path.Combine(
                ETA.Services.Common.AppPaths.RootPath,
                "Assets", "icons", "renewus_vertical_black.png");
            if (File.Exists(resolved))
            {
                using var logoStream = File.OpenRead(resolved);
                var imagePart = headerPart.AddImagePart(ImagePartType.Png);
                imagePart.FeedData(logoStream);
                logoRelId = headerPart.GetIdOfPart(imagePart);
            }
        }
        catch { /* 로고 없으면 텍스트 워터마크로 폴백 */ }

        string esc = System.Net.WebUtility.HtmlEncode(watermarkText ?? "ETA");
        string headerXml = !string.IsNullOrEmpty(logoRelId)
            ? BuildImageWatermarkHeaderXml(logoRelId)
            : BuildTextWatermarkHeaderXml(esc);

        using (var s = headerPart.GetStream(FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(s, new System.Text.UTF8Encoding(false)))
            sw.Write(headerXml);

        // ── 섹션 구성 ───────────────────────────────────────────────────────
        // 단일 섹션. 다중페이지 양식은 결재라인 직전에 PageBreakBefore 로 새 페이지 분리.
        // (섹션 브레이크 NextPage 는 빈 페이지 발생 이슈 있어 단일 섹션 + 페이지 브레이크 조합 사용.)
        // multiPage 분기는 본문 빌드 단계에서 처리.
        _ = multiPage; // 향후 섹션별 마진 차등화 시 사용 — 현재는 단일 섹션
        var headerId = main.GetIdOfPart(headerPart);
        var footerId = main.GetIdOfPart(footerPart);

        _pendingSectionProps = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
            new FooterReference { Type = HeaderFooterValues.Default, Id = footerId },
            new PageSize  { Width = 11906U, Height = 16838U },
            new PageMargin
            {
                Top    = 720, Right  = 720, Bottom = 2880, Left = 720,
                Header = 360, Footer = 1200, Gutter = 0,
            });
    }

    /// <summary>페이지 footer XML — 회사명(24pt 굵게) + 조건부 페이지번호.
    /// IF 필드: NUMPAGES > 1 일 때만 "PAGE/NUMPAGES" 표시, 1장이면 빈 문자열.
    /// 페이지번호는 회사명 위에 작게(8pt, #555) 배치 — 직인 영역(하단 3.8cm) 침범 안 함.</summary>
    private static string BuildFooterXml()
    {
        return
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:ftr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:p>
    <w:pPr>
      <w:jc w:val=""center""/>
      <w:spacing w:before=""0"" w:after=""60""/>
    </w:pPr>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText xml:space=""preserve""> IF </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:instrText xml:space=""preserve""> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText xml:space=""preserve""> &gt; 1 &quot;</w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:instrText xml:space=""preserve""> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText>/</w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:instrText xml:space=""preserve""> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText xml:space=""preserve"">&quot; &quot;&quot;</w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
  </w:p>
  <w:p>
    <w:pPr>
      <w:jc w:val=""center""/>
      <w:spacing w:before=""0"" w:after=""0"" w:line=""720"" w:lineRule=""atLeast""/>
    </w:pPr>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/></w:rPr>
      <w:instrText xml:space=""preserve""> IF </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:instrText xml:space=""preserve""> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/></w:rPr>
      <w:instrText xml:space=""preserve""> = </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:instrText xml:space=""preserve""> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""54""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/></w:rPr>
      <w:instrText xml:space=""preserve""> &quot;리뉴어스 수질분석센터&quot; &quot;&quot;</w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/></w:rPr>
      <w:fldChar w:fldCharType=""separate""/>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/><w:spacing w:val=""100""/><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/></w:rPr>
      <w:t>리뉴어스 수질분석센터</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:b/><w:sz w:val=""54""/><w:color w:val=""1F4E79""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
  </w:p>
</w:ftr>";
    }

    /// <summary>로고 PNG 임베드 워터마크 — behindDoc=1, alphaModFix 로 40% 불투명.</summary>
    private static string BuildImageWatermarkHeaderXml(string relId)
    {
        // renewus 로고 비율 ≈ 1.44 (가로:세로). 173pt × 120pt = 2197100 × 1524000 EMU
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
        stylePart.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                        new FontSize { Val = "16" } // 8pt
                    )),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { After = "0", Before = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                    ))));
        stylePart.Styles.Save();
    }

    // ─── 우상단 성적서번호 라인 ────────────────────────────────────────────
    private static Paragraph BuildReportNoLine()
    {
        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Right },
                new SpacingBetweenLines { Before = "0", After = "120" }),
            new Run(
                new RunProperties(
                    new FontSize { Val = "16" }, // 8pt
                    new Color { Val = "555555" },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
                new Text("성적서번호: {{성적서번호}}") { Space = SpaceProcessingModeValues.Preserve }));
    }

    // ─── 메인 타이틀 ────────────────────────────────────────────────────────
    private static Paragraph BuildTitle()
    {
        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "240", Before = "120" }),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "36" }, // 18pt
                    new CharacterScale { Val = 100 },
                    new Spacing { Val = 60 },     // 자간 0.3pt
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
                new Text("시  험  성  적  서") { Space = SpaceProcessingModeValues.Preserve }));
    }

    // ─── 섹션 타이틀 (파란 바 + 텍스트) ───────────────────────────────────────
    private static void AppendSectionTitle(Body body, string title)
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
            new Run(
                new RunProperties(new Bold(), new FontSize { Val = "18" },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily },
                    new Color { Val = "1F4E79" }),
                new Text(title))));

        t.Append(new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 200U, HeightType = HeightRuleValues.AtLeast }),
            barCell, txtCell));

        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "0" })));
        body.Append(t);
        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "30" })));
    }

    private static Paragraph SmallSpacer() =>
        new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "60" }));

    // ─── 메타정보 4열 페어 표 (라벨|값|라벨|값) ──────────────────────────
    // 원본 Excel 시험성적서 헤더 구조 대응:
    //   회사명 / 채취일자
    //   대표자 / 채취자
    //   시료명 / 입회자
    //   용도   / 분석종료일
    private static Table BuildHeaderTable()
    {
        // 4열 너비: 라벨1 1500 / 값1 3700 / 라벨2 1500 / 값2 3700 = 10400
        int[] widths = { 1500, 3700, 1500, 3700 };
        var table = NewBorderedTable();
        AddColumnGrid(table, widths);

        table.Append(MakePairRow("회사명", "{{회사명}}",   "채취일자",   "{{채취일자}}",   widths));
        table.Append(MakePairRow("대표자", "{{대표자}}",   "채취자",     "{{채취자}}",     widths));
        table.Append(MakePairRow("시료명", "{{시료명}}",   "입회자",     "{{입회자}}",     widths));
        table.Append(MakePairRow("용도",   "{{용도}}",     "분석종료일", "{{분석종료일}}", widths));
        return table;
    }

    private static TableRow MakePairRow(
        string l1, string v1, string l2, string v2, int[] widths)
    {
        var row = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 360U, HeightType = HeightRuleValues.AtLeast }));
        row.Append(MakeCell(l1, width: widths[0], isLabel: true, alignCenter: true));
        row.Append(MakeCell(v1, width: widths[1], alignLeft: true));
        row.Append(MakeCell(l2, width: widths[2], isLabel: true, alignCenter: true));
        row.Append(MakeCell(v2, width: widths[3], alignLeft: true));
        return row;
    }

    // ─── 분석결과 표 ──────────────────────────────────────────────────────────
    // 헤더 행 + 시료 행 1개(치환자) — TestReportWordPrintService 가 시료 항목 수만큼 복제.
    private static Table BuildResultTable()
    {
        // 번호 / 구분 / 항목 / ES / 결과 / 단위 / 기준 = 10400
        // 구분 2200: "휘발성유기화합물"(8자) 한 줄 들어가도록 충분히 확보
        int[] widths = { 500, 2200, 2500, 1500, 1400, 900, 1400 };
        var table = NewBorderedTable();
        AddColumnGrid(table, widths);

        // 헤더 — 옅은 회색 배경, 굵게, 가운데
        string[] headers = { "번호", "구분", "항목", "ES", "결과", "단위", "기준" };
        var hdr = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 320U, HeightType = HeightRuleValues.AtLeast }));
        for (int i = 0; i < headers.Length; i++)
            hdr.Append(MakeCell(headers[i], width: widths[i], isLabel: true, alignCenter: true, headerShade: true));
        table.Append(hdr);

        // 시료 항목 행 — 치환자 (모든 셀 가운데 정렬)
        var row = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 280U, HeightType = HeightRuleValues.AtLeast }));
        row.Append(MakeCell("{{번호}}", width: widths[0], alignCenter: true));
        row.Append(MakeCell("{{구분}}", width: widths[1], alignCenter: true));
        row.Append(MakeCell("{{항목}}", width: widths[2], alignCenter: true));
        row.Append(MakeCell("{{ES}}",   width: widths[3], alignCenter: true));
        row.Append(MakeCell("{{결과}}", width: widths[4], alignCenter: true));
        row.Append(MakeCell("{{단위}}", width: widths[5], alignCenter: true));
        row.Append(MakeCell("{{기준}}", width: widths[6], alignCenter: true));
        table.Append(row);

        return table;
    }

    // ─── 마지막 페이지 하단 anchor 표 — 서명 + 비고문구 ─────────────────
    // tblpPr 로 PageMargin 기준 하단 정렬 → 항상 마지막 페이지 가장 밑에 위치.
    // 1행: "품질책임 수질분야 환경측정분석사  {{품질책임자}}  (서명)" — 가운데정렬
    //      마지막 "(서명)" Run 만 이탤릭/회색 (별도 Run 으로 분리해 스타일 보존)
    // 2행: "{{비고문구}}" — 좌측정렬 8pt
    private static Table BuildFooterAnchored()
    {
        // 본문 하단 (PageMargin.Bottom 위치) 에 anchor — 결재라인은 마지막 페이지 가장 밑.
        // 가로선은 제거 (사용자 요청: "선은 없애고").
        var table = new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder    { Val = BorderValues.None },
                    new LeftBorder   { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.None },
                    new RightBorder  { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.None },
                    new InsideVerticalBorder   { Val = BorderValues.None }),
                new TableLayout { Type = TableLayoutValues.Fixed }));

        // 표 자체를 페이지 하단 마진 기준으로 anchor (스키마: tblpPr 는 tblPr 의 첫 자식)
        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null)
        {
            tp.InsertAt(new TablePositionProperties
            {
                LeftFromText  = 0, RightFromText = 0,
                TopFromText   = 0, BottomFromText = 0,
                VerticalAnchor   = VerticalAnchorValues.Margin,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            }, 0);
        }

        AddColumnGrid(table, 10400);

        // 1행: 서명 라인 (3 Run — label / 품책자 / "(서명)" 이탤릭)
        var sigRow = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 420U, HeightType = HeightRuleValues.AtLeast }));
        sigRow.Append(BuildSignatureCell());
        table.Append(sigRow);

        // 2행: 비고문구 (▩ ES 안내) — 7pt
        var remarkRow = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 320U, HeightType = HeightRuleValues.AtLeast }));
        remarkRow.Append(MakeCell("{{비고문구}}", width: 10400, alignCenter: true, fontHalfPt: 14));
        table.Append(remarkRow);

        return table;
    }

    /// <summary>서명 셀 — Run 3개로 구성하여 "(서명)" 만 이탤릭/회색 유지.
    /// 치환 로직(서비스)은 Text 노드 단위로 먼저 적용하므로 Run 스타일이 보존됨.</summary>
    private static TableCell BuildSignatureCell()
    {
        var cellProps = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "10400" },
            new TableCellBorders(
                new LeftBorder  { Val = BorderValues.None },
                new RightBorder { Val = BorderValues.None }),
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        var paraProps = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "60", After = "60" });

        // Run 1: "품질책임 수질분야 환경측정분석사       " — 8pt
        var runLabel = new Run(
            new RunProperties(
                new FontSize { Val = "16" }, // 8pt
                new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
            new Text("품질책임 수질분야 환경측정분석사       ") { Space = SpaceProcessingModeValues.Preserve });

        // Run 2: "{{품질책임자}}" — 8pt (단독 Text 라 per-Text 치환 시 스타일 유지)
        var runName = new Run(
            new RunProperties(
                new FontSize { Val = "16" },
                new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
            new Text("{{품질책임자}}") { Space = SpaceProcessingModeValues.Preserve });

        // Run 3: "       (서명)" 이탤릭 + 회색 (#777) 8pt
        var runSeal = new Run(
            new RunProperties(
                new Italic(),
                new FontSize { Val = "16" }, // 8pt
                new Color { Val = "777777" },
                new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
            new Text("       (서명)") { Space = SpaceProcessingModeValues.Preserve });

        return new TableCell(cellProps, new Paragraph(paraProps, runLabel, runName, runSeal));
    }

    // ─── 공통 헬퍼 (시험기록부 익스포터와 동일 스타일) ───────────────────
    private static Table NewBorderedTable()
    {
        return new Table(
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
    }

    private static void AddColumnGrid(Table table, params int[] widthsDxa)
    {
        var grid = new TableGrid();
        foreach (var w in widthsDxa) grid.Append(new GridColumn { Width = w.ToString() });
        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null) tp.InsertAfterSelf(grid);
        else table.InsertAt(grid, 0);
    }

    private static TableCell MakeCell(string text, int? width = null,
        bool isLabel = false, bool alignRight = false, bool alignCenter = false,
        bool alignLeft = false, bool headerShade = false, int? fontHalfPt = null)
    {
        var cellProps = new TableCellProperties();
        if (width.HasValue)
            cellProps.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.Value.ToString() });
        if (headerShade)
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F1ECE0" });
        cellProps.Append(new TableCellBorders(
            new LeftBorder  { Val = BorderValues.None },
            new RightBorder { Val = BorderValues.None }));
        cellProps.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        var paraProps = new ParagraphProperties(
            new SpacingBetweenLines { Before = "20", After = "20" });
        if (alignRight)       paraProps.Append(new Justification { Val = JustificationValues.Right });
        else if (alignLeft)   paraProps.Append(new Justification { Val = JustificationValues.Left });
        else if (alignCenter || isLabel)
                              paraProps.Append(new Justification { Val = JustificationValues.Center });

        var runProps = new RunProperties();
        if (isLabel) runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = (fontHalfPt ?? 16).ToString() });

        var run = new Run(runProps,
            new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        return new TableCell(cellProps, new Paragraph(paraProps, run));
    }
}
