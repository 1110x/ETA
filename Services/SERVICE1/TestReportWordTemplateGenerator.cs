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
///     1행: "{{시험성적서서명}}  (서명)" — 라벨 전체를 결재정보(설정→결재정보)에서 자유 입력
///          (마지막 "(서명)" 만 이탤릭/회색)
///     2행: {{비고문구}} — ▩ ES 안내 문구
///
/// 치환자 (TestReportWordPrintService.FillAndSave 가 채움):
///   헤더: {{성적서번호}} {{회사명}} {{대표자}} {{시료명}} {{채취일자}} {{채취자}}
///         {{입회자}} {{분석종료일}} {{용도}} {{비고문구}} {{시험성적서서명}}
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

        // 0) 성적서번호는 헤더로 이동 — 본문 첫 단락에는 제목만 (BuildReportNoLine 제거)

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

        // 모든 페이지 동일 헤더 — 성적서번호 + 페이지번호. 워터마크 없음.
        _ = watermarkText; _ = logoPath;
        var headerPart = main.AddNewPart<HeaderPart>();
        using (var s = headerPart.GetStream(FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(s, new System.Text.UTF8Encoding(false)))
            sw.Write(BuildEmptyHeaderXml());

        // ── 섹션 구성 ───────────────────────────────────────────────────────
        // 단일 섹션. 다중페이지 양식은 결재라인 직전에 PageBreakBefore 로 새 페이지 분리.
        // (섹션 브레이크 NextPage 는 빈 페이지 발생 이슈 있어 단일 섹션 + 페이지 브레이크 조합 사용.)
        // multiPage 분기는 본문 빌드 단계에서 처리.
        _ = multiPage; // 향후 섹션별 마진 차등화 시 사용 — 현재는 단일 섹션
        var headerId = main.GetIdOfPart(headerPart);
        var footerId = main.GetIdOfPart(footerPart);

        // 본문 영역 — Top 1134(2cm, 32항목 한 페이지 수용 위해 3cm → 2cm 축소).
        // Bottom 1800 (직인 영역 진입 방지).
        _pendingSectionProps = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
            new FooterReference { Type = HeaderFooterValues.Default, Id = footerId },
            new PageSize  { Width = 11906U, Height = 16838U },
            new PageMargin
            {
                Top    = 1134, Right  = 720, Bottom = 1800, Left = 720,
                Header = 360,  Footer = 720, Gutter = 0,
            });
    }

    /// <summary>페이지 footer XML — 회사명(24pt 굵게) 만 표시.
    /// 페이지번호는 더 이상 footer 가 아니라 헤더(성적서번호 옆)로 이동.</summary>
    private static string BuildFooterXml()
    {
        return
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:ftr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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

    /// <summary>헤더 XML — 우측 상단에 "성적서번호: {{성적서번호}}     (PAGE/NUMPAGES)".
    /// {{성적서번호}} 는 서비스 치환자, PAGE/NUMPAGES 는 Word 자동 갱신 필드.</summary>
    private static string BuildEmptyHeaderXml() =>
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:p>
    <w:pPr>
      <w:jc w:val=""right""/>
      <w:spacing w:before=""0"" w:after=""0""/>
    </w:pPr>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/></w:rPr>
      <w:t xml:space=""preserve"">성적서번호: {{성적서번호}}     (</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText xml:space=""preserve""> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""separate""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:t>/</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""begin""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:instrText xml:space=""preserve""> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""separate""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:fldChar w:fldCharType=""end""/>
    </w:r>
    <w:r>
      <w:rPr><w:sz w:val=""16""/><w:color w:val=""555555""/></w:rPr>
      <w:t>)</w:t>
    </w:r>
  </w:p>
</w:hdr>";

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

    // ─── 메인 타이틀 ────────────────────────────────────────────────────────
    private static Paragraph BuildTitle()
    {
        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "120", Before = "0" }),
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

    // ─── 메타정보 6열 페어 표 (라벨|값|라벨|값|라벨|값) — 3행 컴팩트 ───────
    // 32 항목을 1페이지에 수용하기 위해 4행 → 3행으로 압축, 행 높이도 축소.
    //   행1: 회사명   / 대표자 / 분석종료일
    //   행2: 시료명   / 용도   / 입회자
    //   행3: 채취일자 / 채취자
    // 빈 페어 자리는 공백 셀로 채움 (그리드 일관성).
    private static Table BuildHeaderTable()
    {
        // 6열: 라벨1 1000 / 값1 2400 / 라벨2 1000 / 값2 2400 / 라벨3 1000 / 값3 2600 = 10400
        int[] widths = { 1000, 2400, 1000, 2400, 1000, 2600 };
        var table = NewBorderedTable();
        AddColumnGrid(table, widths);

        table.Append(MakePairRow(widths,
            ("회사명",     "{{회사명}}"),
            ("대표자",     "{{대표자}}"),
            ("분석종료일", "{{분석종료일}}")));
        table.Append(MakePairRow(widths,
            ("시료명",     "{{시료명}}"),
            ("용도",       "{{용도}}"),
            ("입회자",     "{{입회자}}")));
        table.Append(MakePairRow(widths,
            ("채취일자",   "{{채취일자}}"),
            ("채취자",     "{{채취자}}"),
            ("",           "")));
        return table;
    }

    private static TableRow MakePairRow(
        int[] widths,
        params (string label, string value)[] pairs)
    {
        // 행 높이 280 → 220 (행마다 60 dxa 절약 × 3행 = 180 dxa 추가 절약).
        var row = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 220U, HeightType = HeightRuleValues.AtLeast }));
        for (int i = 0; i < 3; i++)
        {
            var (l, v) = i < pairs.Length ? pairs[i] : ("", "");
            row.Append(MakeCell(l, width: widths[i * 2],     isLabel: true, alignCenter: true));
            row.Append(MakeCell(v, width: widths[i * 2 + 1], alignLeft: true));
        }
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

        // 헤더 — 옅은 회색 배경, 굵게, 가운데. 셀 spacing 0 + 행 240.
        string[] headers = { "번호", "구분", "항목", "ES", "결과", "단위", "기준" };
        var hdr = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 240U, HeightType = HeightRuleValues.AtLeast }));
        for (int i = 0; i < headers.Length; i++)
            hdr.Append(MakeCell(headers[i], width: widths[i], isLabel: true, alignCenter: true, headerShade: true, tightSpacing: true));
        table.Append(hdr);

        // 시료 항목 행 — 치환자 (모든 셀 가운데 정렬). tightSpacing 으로 행마다 ≈40 dxa 절약.
        var row = new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 240U, HeightType = HeightRuleValues.AtLeast }));
        row.Append(MakeCell("{{번호}}", width: widths[0], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{구분}}", width: widths[1], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{항목}}", width: widths[2], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{ES}}",   width: widths[3], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{결과}}", width: widths[4], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{단위}}", width: widths[5], alignCenter: true, tightSpacing: true));
        row.Append(MakeCell("{{기준}}", width: widths[6], alignCenter: true, tightSpacing: true));
        table.Append(row);

        return table;
    }

    // ─── 마지막 페이지 하단 anchor 표 — 서명 + 비고문구 ─────────────────
    // tblpPr 로 PageMargin 기준 하단 정렬 → 항상 마지막 페이지 가장 밑에 위치.
    // 단일 행 + 단일 셀 + 단락 2개 구조 (서명 단락 + 비고문구 단락) — 행 분리로 페이지가
    // 갈라지는 문제 회피. 셀 자체가 한 단위라 anchored 표 내용은 절대 분리되지 않음.
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

        // 표 자체를 "페이지" 자체 기준 하단으로 anchor — Margin 기준이면 본문이 짧은 페이지에서
        // 본문 bottom 이 페이지 상단쪽으로 올라와 결재라인이 위로 떠 보이는 문제가 있어 Page 로 변경.
        // A4 높이 = 16838 dxa. 표 top = 14000 dxa 로 고정 → 표 bottom ≈ 14740
        // → footer(720) 위까지 약 16838 − 720 − 14740 = 1378 dxa ≈ 2.4cm (직인 영역 확보).
        // (32항목 + 결재라인 + footer 가 한 장에 들어가도록 TablePositionY 13200 → 14000.)
        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null)
        {
            tp.InsertAt(new TablePositionProperties
            {
                LeftFromText  = 0, RightFromText = 0,
                TopFromText   = 0, BottomFromText = 0,
                VerticalAnchor   = VerticalAnchorValues.Page,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionY          = 14000,
            }, 0);
        }

        AddColumnGrid(table, 10400);

        // 단일 행 + 단일 셀 — 행 분리로 인한 페이지 갈라짐 방지.
        // cantSplit 으로 셀이 두 페이지에 걸치지 못하게 강제.
        var row = new TableRow(
            new TableRowProperties(
                new CantSplit(),
                new TableRowHeight { Val = 740U, HeightType = HeightRuleValues.AtLeast }));
        row.Append(BuildFooterCombinedCell());
        table.Append(row);

        return table;
    }

    /// <summary>서명 단락 + 비고문구 단락을 한 셀에 묶음.
    /// 단일 행/단일 셀 안에 두 단락이 들어가므로 페이지 경계에서 분리되지 않음.</summary>
    private static TableCell BuildFooterCombinedCell()
    {
        var cellProps = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "10400" },
            new TableCellBorders(
                new LeftBorder   { Val = BorderValues.None },
                new RightBorder  { Val = BorderValues.None },
                new TopBorder    { Val = BorderValues.None },
                new BottomBorder { Val = BorderValues.None }),
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        // 단락 1: 서명 — "{{시험성적서서명}}  (서명)" 가운데정렬
        var sigPara = new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "60", After = "60" },
                new KeepNext()),                                                // 다음 단락(비고문구)과 같은 페이지 보장
            new Run(
                new RunProperties(
                    new FontSize { Val = "16" },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
                new Text("{{시험성적서서명}}") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(
                new RunProperties(
                    new Italic(),
                    new FontSize { Val = "16" },
                    new Color { Val = "777777" },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
                new Text("       (서명)") { Space = SpaceProcessingModeValues.Preserve }));

        // 단락 2: 비고문구 — 7pt 가운데정렬
        var remarkPara = new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "60", After = "0" }),
            new Run(
                new RunProperties(
                    new FontSize { Val = "14" },
                    new RunFonts { Ascii = FontFamily, EastAsia = FontFamily, HighAnsi = FontFamily }),
                new Text("{{비고문구}}") { Space = SpaceProcessingModeValues.Preserve }));

        return new TableCell(cellProps, sigPara, remarkPara);
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
        bool alignLeft = false, bool headerShade = false, int? fontHalfPt = null,
        bool tightSpacing = false)
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
        // 분석결과 행은 셀 상하 마진을 0 으로 (Word 기본 0 이지만 명시적으로 강제)
        if (tightSpacing)
            cellProps.Append(new TableCellMargin(
                new TopMargin    { Width = "0", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa }));

        // 분석결과 셀은 단락 spacing 을 0 으로 — 32행 × 40 dxa ≈ 1280 dxa 절약.
        var spacing = tightSpacing
            ? new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
            : new SpacingBetweenLines { Before = "20", After = "20" };
        var paraProps = new ParagraphProperties(spacing);
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
