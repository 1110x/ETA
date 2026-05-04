using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using ETA.Models;

namespace ETA.Services.Common;

/// <summary>측정대행 접수/발송 대장 — Word(.docx) 출력. 엑셀과 동일 6컬럼 + 서명 푸터 + 회사명.
/// 다중 페이지 시 표 머리글 자동 반복. 회사 로고를 헤더 워터마크로 삽입.</summary>
public static class ReceiptDispatchWordExporter
{
    private static readonly string[] Headers =
        { "접수번호", "접수일", "시료명", "의뢰인 및 업체명", "분석항목", "발송일" };

    /// <summary>컬럼 너비(dxa) — A4 세로 페이지 폭 ~10466 dxa 기준. 분석항목 가장 넓게.</summary>
    private static readonly int[] ColWidths = { 1500, 1000, 1900, 1900, 3050, 1100 }; // 합계 10450

    public static void Export(string path, List<ReceiptDispatchEntry> rows)
    {
        if (File.Exists(path)) File.Delete(path);

        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = mainPart.Document.AppendChild(new Body());

        // 페이지 설정 — A4 세로 (11906 × 16838 dxa). 마지막 섹션은 Bottom 정렬로 결재칸을 페이지 하단에 고정.
        var sectPr = new SectionProperties(
            new PageSize { Width = 11906U, Height = 16838U, Orient = PageOrientationValues.Portrait },
            new PageMargin { Top = 720, Bottom = 720, Left = 720, Right = 720, Header = 360, Footer = 360, Gutter = 0 },
            new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Bottom });
        body.AppendChild(sectPr);

        // 1) 제목 — 별도 섹션 (Top 정렬) 으로 분리: 데이터/제목은 위에서부터, 결재칸만 마지막 페이지 하단으로
        var year = DateTime.Now.Year;
        body.InsertBefore(MakeTitle($"{year}년 측정대행기록부 접수/발송 대장"), sectPr);

        // 2) 데이터 표 (헤더 행 자동 반복 — Word 가 페이지 분할 시 헤더 재출력)
        var table = new Table();
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder    { Val = BorderValues.Single, Size = 8 },
                new BottomBorder { Val = BorderValues.Single, Size = 8 },
                new LeftBorder   { Val = BorderValues.Single, Size = 8 },
                new RightBorder  { Val = BorderValues.Single, Size = 8 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4 }),
            new TableJustification { Val = TableRowAlignmentValues.Center },
            new TableLayout { Type = TableLayoutValues.Fixed });
        table.AppendChild(tblProps);

        var tblGrid = new TableGrid();
        foreach (var w in ColWidths)
            tblGrid.Append(new GridColumn { Width = w.ToString() });
        table.AppendChild(tblGrid);

        // 헤더 행 — 페이지 분할 시 자동 반복
        table.AppendChild(MakeHeaderRow());

        foreach (var e in rows)
        {
            table.AppendChild(MakeDataRow(
                e.접수번호, e.접수일, e.시료명, e.업체명, e.분석항목, e.발송일));
        }
        body.InsertBefore(table, sectPr);

        // 섹션 분리 — 데이터 표(Section1, Top 정렬) / 결재칸+회사명(Section2, Bottom 정렬)
        // Continuous section break: 새 페이지 강제 없이 같은 페이지 흐름 유지하되 정렬만 분리
        var section1Pr = new SectionProperties(
            new SectionType { Val = SectionMarkValues.Continuous },
            new PageSize { Width = 11906U, Height = 16838U, Orient = PageOrientationValues.Portrait },
            new PageMargin { Top = 720, Bottom = 720, Left = 720, Right = 720, Header = 360, Footer = 360, Gutter = 0 },
            new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Top });
        var sectionBreakPara = new Paragraph(new ParagraphProperties(section1Pr));
        body.InsertBefore(sectionBreakPara, sectPr);

        // 3) 서명 푸터 표 (시료접수 / 시험성적서 발송) — Section 2 시작 (Bottom 정렬)
        body.InsertBefore(MakeSpacing(), sectPr);
        body.InsertBefore(MakeSignatureBlock(), sectPr);

        // 4) 회사명 (큰 글씨, 가운데)
        body.InsertBefore(MakeSpacing(), sectPr);
        body.InsertBefore(MakeCompanyFooter("리뉴어스 수질분석센터"), sectPr);

        // 5) 회사 로고 워터마크 — 헤더 파트에 삽입
        TryAddLogoWatermark(mainPart);

        mainPart.Document.Save();
    }

    private static Paragraph MakeTitle(string text)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "120", After = "120" });
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new Bold(),
            new FontSize { Val = "32" },
            new Color { Val = "1F4E78" });
        var run = new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(pPr, run);
    }

    private static TableRow MakeHeaderRow()
    {
        var trPr = new TableRowProperties(new TableHeader());   // 헤더 반복 표시
        var row = new TableRow(trPr);
        for (int i = 0; i < Headers.Length; i++)
        {
            row.AppendChild(MakeHeaderCell(Headers[i], ColWidths[i]));
        }
        return row;
    }

    private static TableCell MakeHeaderCell(string text, int width)
    {
        var tcPr = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new Shading { Fill = "D9E2F3", Val = ShadingPatternValues.Clear, Color = "auto" },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center });
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new Bold(),
            new FontSize { Val = "20" },
            new Color { Val = "1F4E78" });
        return new TableCell(tcPr, new Paragraph(pPr, new Run(rPr, new Text(text))));
    }

    private static TableRow MakeDataRow(params string[] cells)
    {
        // 기본 1줄(360 dxa) — 줄바꿈 발생 시 자동 확장 (AtLeast)
        var trPr = new TableRowProperties(
            new TableRowHeight { Val = 360U, HeightType = HeightRuleValues.AtLeast });
        var row = new TableRow(trPr);
        // 모든 데이터 셀 통일된 글자 크기 — 7pt (FontSize Val=14, half-points)
        for (int i = 0; i < cells.Length; i++)
        {
            row.AppendChild(MakeDataCell(cells[i] ?? "", ColWidths[i], 14));
        }
        return row;
    }

    private static TableCell MakeDataCell(string text, int width, int fontHalfPt = 18)
    {
        var tcPr = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center });
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new FontSize { Val = fontHalfPt.ToString() });
        return new TableCell(tcPr, new Paragraph(pPr, new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static Paragraph MakeSpacing()
        => new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "120", After = "120" }));

    private static Table MakeSignatureBlock()
    {
        var tbl = new Table();
        var tblPr = new TableProperties(
            new TableBorders(
                new TopBorder    { Val = BorderValues.Single, Size = 8 },
                new BottomBorder { Val = BorderValues.Single, Size = 8 },
                new LeftBorder   { Val = BorderValues.Single, Size = 8 },
                new RightBorder  { Val = BorderValues.Single, Size = 8 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4 }),
            new TableJustification { Val = TableRowAlignmentValues.Center },
            new TableLayout { Type = TableLayoutValues.Fixed });
        tbl.AppendChild(tblPr);

        // 4 열: 라벨1(시료접수) | 서명1 | 라벨2(시험성적서 발송) | 서명2
        int[] sigW = { 1500, 3725, 1500, 3725 }; // 합 10450 (세로 양식)
        var grid = new TableGrid();
        foreach (var w in sigW) grid.Append(new GridColumn { Width = w.ToString() });
        tbl.AppendChild(grid);

        string[] sigLines =
        {
            "담  당  자 :                       (서명)",
            "기술책임자 :                       (서명)",
            "품질책임자 :                       (서명)",
        };

        for (int i = 0; i < 3; i++)
        {
            // 컴팩트한 행 높이 — 한 줄만 차지 (320 dxa ≈ 16pt)
            var row = new TableRow(new TableRowProperties(
                new TableRowHeight { Val = 320U, HeightType = HeightRuleValues.Exact }));

            // 시료접수 라벨 — 첫 행만 vMerge restart, 나머지 continue
            var labelCell1 = MakeMergedLabelCell("시료접수", sigW[0], i == 0, "E8F4E8", "1F5E1F");
            row.AppendChild(labelCell1);
            row.AppendChild(MakeSignatureCell(sigLines[i], sigW[1]));

            var labelCell2 = MakeMergedLabelCell("시험성적서 발송", sigW[2], i == 0, "E2EAF6", "1F4E78");
            row.AppendChild(labelCell2);
            row.AppendChild(MakeSignatureCell(sigLines[i], sigW[3]));

            tbl.AppendChild(row);
        }
        return tbl;
    }

    private static TableCell MakeMergedLabelCell(string text, int width, bool restart, string fillHex, string textHex)
    {
        var tcPr = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new VerticalMerge { Val = restart ? MergedCellValues.Restart : MergedCellValues.Continue },
            new Shading { Fill = fillHex, Val = ShadingPatternValues.Clear, Color = "auto" },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        if (!restart)
            return new TableCell(tcPr, new Paragraph());

        var pPr = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new Bold(),
            new FontSize { Val = "18" },   // 9pt — 컴팩트
            new Color { Val = textHex });
        return new TableCell(tcPr, new Paragraph(pPr, new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static TableCell MakeSignatureCell(string text, int width)
    {
        var tcPr = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        var pPr = new ParagraphProperties(new Justification { Val = JustificationValues.Left });
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new FontSize { Val = "16" });   // 8pt
        return new TableCell(tcPr, new Paragraph(pPr, new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static Paragraph MakeCompanyFooter(string text)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "120", After = "120" },
            new ParagraphMarkRunProperties());
        // 셀 배경처럼 보이게 회색 사각형 — 단순하게 큰 글씨로
        var rPr = new RunProperties(
            new RunFonts { Ascii = "맑은 고딕", EastAsia = "맑은 고딕", HighAnsi = "맑은 고딕" },
            new Bold(),
            new FontSize { Val = "36" },
            new Color { Val = "1F4E78" });
        return new Paragraph(pPr, new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    /// <summary>회사 로고 워터마크 — 헤더 파트에 반투명 이미지 삽입.</summary>
    private static void TryAddLogoWatermark(MainDocumentPart mainPart)
    {
        try
        {
            var logoPath = Path.Combine(AppPaths.RootPath, "Assets", "icons", "renewus_vertical_black.png");
            if (!File.Exists(logoPath)) return;

            // 헤더 파트 추가 (모든 페이지 공통)
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var imagePart = headerPart.AddImagePart(ImagePartType.Png);
            using (var fs = File.OpenRead(logoPath))
                imagePart.FeedData(fs);

            string relId = headerPart.GetIdOfPart(imagePart);

            // 워터마크 — 페이지 가운데, 큰 사이즈, 반투명
            // EMU: 914400 = 1 inch
            long widthEmu  = (long)(5.5 * 914400);
            long heightEmu = (long)(3.8 * 914400);

            var element = new Drawing(
                new DW.Anchor(
                    new DW.SimplePosition { X = 0, Y = 0 },
                    new DW.HorizontalPosition(new DW.HorizontalAlignment("center")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
                    new DW.VerticalPosition(new DW.VerticalAlignment("center")) { RelativeFrom = DW.VerticalRelativePositionValues.Page },
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    new DW.WrapNone(),
                    new DW.DocProperties { Id = 1U, Name = "Watermark" },
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = "logo.png" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(new A.AlphaModulationFixed { Amount = 25000 }) { Embed = relId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0, Y = 0 },
                                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                {
                    DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U,
                    SimplePos = false, RelativeHeight = 0U, BehindDoc = true, Locked = false, LayoutInCell = true, AllowOverlap = true,
                });

            var headerPara = new Paragraph(new Run(element));
            headerPart.Header = new Header(headerPara);
            headerPart.Header.Save();

            string headerRelId = mainPart.GetIdOfPart(headerPart);
            // section 에 헤더 참조 추가
            var sectPr = mainPart.Document.Body!.Elements<SectionProperties>().LastOrDefault();
            sectPr?.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelId });
        }
        catch { /* 워터마크 실패해도 본문은 출력 */ }
    }
}
