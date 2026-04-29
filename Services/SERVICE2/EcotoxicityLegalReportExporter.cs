using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 환경부 ES 04704.1c 별지 1 — 원수 시료 결과 기록부 (생태독성 법정양식) Word 출력기.
/// 시료 정보 + 현장/수질 측정항목 + 독성시험결과 매트릭스 + EC₅₀/TU + 기입자 서명란.
/// </summary>
public static class EcotoxicityLegalReportExporter
{
    private const string Font   = "맑은 고딕";
    private const int    BodyPt = 8;     // 본문 8pt
    private const int    HdrPt  = 9;     // 섹션 헤더 9pt
    private const int    TitlePt = 14;   // 제목 14pt

    public sealed record FormData(
        string 배출시설,
        string 시료채취장소,
        string 시료번호,
        string 채취일자,
        // 현장 측정항목
        string 온도,
        string pH,
        string 용존산소,
        string 유입수량,
        string 전기전도도,
        string 염분,
        // 수질 측정항목
        string 잔류염소,
        string 암모니아,
        string 경도,
        // 독성시험결과 (행: Control, 6.25, 12.5, 25, 50, 100 / 열: 1~4 마리수, 유영저해율)
        string[][] 독성결과,   // 6 rows × 5 cols
        // 독성시험항목
        string EC50,
        string TU,
        string 통계분석법,
        // 작성
        string 작성일,        // YYYY-MM-DD
        string 기입자성명);

    public static (bool ok, string msg) Export(FormData d, string savePath)
    {
        try
        {
            using var doc = WordprocessingDocument.Create(savePath, WordprocessingDocumentType.Document);
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            // 페이지 / 여백
            var sectPr = new SectionProperties(
                new PageSize     { Width = 11906U, Height = 16838U },
                new PageMargin   { Top = 720, Right = 720, Bottom = 720, Left = 720, Header = 360, Footer = 360, Gutter = 0 });

            // 별지 1 표시 (표 외부)
            body.Append(P("<별지 1>", BodyPt, align: JustificationValues.Left));

            // ── 표 시작 ────────────────────────────────────────────────────
            var t = new Table(StdTableProps());
            // 6 컬럼 균등
            var grid = new TableGrid();
            for (int i = 0; i < 6; i++) grid.Append(new GridColumn { Width = "1700" });
            t.Append(grid);

            // 제목 행 — 표 최상단 (span 6, 가운데 정렬, 굵게)
            t.Append(RowNoVert(TitleCell("원수 시료 결과 기록부", 6)));

            // 배출시설 / 시료채취장소 / 시료번호
            t.Append(RowNoVert(
                LblCell("배출시설:", 1),  ValCell(d.배출시설, 1),
                LblCell("시료채취장소:", 1), ValCell(d.시료채취장소, 1),
                LblCell("시료 번호:", 1), ValCell(d.시료번호, 1)));

            // 채취일자 (span 6)
            t.Append(RowNoVert(LblCell("채취일자", 1), ValCell(d.채취일자, 5)));

            // 채취방법
            t.Append(RowNoVert(
                LblCell("채취방법:", 1),
                ValCell("☐ Grab", 2),
                ValCell("☐ Composite", 3)));

            // 채취시간 — 헤더
            t.Append(RowNoVert(SectionHeaderCell("채취시간", 6)));
            // 채취시간 본문 — 일/시/분 사이 비파괴 공백으로 수기 입력 공간 확보
            const string nbsp9 = "         "; // 9 NBSP
            string timeSlots = $"{nbsp9}일{nbsp9}시{nbsp9}분{nbsp9}";
            t.Append(RowNoVert(
                LblCell("시작:", 1), ValCell(timeSlots, 2),
                LblCell("종료:", 1), ValCell(timeSlots, 2)));

            // 시료채취 1일 전의 폐수 특성 (span 6)
            t.Append(RowNoVert(
                LblCell("시료채취 1일 전의 폐수 특성:", 2),
                ValCell("☐ 정상    ☐ 비정상(내용:                          )", 4)));

            // 시료 운반 조건
            t.Append(RowNoVert(
                LblCell("시료 운반 조건:", 2),
                ValCell("☐ 냉장    ☐ 실온", 4)));

            // 시료채취량 / 유효기간
            t.Append(RowNoVert(
                LblCell("시료채취량(L):", 1), ValCell("", 2),
                LblCell("시료유효기간:", 1), ValCell("", 2)));

            // ── 현장 측정항목 ──────────────────────────────────────────────
            t.Append(RowNoVert(SectionHeaderCell("현장 측정항목", 6)));
            t.Append(RowNoVert(
                LblCell("온도(℃):", 1), ValCell(d.온도, 1),
                LblCell("pH:", 1),       ValCell(d.pH, 1),
                LblCell("시료색깔:", 1), ValCell("☐ 유  ☐ 무(      )", 1)));
            t.Append(RowNoVert(
                LblCell("용존산소(mg/L):", 1), ValCell(d.용존산소, 2),
                LblCell("유입수량(m³/일):", 1), ValCell(d.유입수량, 2)));
            t.Append(RowNoVert(
                LblCell("전기전도도(μS/cm):", 1), ValCell(d.전기전도도, 2),
                LblCell("염분(‰):", 1), ValCell(d.염분, 2)));

            // ── 수질 측정항목 ──────────────────────────────────────────────
            // 라벨이 길어 줄바꿈되는 경우가 있어 라벨만 7pt로 축소
            t.Append(RowNoVert(SectionHeaderCell("수질 측정항목", 6)));
            t.Append(RowNoVert(
                LblCell("잔류염소(mg/L):", 1, ptOverride: 7),     ValCell(d.잔류염소, 1),
                LblCell("암모니아(NH₃, mg/L):", 1, ptOverride: 7), ValCell(d.암모니아, 1),
                LblCell("경도(mg/L):", 1, ptOverride: 7),         ValCell(d.경도, 1)));

            // ── 독성시험결과 ──────────────────────────────────────────────
            t.Append(RowNoVert(SectionHeaderCell("독성시험결과", 6)));
            // 헤더 행: 농도(%) | 반복횟수(1~4 — span 4) | 유영저해율(%)
            t.Append(Row(
                HdrCell("농도(%) / 반복횟수", 1),
                HdrCell("1", 1), HdrCell("2", 1), HdrCell("3", 1), HdrCell("4", 1),
                HdrCell("유영저해율(%)", 1)));
            // 데이터 행 6개 — 매트릭스라 세로선 유지.
            // 농도별 1~4 반복 + 유영저해율은 인쇄 후 수기 입력 받도록 빈 셀로 둠.
            string[] rowLabels = { "Control", "6.25", "12.5", "25", "50", "100" };
            for (int r = 0; r < 6; r++)
            {
                t.Append(Row(
                    LblCell(rowLabels[r], 1, align: JustificationValues.Center),
                    ValCell("", 1),
                    ValCell("", 1),
                    ValCell("", 1),
                    ValCell("", 1),
                    ValCell("", 1)));
            }

            // ── 독성시험항목 ──────────────────────────────────────────────
            t.Append(RowNoVert(SectionHeaderCell("독성시험항목", 6)));
            t.Append(RowNoVert(
                LblCell("물벼룩: 24시간 급성독성", 2),
                LblCell("EC₅₀ 값:", 1), ValCell(d.EC50, 1),
                LblCell("TU(100/EC₅₀ 값):", 1), ValCell(d.TU, 1)));
            t.Append(RowNoVert(
                LblCell("사용한 통계분석법:", 2),
                ValCell(d.통계분석법, 4)));

            // 최종 작성일
            string yy = "", mm = "", dd = "";
            if (DateTime.TryParse(d.작성일, out var date))
            {
                yy = date.Year.ToString();
                mm = date.Month.ToString();
                dd = date.Day.ToString();
            }
            t.Append(RowNoVert(
                LblCell("최종 작성일:", 2),
                ValCell($"{yy}", 1, align: JustificationValues.Center),
                LblCell("년", 1, align: JustificationValues.Left),
                ValCell($"{mm}", 1, align: JustificationValues.Center),
                LblCell("월", 1, align: JustificationValues.Left)));

            body.Append(t);

            // ── 기입자 서명란 ──────────────────────────────────────────────
            body.Append(P("", BodyPt, spacingAfter: 200));
            // "기입자 성명 :    {이름}                                      (서명)"
            // 비파괴 공백( )을 사용해 Word 가 단일 공백으로 축약하지 않도록 한다.
            string nbsp = " ";
            string padBeforeName = string.Concat(System.Linq.Enumerable.Repeat(nbsp, 4));
            string padAfterName  = string.Concat(System.Linq.Enumerable.Repeat(nbsp, 30));
            body.Append(P($"기입자 성명 :{padBeforeName}{d.기입자성명}{padAfterName}(서명)",
                BodyPt, align: JustificationValues.Center));

            body.Append(sectPr);
            main.Document.Save();
            return (true, "");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    // ─── 헬퍼 ────────────────────────────────────────────────────────────────
    private static string Get(string[] arr, int i) => i < arr.Length ? (arr[i] ?? "") : "";

    private static Paragraph P(string text, int pt,
        bool bold = false,
        JustificationValues? align = null,
        int spacingBefore = 0, int spacingAfter = 0)
    {
        var rp = new RunProperties(
            new RunFonts { Ascii = Font, EastAsia = Font, HighAnsi = Font },
            new FontSize { Val = (pt * 2).ToString() });
        if (bold) rp.Append(new Bold());

        var pp = new ParagraphProperties();
        if (align != null) pp.Append(new Justification { Val = align.Value });
        if (spacingBefore > 0 || spacingAfter > 0)
            pp.Append(new SpacingBetweenLines { Before = spacingBefore.ToString(), After = spacingAfter.ToString() });

        return new Paragraph(pp, new Run(rp, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static TableProperties StdTableProps() => new TableProperties(
        new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },  // 100% (50.00%/100 단위)
        new TableBorders(
            new TopBorder    { Val = BorderValues.Single, Size = 6 },
            new LeftBorder   { Val = BorderValues.Single, Size = 6 },
            new BottomBorder { Val = BorderValues.Single, Size = 6 },
            new RightBorder  { Val = BorderValues.Single, Size = 6 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4 }),
        new TableLayout { Type = TableLayoutValues.Fixed });

    private static TableRow Row(params TableCell[] cells)
    {
        var r = new TableRow();
        foreach (var c in cells) r.Append(c);
        return r;
    }

    /// <summary>내부 세로선 없는 행 — 마지막 셀 제외 모든 셀의 Right border 를 None 으로 override.</summary>
    private static TableRow RowNoVert(params TableCell[] cells)
    {
        var r = new TableRow();
        for (int i = 0; i < cells.Length; i++)
        {
            if (i < cells.Length - 1)
            {
                var props = cells[i].GetFirstChild<TableCellProperties>() ?? new TableCellProperties();
                if (cells[i].GetFirstChild<TableCellProperties>() == null) cells[i].PrependChild(props);
                props.Append(new TableCellBorders(
                    new RightBorder { Val = BorderValues.None }));
            }
            r.Append(cells[i]);
        }
        return r;
    }

    private static TableCell LblCell(string text, int span,
        JustificationValues? align = null, int? ptOverride = null)
        => MakeCell(text, span, bold: false, fillHex: "F2F2F2",
            align: align ?? JustificationValues.Left, ptOverride: ptOverride);

    private static TableCell ValCell(string text, int span,
        JustificationValues? align = null)
        => MakeCell(text, span, bold: false, fillHex: null, align: align ?? JustificationValues.Center);

    private static TableCell HdrCell(string text, int span)
        => MakeCell(text, span, bold: true, fillHex: "DDE5F0", align: JustificationValues.Center, ptOverride: HdrPt);

    private static TableCell SectionHeaderCell(string text, int span)
        => MakeCell(text, span, bold: true, fillHex: "B7CDE5", align: JustificationValues.Left, ptOverride: HdrPt);

    private static TableCell TitleCell(string text, int span)
        => MakeCell(text, span, bold: true, fillHex: null, align: JustificationValues.Center, ptOverride: TitlePt);

    private static TableCell MakeCell(string text, int span, bool bold, string? fillHex,
        JustificationValues align, int? ptOverride = null)
    {
        var props = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "0" });
        if (span > 1) props.Append(new GridSpan { Val = span });
        if (!string.IsNullOrEmpty(fillHex))
            props.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fillHex });
        props.Append(new TableCellMargin(
            new TopMargin    { Width = "60", Type = TableWidthUnitValues.Dxa },
            new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new LeftMargin   { Width = "100", Type = TableWidthUnitValues.Dxa },
            new RightMargin  { Width = "100", Type = TableWidthUnitValues.Dxa }));

        int pt = ptOverride ?? BodyPt;
        var rp = new RunProperties(
            new RunFonts { Ascii = Font, EastAsia = Font, HighAnsi = Font },
            new FontSize { Val = (pt * 2).ToString() });
        if (bold) rp.Append(new Bold());

        var pp = new ParagraphProperties(
            new Justification { Val = align },
            new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });

        var p = new Paragraph(pp, new Run(rp, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));
        return new TableCell(props, p);
    }
}
