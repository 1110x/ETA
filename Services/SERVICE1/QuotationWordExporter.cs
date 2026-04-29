using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 견적서 Word(.docx) 출력 — 기존 Excel 템플릿 대안 프로토타입.
/// 시각언어: 시험기록부 Word 와 통일 (가로선 표 / 8pt 본문 / 9pt 섹션 / renewus 워터마크).
/// </summary>
public static class QuotationWordExporter
{
    private const string FontFamily = "맑은 고딕";
    private const int PageDxa        = 10400;
    private const int MinColWidth    = 700;
    private const int MaxColWidth    = 6000;
    private const int CellPaddingDxa = 200;
    private const int AsciiCharDxa   = 110;
    private const int CjkCharDxa     = 175;

    [ThreadStatic] private static SectionProperties? _sectPr;

    public static (bool ok, string msg) ExportWord(QuotationIssue issue, string savePath)
    {
        try
        {
            // 1) 데이터 수집 — Excel 출력과 동일한 로직
            var row = QuotationService.GetIssueRow(issue.Id);
            var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id","rowid","견적발행일자","업체명","약칭","대표자",
                "견적요청담당","담당자","담당자연락처","담당자 e-Mail",
                "시료명","견적번호","적용구분","적용구분_코드",
                "합계 금액","부가세","총합계","비고","거래명세서번호"
            };

            var items = new List<(string name, decimal qty, decimal unitPrice)>();
            foreach (var kv in row)
            {
                var col = kv.Key.Trim();
                if (fixedCols.Contains(col)) continue;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (!decimal.TryParse(kv.Value.Replace(",", ""), out var qty) || qty == 0) continue;
                decimal up = 0;
                if (row.TryGetValue(col + "단가", out var upStr))
                    decimal.TryParse(upStr.Replace(",", ""), out up);
                items.Add((col, qty, up));
            }
            // 2) 분석항목 메타 (카테고리 / ES) — 정렬에 ES 사용
            var meta = LoadMeta();

            // 분석항목은 ES 번호 오름차순 (분석정보에 없으면 끝으로, 동일 ES면 항목명)
            items.Sort((a, b) =>
            {
                string esA = meta.TryGetValue(a.name.Trim(), out var ma) ? ma.es ?? "" : "";
                string esB = meta.TryGetValue(b.name.Trim(), out var mb) ? mb.es ?? "" : "";
                bool aEmpty = string.IsNullOrWhiteSpace(esA);
                bool bEmpty = string.IsNullOrWhiteSpace(esB);
                if (aEmpty != bEmpty) return aEmpty ? 1 : -1;
                int cmp = string.Compare(esA, esB, StringComparison.OrdinalIgnoreCase);
                return cmp != 0 ? cmp : string.Compare(a.name, b.name, StringComparison.OrdinalIgnoreCase);
            });

            // 3) 받는이/발행자 정보 보완
            string 담당자 = !string.IsNullOrWhiteSpace(issue.담당자) ? issue.담당자
                            : (row.TryGetValue("담당자", out var v1) ? v1 : "");
            string 연락처 = !string.IsNullOrWhiteSpace(issue.담당자연락처) ? issue.담당자연락처
                            : (row.TryGetValue("담당자연락처", out var v2) ? v2 : "");
            string 이메일 = !string.IsNullOrWhiteSpace(issue.담당자이메일) ? issue.담당자이메일
                            : (row.TryGetValue("담당자 e-Mail", out var v3) ? v3
                            :  (row.TryGetValue("담당자e-Mail", out var v4) ? v4 : ""));
            string 대표자 = ContractService.GetRepresentativeByCompany(issue.업체명 ?? "");

            string 견적담당자성명 = "", 견적담당자이메일 = "";
            try
            {
                var uid = CurrentUserManager.Instance.CurrentUserId;
                if (!string.IsNullOrWhiteSpace(uid))
                {
                    using var c = DbConnectionFactory.CreateConnection();
                    c.Open();
                    using var cmd = c.CreateCommand();
                    cmd.CommandText = "SELECT COALESCE(`성명`,''), COALESCE(`Email`,'') FROM `Agent` WHERE `사번`=@id LIMIT 1";
                    cmd.Parameters.AddWithValue("@id", uid);
                    using var rdr = cmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        견적담당자성명   = rdr.IsDBNull(0) ? "" : rdr.GetString(0);
                        견적담당자이메일 = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
                    }
                }
            }
            catch { }

            // 4) Word 빌드
            using (var doc = WordprocessingDocument.Create(savePath, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                var body = main.Document.Body!;

                BuildStyles(main);
                BuildHeaderWatermark(main);
                BuildFooterPageNumber(main);

                body.Append(BuildTitle());

                Section(body, "받는이 / 발행자");
                body.Append(BuildHeaderInfo(issue, 대표자, 담당자, 연락처, 이메일,
                    견적담당자성명, 견적담당자이메일));
                body.Append(Spacer());

                Section(body, "견적 항목");
                decimal supply = items.Sum(it => it.qty * it.unitPrice);
                body.Append(BuildItemsTable(items, meta));
                body.Append(Spacer());

                if (row.TryGetValue("비고", out var memo) && !string.IsNullOrWhiteSpace(memo))
                {
                    Section(body, "비고");
                    body.Append(P(memo, 9));
                    body.Append(Spacer());
                }

                // 금액합계 — 페이지 최하단 anchor (TablePositionProperties)
                decimal vat = Math.Round(supply * 0.1m);
                decimal total = supply + vat;
                body.Append(BuildBottomTotals(supply, vat, total));

                if (_sectPr != null) { body.Append(_sectPr); _sectPr = null; }
                main.Document.Save();
            }

            return (true, "");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    // ─── 데이터 ──────────────────────────────────────────────────────────────
    /// <summary>분석정보 테이블에서 Analyte → (Category, ES) 맵 — Excel 출력과 동일.</summary>
    private static Dictionary<string, (string category, string es)> LoadMeta()
    {
        var m = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `Analyte`, COALESCE(`Category`,''), COALESCE(`ES`,'') FROM `분석정보`";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var a = rdr.IsDBNull(0) ? "" : rdr.GetString(0).Trim();
                if (string.IsNullOrWhiteSpace(a)) continue;
                var c  = rdr.IsDBNull(1) ? "" : rdr.GetString(1).Trim();
                var es = rdr.IsDBNull(2) ? "" : rdr.GetString(2).Trim();
                if (!m.ContainsKey(a)) m[a] = (c, es);
            }
        }
        catch { }
        return m;
    }

    // ─── 페이지 / 헤더 / 워터마크 ────────────────────────────────────────────
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
                        new SpacingBetweenLines { After = "0", Before = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }))));
        sp.Styles.Save();
    }

    private static void BuildFooterPageNumber(MainDocumentPart main)
    {
        var fp = main.AddNewPart<FooterPart>();
        string xml =
@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<w:ftr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:p>
    <w:pPr><w:jc w:val=""center""/><w:spacing w:before=""0"" w:after=""0""/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:fldChar w:fldCharType=""begin""/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:instrText xml:space=""preserve"">PAGE</w:instrText></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:fldChar w:fldCharType=""end""/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:t xml:space=""preserve"">/</w:t></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:fldChar w:fldCharType=""begin""/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:instrText xml:space=""preserve"">NUMPAGES</w:instrText></w:r>
    <w:r><w:rPr><w:rFonts w:ascii=""맑은 고딕"" w:eastAsia=""맑은 고딕"" w:hAnsi=""맑은 고딕""/><w:sz w:val=""18""/></w:rPr><w:fldChar w:fldCharType=""end""/></w:r>
  </w:p>
</w:ftr>";
        using (var s = fp.GetStream(FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(s, new System.Text.UTF8Encoding(false))) sw.Write(xml);

        // SectionProperties 가 BuildHeaderWatermark 에서 이미 생성됨 — FooterReference 만 추가
        _sectPr?.PrependChild(new FooterReference
        {
            Type = HeaderFooterValues.Default,
            Id   = main.GetIdOfPart(fp),
        });
    }

    private static void BuildHeaderWatermark(MainDocumentPart main)
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
            new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720, Header = 360, Footer = 360, Gutter = 0 });
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

    // ─── 본문 빌더 ──────────────────────────────────────────────────────────
    private static Paragraph BuildTitle() =>
        new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "120", After = "60" }),
            new Run(new RunProperties(new Bold(), new FontSize { Val = "44" }),
                new Text("견 적 서")));

    /// <summary>1행 표 — 좁은 파란 셀(액센트바) + 텍스트 셀. 막대 길이 = 텍스트 행 높이만큼.</summary>
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
            new Run(new RunProperties(new Bold(), new FontSize { Val = "18" }, new Color { Val = "1F4E79" }),
                new Text(title))));

        t.Append(new TableRow(
            new TableRowProperties(new TableRowHeight { Val = 200U, HeightType = HeightRuleValues.AtLeast }),
            barCell, txtCell));

        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "0" })));
        body.Append(t);
        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "30" })));
    }

    private static Paragraph Spacer() =>
        new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "20", After = "20" }));

    private static Paragraph P(string text, int pt) =>
        new Paragraph(new Run(
            new RunProperties(new FontSize { Val = (pt * 2).ToString() }),
            new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));

    private static Table BuildHeaderInfo(QuotationIssue issue, string 대표자,
        string 담당자, string 연락처, string 이메일, string 견적담당자, string 견적담당자이메일)
    {
        // 좌측(받는이) | 우측(발행자) — 두 컬럼으로 시각 분리
        var rows = new[] {
            new[] { ("업체명", issue.업체명 ?? ""), ("견적번호", issue.견적번호 ?? "") },
            new[] { ("대표자", 대표자), ("발행일", issue.발행일 ?? "") },
            new[] { ("담당자", 담당자), ("견적담당자", 견적담당자) },
            new[] { ("연락처", 연락처), ("견적담당자 E-mail", 견적담당자이메일) },
            new[] { ("E-mail", 이메일), ("시료명", issue.시료명 ?? "") },
        };
        return BuildPackedTable(rows, Align.Left);
    }

    private static Table BuildItemsTable(List<(string name, decimal qty, decimal unitPrice)> items,
        Dictionary<string, (string category, string es)> meta)
    {
        var headers = new[] { "No", "구분", "분석항목", "ES", "수량", "단가", "소계" };
        var rows = new List<string[]> { headers };
        for (int i = 0; i < items.Count; i++)
        {
            var (name, qty, up) = items[i];
            decimal sub = qty * up;
            string cat = "", es = "";
            if (meta.TryGetValue(name.Trim(), out var v)) { cat = v.category; es = v.es; }
            rows.Add(new[]
            {
                (i + 1).ToString(),
                cat,
                name,
                es,
                qty.ToString("0.##"),
                up.ToString("#,##0"),
                sub.ToString("#,##0"),
            });
        }

        int[] widths = ColWidths(rows.ToArray());
        var t = NewTable();
        AddGrid(t, widths);

        var hdr = new TableRow();
        for (int i = 0; i < headers.Length; i++)
            hdr.Append(Cell(headers[i], widths[i], isLabel: true, align: Align.Center));
        t.Append(hdr);

        for (int rIdx = 1; rIdx < rows.Count; rIdx++)
        {
            var tr = new TableRow();
            for (int c = 0; c < headers.Length; c++)
            {
                // 분석항목 좌측, 나머지 가운데(숫자 컬럼 포함 → 가독성 우선)
                Align align = c == 2 ? Align.Left : Align.Center;
                tr.Append(Cell(rows[rIdx][c], widths[c], align: align));
            }
            t.Append(tr);
        }
        return t;
    }

    /// <summary>금액 합계 표 — 페이지 최하단에 anchor (TablePositionProperties).
    /// 공급가액 / 부가세 / 총합계 3행, 라벨/값 페어 우측 정렬.</summary>
    private static Table BuildBottomTotals(decimal supply, decimal vat, decimal total)
    {
        var t = NewTable();
        var tp = t.GetFirstChild<TableProperties>();
        if (tp != null)
            tp.InsertAt(new TablePositionProperties
            {
                VerticalAnchor   = VerticalAnchorValues.Margin,
                HorizontalAnchor = HorizontalAnchorValues.Margin,
                TablePositionXAlignment = HorizontalAlignmentValues.Center,
                TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            }, 0);

        AddGrid(t, 4800, 5600);

        TableRow MakeRow(string label, string value, bool emphasize = false)
        {
            var lbl = Cell(label, 4800, isLabel: true, align: Align.Right);
            var val = Cell(value, 5600,
                isLabel: emphasize,
                align: Align.Right);
            return new TableRow(lbl, val);
        }

        t.Append(MakeRow("공급가액",      supply.ToString("#,##0") + " 원"));
        t.Append(MakeRow("부가세 (10%)",  vat.ToString("#,##0")    + " 원"));
        t.Append(MakeRow("총 합계",        total.ToString("#,##0")  + " 원", emphasize: true));
        return t;
    }

    // ─── 공통 셀/표 빌드 — EcotoxicityWordExporter 와 동일 스타일 ──────────
    private static Table BuildPackedTable((string label, string value)[][] rows, Align align)
    {
        int[] widths = PackedColWidths(rows);
        var t = NewTable();
        AddGrid(t, widths);
        foreach (var r in rows)
        {
            var tr = new TableRow();
            for (int i = 0; i < r.Length; i++)
            {
                int w = i < widths.Length ? widths[i] : MinColWidth;
                tr.Append(PackedCell(r[i].label, r[i].value, w, align));
            }
            t.Append(tr);
        }
        return t;
    }

    private static int[] PackedColWidths((string label, string value)[][] rows)
    {
        var asText = rows.Select(r => r.Select(c => $"{c.label}: {c.value}").ToArray()).ToArray();
        return ColWidths(asText);
    }

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
        if (sum == 0) { for (int i = 0; i < n; i++) widths[i] = PageDxa / n; return widths; }
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

    private static TableCell Cell(string text, int width, bool isLabel = false, Align align = Align.Left)
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

        var rp = new RunProperties();
        if (isLabel) rp.Append(new Bold());
        rp.Append(new FontSize { Val = "16" });

        return new TableCell(cp, new Paragraph(pp,
            new Run(rp, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static TableCell PackedCell(string label, string value, int width, Align align)
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
}
