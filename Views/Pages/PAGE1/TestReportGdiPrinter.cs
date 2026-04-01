using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using ETA.Models;

namespace ETA.Views.Pages.PAGE1;

/// <summary>시험성적서를 GDI+ 로 직접 프린터 출력</summary>
public static class TestReportGdiPrinter
{
    // ── 페이지 여백 (mm) ── 템플릿 pageMargins 기준
    private const float ML  = 18.0f;   // 0.70866" × 25.4
    private const float MR  = 18.0f;
    private const float MT  = 27.0f;   // 1.06299" × 25.4
    private const float MB  = 24.0f;   // 0.94488" × 25.4
    private const float MH  = 16.0f;   // header margin 0.62992" × 25.4
    private const float MF  = 18.0f;   // footer margin

    private const float PH  = 297.0f;
    private const float UW  = 174.0f;  // 210 - ML - MR

    // ── 행 높이 (mm) — 엑셀 pt × (25.4/72) × 0.59 스케일
    private const float RH   = 4.68f;   // 22.5pt × 0.5864
    private const float RH42 = 4.52f;   // 21.75pt
    private const float RH44 = 10.44f;  // 50.25pt

    // ── 헤더 구역 X 좌표 (mm, 페이지 절대)
    // A+B | C+D | E | F+G+H
    private const float HX0 = ML;             // 18
    private const float HX1 = ML + 24.74f;    // 42.74  (A+B)
    private const float HX2 = ML + 86.73f;    // 104.73 (C+D 끝 = E 시작)
    private const float HX3 = ML + 113.88f;   // 131.88 (E 끝 = F 시작)
    private const float HX4 = ML + UW;        // 192

    // ── 데이터 구역 X 좌표 (mm)
    // A | B+C | D | E | F | G | H
    private const float DX0 = ML;
    private const float DX1 = ML + 10.22f;
    private const float DX2 = ML + 39.29f;
    private const float DX3 = ML + 86.73f;
    private const float DX4 = ML + 113.88f;
    private const float DX5 = ML + 130.81f;
    private const float DX6 = ML + 143.11f;
    private const float DX7 = ML + UW;

    // ── 색상 (템플릿: 배경색 없음 → 흰 바탕, 헤더만 아주 옅은 회색)
    private static readonly Color BgWhite  = Color.White;
    private static readonly Color BgLabel  = Color.FromArgb(242, 242, 242); // 아주 연한 회색
    private static readonly Color BgColHdr = Color.FromArgb(220, 220, 220);
    private static readonly Color ClrBord  = Color.FromArgb(140, 140, 140);
    private static readonly Color ClrBordH = Color.Black;

    // ═══════════════════════════════════════════════════════════════════════
    public static void Print(
        SampleRequest           sample,
        List<AnalysisResultRow> rows,
        Dictionary<string, string> stdMap,
        string reportNo,
        string qualityMgr,
        string companyName,
        string representative,
        bool   isQC)
    {
        int total  = Math.Max(1, (int)Math.Ceiling(rows.Count / 32.0));
        int pageNo = 0;

        var pd = new PrintDocument();
        pd.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);
        pd.DefaultPageSettings.Landscape = false;
        pd.DefaultPageSettings.Margins   = new Margins(0, 0, 0, 0);

        using var dlg = new System.Windows.Forms.PrintDialog
        {
            Document        = pd,
            UseEXDialog     = true,
            AllowSelection  = false,
            AllowCurrentPage = false,
            AllowSomePages  = false,
        };
        if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

        pd.PrintPage += (_, e) =>
        {
            var g = e.Graphics!;
            g.PageUnit          = GraphicsUnit.Millimeter;
            g.SmoothingMode     = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            DrawPage(g, sample, rows.Skip(pageNo * 32).Take(32).ToList(),
                     stdMap, reportNo, qualityMgr, companyName, representative,
                     isQC, pageNo, total);
            pageNo++;
            e.HasMorePages = pageNo < total;
        };
        pd.Print();
    }

    // ═══════════════════════════════════════════════════════════════════════
    private static void DrawPage(
        Graphics g,
        SampleRequest sample, List<AnalysisResultRow> rows,
        Dictionary<string, string> stdMap,
        string reportNo, string qualityMgr,
        string companyName, string representative,
        bool isQC, int pageIdx, int totalPages)
    {
        string no     = string.IsNullOrEmpty(reportNo)
            ? $"WAC-{DateTime.Now:yyyyMMdd}-{sample.약칭}" : reportNo;
        string suffix = pageIdx == 0 ? "-A" : $"-A-{pageIdx + 1}";

        // ── 페이지 헤더: "ANALYSIS REPORT" ─────────────────────────────
        DrawPageHeader(g, totalPages > 1 ? $"ANALYSIS REPORT  ({pageIdx + 1}/{totalPages})" : "ANALYSIS REPORT");

        // ── 페이지 푸터 ─────────────────────────────────────────────────
        DrawPageFooter(g, "리뉴어스주식회사 - 수질분석센터");

        float y = MT;

        // ── Row 1: 성적서 번호 ──────────────────────────────────────────
        DrawCell(g, HX0, y, HX1-HX0, RH, "성 적 서  번 호", lbl:true, a:TA.Center);
        DrawCell(g, HX1, y, HX2-HX1, RH, no + suffix,       lbl:false, a:TA.Left);
        DrawCell(g, HX2, y, HX4-HX2, RH, "",                lbl:false, a:TA.Left, noBorder:true);
        y += RH;

        // ── Row 2: 빈 행 ────────────────────────────────────────────────
        y += RH;

        // ── Row 3: 업체명 / 채수일자 ────────────────────────────────────
        DrawInfoRow(g, y, "업 체 명",
            string.IsNullOrEmpty(companyName) ? sample.의뢰사업장 : companyName,
            "채 수 일 자", FormatDate(sample.채취일자));
        y += RH;

        // ── Row 4: 대표자 / 채수담당자 ──────────────────────────────────
        string sampler = $"{sample.시료채취자1} {sample.시료채취자2}".Trim();
        DrawInfoRow(g, y, "대 표 자", representative, "채수담당자", sampler);
        y += RH;

        // ── Row 5: 채수입회자 / 분석완료일 ──────────────────────────────
        DrawInfoRow(g, y, "채수입회자", sample.입회자, "분석완료일", FormatDate(sample.분석종료일));
        y += RH;

        // ── Row 6: 빈 행 ────────────────────────────────────────────────
        y += RH;

        // ── Row 7: 시료명 / 의뢰정보 ────────────────────────────────────
        DrawInfoRow(g, y, "시 료 명", sample.시료명,
            "의뢰정보(용도)", isQC ? "정도보증 적용" : "참고용");
        y += RH;

        // ── Row 8: 시험결과 배너 ─────────────────────────────────────────
        DrawCell(g, HX0, y, HX1-HX0, RH, "시험결과", lbl:true, a:TA.Center);
        using (var p = new Pen(ClrBord, 0.15f))
            g.DrawLine(p, HX1, y+RH, HX4, y+RH);
        y += RH;

        // ── Row 9: 컬럼 헤더 ─────────────────────────────────────────────
        // 상단 굵은 선
        using (var p = new Pen(ClrBordH, 0.4f))
            g.DrawLine(p, DX0, y, DX7, y);
        DrawHdrCell(g, DX0, y, DX1-DX0, RH, "번호");
        DrawHdrCell(g, DX1, y, DX2-DX1, RH, "항목 구분");
        DrawHdrCell(g, DX2, y, DX3-DX2, RH, "시험 항목");
        DrawHdrCell(g, DX3, y, DX4-DX3, RH, "적용 시험방법");
        DrawHdrCell(g, DX4, y, DX5-DX4, RH, "결과");
        DrawHdrCell(g, DX5, y, DX6-DX5, RH, "단위");
        DrawHdrCell(g, DX6, y, DX7-DX6, RH, "비고");
        // 하단 이중선
        using (var p = new Pen(ClrBordH, 0.4f))
        {
            g.DrawLine(p, DX0, y+RH,       DX7, y+RH);
            g.DrawLine(p, DX0, y+RH+0.5f,  DX7, y+RH+0.5f);
        }
        y += RH;

        // ── Rows 10~41: 데이터 32행 ─────────────────────────────────────
        for (int i = 0; i < 32; i++)
        {
            if (i < rows.Count)
            {
                var r = rows[i];
                DrawDataRow(g, y, pageIdx*32+i+1,
                    r.Category ?? "", r.항목명 ?? "", r.ES ?? "",
                    FormatResult(r.결과값), r.단위 ?? "",
                    stdMap.GetValueOrDefault(r.항목명 ?? "", ""));
            }
            else
            {
                DrawEmptyRow(g, y);
            }
            y += RH;
        }

        // ── Row 42: 마감 이중선 ──────────────────────────────────────────
        using (var p = new Pen(ClrBordH, 0.4f))
        {
            g.DrawLine(p, DX0, y,       DX7, y);
            g.DrawLine(p, DX0, y+0.5f,  DX7, y+0.5f);
        }
        y += RH42;

        // ── Row 43: 서명 ─────────────────────────────────────────────────
        if (!string.IsNullOrEmpty(qualityMgr))
        {
            using var fn = F(8f);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString($"품질책임 수질분야 환경측정분석사       {qualityMgr}       (서명)",
                fn, Brushes.Black, new RectangleF(ML, y, UW, RH), sf);
        }
        y += RH;

        // ── Row 44: 면책고지 ─────────────────────────────────────────────
        string discl = isQC
            ? "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분석결과 입니다."
            : "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분석결과입니다.";
        using var fd = F(7f);
        var sfD = new StringFormat { LineAlignment = StringAlignment.Center };
        g.DrawString(discl, fd, Brushes.Black, new RectangleF(ML, y, UW, RH44), sfD);
    }

    // ── 페이지 헤더 / 푸터 ──────────────────────────────────────────────
    private static void DrawPageHeader(Graphics g, string text)
    {
        using var fn = F(22f, bold: true);
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        // 헤더 마진 기준 위치
        g.DrawString(text, fn, Brushes.Black,
            new RectangleF(ML, MH - 6f, UW, 12f), sf);
    }

    private static void DrawPageFooter(Graphics g, string text)
    {
        using var fn = F(16f, bold: true);
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(text, fn, Brushes.Black,
            new RectangleF(ML, PH - MF - 4f, UW, 8f), sf);
    }

    // ── 행 드로잉 ───────────────────────────────────────────────────────
    private static void DrawInfoRow(Graphics g, float y,
        string lbl1, string val1, string lbl2, string val2)
    {
        DrawCell(g, HX0, y, HX1-HX0, RH, lbl1, lbl:true,  a:TA.Center);
        DrawCell(g, HX1, y, HX2-HX1, RH, val1, lbl:false, a:TA.Left);
        DrawCell(g, HX2, y, HX3-HX2, RH, lbl2, lbl:true,  a:TA.Center);
        DrawCell(g, HX3, y, HX4-HX3, RH, val2, lbl:false, a:TA.Left);
    }

    private static void DrawHdrCell(Graphics g, float x, float y, float w, float h, string text)
        => DrawCell(g, x, y, w, h, text, lbl:true, a:TA.Center, bg:BgColHdr);

    private static void DrawDataRow(Graphics g, float y, int no,
        string cat, string item, string method,
        string result, string unit, string standard)
    {
        DrawCell(g, DX0, y, DX1-DX0, RH, no.ToString(), false, TA.Center);
        DrawCell(g, DX1, y, DX2-DX1, RH, cat,           false, TA.Left);
        DrawCell(g, DX2, y, DX3-DX2, RH, item,          false, TA.Left);
        DrawCell(g, DX3, y, DX4-DX3, RH, method,        false, TA.Center);
        DrawCell(g, DX4, y, DX5-DX4, RH, result,        false, TA.Right);
        DrawCell(g, DX5, y, DX6-DX5, RH, unit,          false, TA.Center);
        DrawCell(g, DX6, y, DX7-DX6, RH, standard,      false, TA.Center);
    }

    private static void DrawEmptyRow(Graphics g, float y)
    {
        float[] xs = { DX0, DX1, DX2, DX3, DX4, DX5, DX6, DX7 };
        for (int i = 0; i < 7; i++)
            DrawCell(g, xs[i], y, xs[i+1]-xs[i], RH, "", false, TA.Left);
    }

    // ── 셀 기본 ─────────────────────────────────────────────────────────
    private enum TA { Left, Center, Right }

    private static void DrawCell(
        Graphics g, float x, float y, float w, float h,
        string text, bool lbl, TA a,
        Color? bg = null, bool noBorder = false)
    {
        Color fill = bg ?? (lbl ? BgLabel : BgWhite);
        using var brush = new SolidBrush(fill);
        g.FillRectangle(brush, x, y, w, h);

        if (!noBorder)
        {
            using var pen = new Pen(ClrBord, 0.15f);
            g.DrawRectangle(pen, x, y, w, h);
        }

        if (string.IsNullOrEmpty(text)) return;

        using var font = F(lbl ? 7.5f : 7f);
        var sa = a == TA.Center ? StringAlignment.Center
               : a == TA.Right  ? StringAlignment.Far
               :                  StringAlignment.Near;
        var sf = new StringFormat
        {
            Alignment     = sa,
            LineAlignment = StringAlignment.Center,
            FormatFlags   = StringFormatFlags.NoWrap,
            Trimming      = StringTrimming.EllipsisCharacter,
        };
        g.DrawString(text, font, Brushes.Black,
            new RectangleF(x+0.5f, y+0.2f, w-1f, h-0.4f), sf);
    }

    // ── 폰트 ─────────────────────────────────────────────────────────────
    private static Font F(float pt, bool bold = false)
    {
        var style = bold ? FontStyle.Bold : FontStyle.Regular;
        try   { return new Font("KBIZ한마음고딕 M", pt, style, GraphicsUnit.Point); }
        catch { return new Font("맑은 고딕",         pt, style, GraphicsUnit.Point); }
    }

    // ── 날짜 / 결과 포맷 ─────────────────────────────────────────────────
    private static string FormatDate(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        return DateTime.TryParse(raw, out var dt) ? dt.ToString("yyyy-MM-dd") : raw;
    }

    private static string FormatResult(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        raw = raw.Trim();
        if (raw.StartsWith('<') || raw.StartsWith('>')) return raw;
        if (decimal.TryParse(raw,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
            return d.ToString("G6");
        return raw;
    }
}
