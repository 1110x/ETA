using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ETA.Models;

namespace ETA.Views;

/// <summary>
/// 폐수 분석결과 추이 선 그래프 (각 항목 독립 정규화)
/// </summary>
internal sealed class WasteChartControl : Control
{
    private readonly string _title;
    private readonly List<WasteAnalysisResult> _data;
    private readonly HashSet<string> _enabled;   // 표시할 항목

    internal static readonly (string Label, Color Color, Func<WasteAnalysisResult, double?> Get)[] Series =
    {
        ("BOD",        Color.Parse("#ff7755"), r => r.BOD),
        ("TOC(TC-IC)", Color.Parse("#5588ff"), r => r.TOC_TCIC),
        ("TOC(NPOC)",  Color.Parse("#44aaee"), r => r.TOC_NPOC),
        ("SS",         Color.Parse("#ffcc33"), r => r.SS),
        ("T-N",        Color.Parse("#55dd55"), r => r.TN),
        ("T-P",        Color.Parse("#dd55dd"), r => r.TP),
        ("Phenols",    Color.Parse("#55ddcc"), r => r.Phenols),
        ("N-Hexan",    Color.Parse("#ff9944"), r => r.NHexan),
    };

    public WasteChartControl(string title, List<WasteAnalysisResult> data, HashSet<string>? enabled = null)
    {
        _title   = title;
        _data    = data;
        _enabled = enabled ?? new HashSet<string>(Series.Select(s => s.Label));
        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Stretch;
        VerticalAlignment   = Avalonia.Layout.VerticalAlignment.Stretch;
    }

    /// <summary>항목 켜기/끄기 후 화면 갱신</summary>
    public void SetEnabled(string label, bool on)
    {
        if (on) _enabled.Add(label); else _enabled.Remove(label);
        InvalidateVisual();
    }

    protected override Size MeasureOverride(Size available)
    {
        double w = double.IsInfinity(available.Width)  ? 400 : available.Width;
        double h = double.IsInfinity(available.Height) ? 260 : available.Height;
        return new Size(w, Math.Max(h, 160));
    }

    public override void Render(DrawingContext ctx)
    {
        double w = Math.Max(Bounds.Width,  200);
        double h = Math.Max(Bounds.Height, 160);

        const double padL = 12, padR = 20, padT = 24, padB = 42;
        double plotW = w - padL - padR;
        double plotH = h - padT - padB;

        // ── 배경 ────────────────────────────────────────────────────────
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#12121c")), new Rect(0, 0, w, h));
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#1a1a2a")), new Rect(padL, padT, plotW, plotH));

        // ── 제목 ────────────────────────────────────────────────────────
        DrawText(ctx, _title, Typeface.Default, 10, new SolidColorBrush(Color.Parse("#8899bb")), padL, 4);

        int n = _data.Count;
        if (n < 2) { DrawText(ctx, "데이터 없음", Typeface.Default, 11, Brushes.Gray, padL + 8, padT + plotH / 2); return; }

        // ── 수평 그리드 ─────────────────────────────────────────────────
        var gridPen = new Pen(new SolidColorBrush(Color.Parse("#252535")), 1);
        for (int g = 1; g < 4; g++)
        {
            double y = padT + g * plotH / 4;
            ctx.DrawLine(gridPen, new Point(padL, y), new Point(padL + plotW, y));
        }

        // ── 각 항목 선 ──────────────────────────────────────────────────
        double legendX = padL;
        foreach (var (label, color, getValue) in Series)
        {
            if (!_enabled.Contains(label)) continue;   // 체크 해제된 항목 건너뜀

            var vals    = _data.Select(r => getValue(r)).ToList();
            var nonNull = vals.Where(v => v.HasValue).Select(v => v!.Value).ToList();
            if (nonNull.Count < 2) continue;

            double maxV = nonNull.Max(), minV = nonNull.Min();
            if (Math.Abs(maxV - minV) < 1e-9) { minV -= 0.5; maxV += 0.5; }

            var pen = new Pen(new SolidColorBrush(color), 1.5);
            Point? prev = null;
            for (int i = 0; i < n; i++)
            {
                if (!vals[i].HasValue) { prev = null; continue; }
                double x = padL + (double)i / (n - 1) * plotW;
                double y = padT + plotH * (1.0 - (vals[i]!.Value - minV) / (maxV - minV));
                var pt = new Point(x, y);
                if (prev.HasValue) ctx.DrawLine(pen, prev.Value, pt);
                prev = pt;
            }

            // 범례 (활성 항목만)
            var lb = new SolidColorBrush(color);
            ctx.FillRectangle(lb, new Rect(legendX, h - 13, 8, 8));
            legendX += DrawText(ctx, label, Typeface.Default, 9, lb, legendX + 10, h - 14) + 18;
        }

        // ── X 축 날짜 레이블 ─────────────────────────────────────────────
        int step = Math.Max(1, n / 8);
        for (int i = 0; i < n; i += step)
        {
            double x   = padL + (double)i / (n - 1) * plotW;
            var    lbl = _data[i].채수일.Length >= 10 ? _data[i].채수일[5..10] : _data[i].채수일;
            var    ft  = new FormattedText(lbl, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                             Typeface.Default, 9, Brushes.DimGray);
            ctx.DrawText(ft, new Point(x - ft.Width / 2, padT + plotH + 3));
        }
    }

    private static double DrawText(DrawingContext ctx, string text, Typeface tf, double size,
                                   IBrush brush, double x, double y)
    {
        var ft = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                                   tf, size, brush);
        ctx.DrawText(ft, new Point(x, y));
        return ft.Width;
    }
}
