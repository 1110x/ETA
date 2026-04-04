using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using ETA.Models;

namespace ETA.Views;

/// <summary>
/// 폐수 분석결과 — 단일 항목 Bar/Line 차트
/// </summary>
internal sealed class WasteSingleSeriesChart : Control
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private readonly string _label;
    private readonly Color _color;
    private readonly Func<WasteAnalysisResult, double?> _getValue;
    private readonly List<WasteAnalysisResult> _data;
    private bool _barMode = true;
    private int _hoverIndex = -1;

    // ── 바 등장 애니메이션 ──
    private double _animProgress = 0;
    private readonly Stopwatch _animSw = new();
    private DispatcherTimer? _animTimer;
    private const double AnimDurationMs = 600;

    public WasteSingleSeriesChart(
        string label, Color color,
        Func<WasteAnalysisResult, double?> getValue,
        List<WasteAnalysisResult> data)
    {
        _label    = label;
        _color    = color;
        _getValue = getValue;
        _data     = data;
        HorizontalAlignment = HorizontalAlignment.Stretch;
        VerticalAlignment   = VerticalAlignment.Stretch;
        StartBarAnimation();
    }

    private void StartBarAnimation()
    {
        _animProgress = 0;
        _animSw.Restart();
        _animTimer?.Stop();
        _animTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        _animTimer.Tick += (_, _) =>
        {
            double t = _animSw.ElapsedMilliseconds / AnimDurationMs;
            if (t >= 1.0) { _animProgress = 1.0; _animTimer.Stop(); }
            else { double u = 1.0 - t; _animProgress = 1.0 - u * u * u * u; } // QuarticEaseOut
            InvalidateVisual();
        };
        _animTimer.Start();
    }

    public void SetBarMode(bool bar) { _barMode = bar; InvalidateVisual(); }

    protected override Size MeasureOverride(Size available)
    {
        double w = double.IsInfinity(available.Width)  ? 300 : available.Width;
        double h = double.IsInfinity(available.Height) ? 150 : available.Height;
        return new Size(w, Math.Max(h, 100));
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (_data.Count == 0) return;
        var pos = e.GetPosition(this);
        double w = Math.Max(Bounds.Width, 100);
        const double padL = 44, padR = 8;
        double plotW = w - padL - padR;
        int n = _data.Count;
        if (n == 0 || plotW <= 0) return;
        double slotW = plotW / n;
        int idx = (int)((pos.X - padL) / slotW);
        idx = Math.Clamp(idx, -1, n - 1);
        if (idx != _hoverIndex) { _hoverIndex = idx; InvalidateVisual(); }
    }

    protected override void OnPointerExited(PointerEventArgs e)
    {
        base.OnPointerExited(e);
        if (_hoverIndex >= 0) { _hoverIndex = -1; InvalidateVisual(); }
    }

    public override void Render(DrawingContext ctx)
    {
        double w = Math.Max(Bounds.Width,  100);
        double h = Math.Max(Bounds.Height, 100);

        const double padL = 44, padR = 8, padT = 22, padB = 22;
        double plotW = w - padL - padR;
        double plotH = h - padT - padB;

        // 배경
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#12121c")), new Rect(0, 0, w, h));
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#1a1a2a")), new Rect(padL, padT, plotW, plotH));

        // 제목
        var titleBrush = new SolidColorBrush(_color);
        var titleFt = new FormattedText(_label, CultureInfo.InvariantCulture,
            FlowDirection.LeftToRight, new Typeface(Font), 10, titleBrush);
        ctx.DrawText(titleFt, new Point(padL, 4));

        int n = _data.Count;
        if (n == 0)
        {
            var emptyFt = new FormattedText("—", CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight, Typeface.Default, 10, Brushes.Gray);
            ctx.DrawText(emptyFt, new Point(padL + 8, padT + plotH / 2 - 5));
            return;
        }

        var vals = _data.Select(r => _getValue(r)).ToList();
        var nonNull = vals.Where(v => v.HasValue).Select(v => v!.Value).ToList();
        if (nonNull.Count == 0) return;

        double avg = nonNull.Average();
        double maxV, minV = 0;
        if (_barMode)
        {
            maxV = avg > 0 ? avg * 3.0 : nonNull.Max();
            if (maxV <= 0) maxV = 1;
        }
        else
        {
            maxV = nonNull.Max();
            minV = nonNull.Min();
            if (Math.Abs(maxV - minV) < 1e-9) { minV -= 0.5; maxV += 0.5; }
        }

        // Y축 눈금 (3줄)
        var gridPen = new Pen(new SolidColorBrush(Color.Parse("#252535")), 1);
        for (int g = 1; g <= 3; g++)
        {
            double gy = padT + g * plotH / 4;
            ctx.DrawLine(gridPen, new Point(padL, gy), new Point(padL + plotW, gy));
            double yVal = _barMode
                ? maxV * (1.0 - (double)g / 4)
                : minV + (maxV - minV) * (1.0 - (double)g / 4);
            var yFt = new FormattedText(FormatAxis(yVal), CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight, Typeface.Default, 8, Brushes.DimGray);
            ctx.DrawText(yFt, new Point(padL - yFt.Width - 3, gy - yFt.Height / 2));
        }

        // 상단(max), 하단(min/0) 눈금
        var maxFt = new FormattedText(FormatAxis(maxV), CultureInfo.InvariantCulture,
            FlowDirection.LeftToRight, Typeface.Default, 8, Brushes.DimGray);
        ctx.DrawText(maxFt, new Point(padL - maxFt.Width - 3, padT - maxFt.Height / 2));
        var minFt = new FormattedText(FormatAxis(minV), CultureInfo.InvariantCulture,
            FlowDirection.LeftToRight, Typeface.Default, 8, Brushes.DimGray);
        ctx.DrawText(minFt, new Point(padL - minFt.Width - 3, padT + plotH - minFt.Height / 2));

        double slotW = plotW / n;
        var brush = new SolidColorBrush(_color);
        var pen   = new Pen(brush, 1.8);

        if (_barMode)
        {
            double barW = Math.Min(slotW * 0.6, 24);
            var redColor = Color.Parse("#ee3333");
            for (int i = 0; i < n; i++)
            {
                if (!vals[i].HasValue) continue;
                double v = vals[i]!.Value;
                double ratio = v / maxV;
                ratio = Math.Clamp(ratio, 0, 1);
                double barH = ratio * plotH * _animProgress;
                double cx = padL + (i + 0.5) * slotW;
                double bx = cx - barW / 2;
                double by = padT + plotH - barH;
                byte alpha = (byte)(200 * _animProgress);
                bool isOutlier = avg > 0 && v > avg * 3.0;
                var c = isOutlier ? redColor : _color;
                var barBrush = new SolidColorBrush(Color.FromArgb(alpha, c.R, c.G, c.B));
                ctx.FillRectangle(barBrush, new Rect(bx, by, barW, barH));
            }
        }
        else
        {
            Point? prev = null;
            for (int i = 0; i < n; i++)
            {
                if (!vals[i].HasValue) { prev = null; continue; }
                double cx = padL + (i + 0.5) * slotW;
                double ratio = (vals[i]!.Value - minV) / (maxV - minV);
                double cy = padT + plotH * (1.0 - ratio);
                var pt = new Point(cx, cy);
                if (prev.HasValue) ctx.DrawLine(pen, prev.Value, pt);
                ctx.DrawEllipse(brush, null, pt, 3, 3);
                prev = pt;
            }
        }

        // X축 날짜
        int step = Math.Max(1, n / 6);
        for (int i = 0; i < n; i += step)
        {
            double cx  = padL + (i + 0.5) * slotW;
            string lbl = _data[i].채수일.Length >= 10 ? _data[i].채수일[5..10] : _data[i].채수일;
            var ft = new FormattedText(lbl, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                         Typeface.Default, 8, Brushes.DimGray);
            ctx.DrawText(ft, new Point(cx - ft.Width / 2, padT + plotH + 3));
        }

        // 호버 툴팁
        if (_hoverIndex >= 0 && _hoverIndex < n)
        {
            double hx = padL + _hoverIndex * slotW;
            ctx.FillRectangle(new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)),
                              new Rect(hx, padT, slotW, plotH));

            var d = _data[_hoverIndex];
            var v = _getValue(d);
            string tip = $"📅 {d.채수일}\n{_label}: {(v.HasValue ? FormatVal(v.Value) : "—")}";
            var tipFt = new FormattedText(tip, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                            Typeface.Default, 10, Brushes.White);
            double tipW = tipFt.Width + 14;
            double tipH = tipFt.Height + 8;
            double tipX = hx + slotW + 4;
            double tipY = padT + 4;
            if (tipX + tipW > w - 4) tipX = hx - tipW - 4;
            ctx.FillRectangle(new SolidColorBrush(Color.Parse("#DD222233")),
                              new Rect(tipX, tipY, tipW, tipH), 4);
            ctx.DrawText(tipFt, new Point(tipX + 7, tipY + 4));
        }
    }

    private static string FormatVal(double v) => v >= 100 ? v.ToString("F1") : v.ToString("F3");
    private static string FormatAxis(double v) => v >= 100 ? v.ToString("F0") : v >= 1 ? v.ToString("F1") : v.ToString("F2");
}
