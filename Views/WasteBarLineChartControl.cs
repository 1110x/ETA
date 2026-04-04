using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
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
/// 폐수 분석결과 Bar/Line 전환 차트 — 최근 N건, 항목별 독립 Y축 정규화
/// </summary>
internal sealed class WasteBarLineChartControl : Control
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    internal static readonly (string Label, Color Color, Func<WasteAnalysisResult, double?> Get)[] Series =
    {
        ("BOD",     Color.Parse("#ff7755"), r => r.BOD),
        ("TOC",     Color.Parse("#5588ff"), r => r.TOC_TCIC),
        ("SS",      Color.Parse("#ffcc33"), r => r.SS),
        ("T-N",     Color.Parse("#55dd55"), r => r.TN),
        ("T-P",     Color.Parse("#dd55dd"), r => r.TP),
        ("Phenols", Color.Parse("#55ddcc"), r => r.Phenols),
        ("N-Hexan", Color.Parse("#ff9944"), r => r.NHexan),
    };

    private readonly string _title;
    private readonly List<WasteAnalysisResult> _data;
    private readonly HashSet<string> _enabled;
    private bool _barMode = true;  // true=Bar, false=Line

    // 마우스 호버 — 툴팁용
    private int _hoverIndex = -1;

    // ── 바 등장 애니메이션 ──
    private double _animProgress = 0;
    private readonly Stopwatch _animSw = new();
    private DispatcherTimer? _animTimer;
    private const double AnimDurationMs = 600; // 0.6초

    public WasteBarLineChartControl(string title, List<WasteAnalysisResult> data)
    {
        _title   = title;
        _data    = data;
        _enabled = new HashSet<string>(Series.Select(s => s.Label));
        HorizontalAlignment = HorizontalAlignment.Stretch;
        VerticalAlignment   = VerticalAlignment.Stretch;
        StartBarAnimation();
    }

    private void StartBarAnimation()
    {
        _animProgress = 0;
        _animSw.Restart();
        _animTimer?.Stop();
        _animTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) }; // ~60fps
        _animTimer.Tick += (_, _) =>
        {
            double t = _animSw.ElapsedMilliseconds / AnimDurationMs;
            if (t >= 1.0)
            {
                _animProgress = 1.0;
                _animTimer.Stop();
            }
            else
            {
                // QuarticEaseOut: 빠르게 시작, 끝에 감속
                double u = 1.0 - t;
                _animProgress = 1.0 - u * u * u * u;
            }
            InvalidateVisual();
        };
        _animTimer.Start();
    }

    public void SetEnabled(string label, bool on)
    {
        if (on) _enabled.Add(label); else _enabled.Remove(label);
        InvalidateVisual();
    }

    public void SetBarMode(bool bar)
    {
        _barMode = bar;
        InvalidateVisual();
    }

    protected override Size MeasureOverride(Size available)
    {
        double w = double.IsInfinity(available.Width)  ? 500 : available.Width;
        double h = double.IsInfinity(available.Height) ? 300 : available.Height;
        return new Size(w, Math.Max(h, 180));
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (_data.Count == 0) return;

        var pos = e.GetPosition(this);
        double w = Math.Max(Bounds.Width, 200);
        const double padL = 50, padR = 16;
        double plotW = w - padL - padR;
        int n = _data.Count;
        if (n == 0 || plotW <= 0) return;

        double slotW = plotW / n;
        int idx = (int)((pos.X - padL) / slotW);
        idx = Math.Clamp(idx, -1, n - 1);

        if (idx != _hoverIndex)
        {
            _hoverIndex = idx;
            InvalidateVisual();
        }
    }

    protected override void OnPointerExited(PointerEventArgs e)
    {
        base.OnPointerExited(e);
        if (_hoverIndex >= 0) { _hoverIndex = -1; InvalidateVisual(); }
    }

    public override void Render(DrawingContext ctx)
    {
        double w = Math.Max(Bounds.Width,  200);
        double h = Math.Max(Bounds.Height, 180);

        const double padL = 50, padR = 16, padT = 8, padB = 48;
        double plotW = w - padL - padR;
        double plotH = h - padT - padB;

        // 배경
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#12121c")), new Rect(0, 0, w, h));
        ctx.FillRectangle(new SolidColorBrush(Color.Parse("#1a1a2a")), new Rect(padL, padT, plotW, plotH));

        int n = _data.Count;
        if (n == 0)
        {
            DrawText(ctx, "데이터 없음", Typeface.Default, 11, Brushes.Gray, padL + 8, padT + plotH / 2);
            return;
        }

        // 활성 시리즈 목록
        var active = Series.Where(s => _enabled.Contains(s.Label)).ToArray();
        if (active.Length == 0) return;

        // 수평 그리드
        var gridPen = new Pen(new SolidColorBrush(Color.Parse("#252535")), 1);
        for (int g = 1; g < 4; g++)
        {
            double y = padT + g * plotH / 4;
            ctx.DrawLine(gridPen, new Point(padL, y), new Point(padL + plotW, y));
        }

        // 슬롯 계산
        double slotW   = plotW / n;
        double barZone = slotW * 0.75;
        double barW    = active.Length > 0 ? barZone / active.Length : barZone;
        barW = Math.Min(barW, 28);

        // 각 시리즈 그리기
        int serIdx = 0;
        foreach (var (label, color, getValue) in active)
        {
            var vals    = _data.Select(r => getValue(r)).ToList();
            var nonNull = vals.Where(v => v.HasValue).Select(v => v!.Value).ToList();
            if (nonNull.Count == 0) { serIdx++; continue; }

            double maxV = nonNull.Max(), minV = 0;  // 바 차트는 0부터
            if (_barMode)
            {
                if (maxV <= 0) maxV = 1;
            }
            else
            {
                minV = nonNull.Min();
                if (Math.Abs(maxV - minV) < 1e-9) { minV -= 0.5; maxV += 0.5; }
            }

            var brush = new SolidColorBrush(color);
            var pen   = new Pen(brush, 1.8);

            if (_barMode)
            {
                // Bar 모드
                double totalBarWidth = barW * active.Length;
                double barStart = -totalBarWidth / 2 + serIdx * barW;

                for (int i = 0; i < n; i++)
                {
                    if (!vals[i].HasValue) continue;
                    double ratio = (vals[i]!.Value - minV) / (maxV - minV);
                    ratio = Math.Clamp(ratio, 0, 1);
                    double barH = ratio * plotH * _animProgress; // 애니메이션 적용
                    double cx   = padL + (i + 0.5) * slotW;
                    double bx   = cx + barStart;
                    double by   = padT + plotH - barH;

                    byte alpha = (byte)(200 * _animProgress);
                    var barBrush = new SolidColorBrush(Color.FromArgb(alpha, color.R, color.G, color.B));
                    ctx.FillRectangle(barBrush, new Rect(bx, by, Math.Max(barW - 1, 2), barH));
                }
            }
            else
            {
                // Line 모드
                Point? prev = null;
                for (int i = 0; i < n; i++)
                {
                    if (!vals[i].HasValue) { prev = null; continue; }
                    double cx = padL + (i + 0.5) * slotW;
                    double ratio = (vals[i]!.Value - minV) / (maxV - minV);
                    double cy = padT + plotH * (1.0 - ratio);
                    var pt = new Point(cx, cy);
                    if (prev.HasValue) ctx.DrawLine(pen, prev.Value, pt);
                    // 점
                    ctx.DrawEllipse(brush, null, pt, 3, 3);
                    prev = pt;
                }
            }
            serIdx++;
        }

        // X축 날짜 레이블
        int step = Math.Max(1, n / 10);
        for (int i = 0; i < n; i += step)
        {
            double cx  = padL + (i + 0.5) * slotW;
            string lbl = _data[i].채수일.Length >= 10 ? _data[i].채수일[5..10] : _data[i].채수일;
            var ft = new FormattedText(lbl, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                         Typeface.Default, 9, Brushes.DimGray);
            ctx.DrawText(ft, new Point(cx - ft.Width / 2, padT + plotH + 3));
        }

        // 범례 (하단)
        double legendX = padL;
        double legendY = h - 16;
        foreach (var (label, color, _) in active)
        {
            var lb = new SolidColorBrush(color);
            ctx.FillRectangle(lb, new Rect(legendX, legendY, 8, 8));
            legendX += DrawText(ctx, label, Typeface.Default, 9, lb, legendX + 10, legendY - 1) + 16;
        }

        // 호버 툴팁
        if (_hoverIndex >= 0 && _hoverIndex < n)
        {
            var d = _data[_hoverIndex];
            var lines = new List<string> { $"📅 {d.채수일}" };
            foreach (var (label, _, getValue) in active)
            {
                var v = getValue(d);
                lines.Add(v.HasValue ? $"{label}: {FormatVal(v.Value)}" : $"{label}: —");
            }
            string tip = string.Join("\n", lines);

            var tipFt = new FormattedText(tip, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                            Typeface.Default, 10, Brushes.White);
            double tipW = tipFt.Width + 16;
            double tipH = tipFt.Height + 10;
            double tipX = padL + (_hoverIndex + 0.5) * slotW + 10;
            double tipY = padT + 10;
            if (tipX + tipW > w - 4) tipX = tipX - tipW - 20;

            ctx.FillRectangle(new SolidColorBrush(Color.Parse("#DD222233")), new Rect(tipX, tipY, tipW, tipH), 4);
            ctx.DrawText(tipFt, new Point(tipX + 8, tipY + 5));

            // 호버 컬럼 하이라이트
            double hx = padL + _hoverIndex * slotW;
            ctx.FillRectangle(new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)),
                              new Rect(hx, padT, slotW, plotH));
        }
    }

    private static string FormatVal(double v) => v >= 100 ? v.ToString("F1") : v.ToString("F3");

    private static double DrawText(DrawingContext ctx, string text, Typeface tf, double size,
                                   IBrush brush, double x, double y)
    {
        var ft = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                                   tf, size, brush);
        ctx.DrawText(ft, new Point(x, y));
        return ft.Width;
    }
}
