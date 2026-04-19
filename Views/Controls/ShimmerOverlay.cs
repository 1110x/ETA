using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;

namespace ETA.Views.Controls;

/// <summary>
/// 재사용 shimmer sweep — wireframe v0.1 디자인 시스템.
/// CSS `shimmerSweep 2.2s` 애니메이션을 Avalonia Border로 이식.
/// DetachedFromVisualTree 시 타이머 자동 정리.
/// </summary>
public static class ShimmerOverlay
{
    /// <summary>주어진 Border에 shimmer sweep 배경을 덧씌운다 (배경색 위 하얀 빛띠가 좌→우로 흐름).</summary>
    public static void Attach(Border target, Color? sweepColor = null, double durationSec = 2.2)
    {
        var c = sweepColor ?? Color.FromArgb(220, 255, 255, 255);
        var transparent = Color.FromArgb(0, c.R, c.G, c.B);

        // 원본 배경 보존, 위에 덮을 Panel을 Child로 감싼다. 이미 Child가 있으면 Grid로 래핑.
        var sweep = new Border
        {
            IsHitTestVisible = false,
            Background = MakeGradient(transparent, c, -0.3),
        };

        if (target.Child is null)
        {
            target.Child = sweep;
        }
        else
        {
            var grid = new Grid();
            var existing = target.Child;
            target.Child = null;
            grid.Children.Add(existing);
            grid.Children.Add(sweep);
            target.Child = grid;
        }

        double offset = -0.3;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(30) };
        double step = 1.6 / (durationSec * 33.3); // 1 sweep = durationSec초
        timer.Tick += (_, _) =>
        {
            offset += step;
            if (offset > 1.3) offset = -0.3;
            sweep.Background = MakeGradient(transparent, c, offset);
        };
        timer.Start();

        target.DetachedFromVisualTree += (_, _) => timer.Stop();
    }

    private static LinearGradientBrush MakeGradient(Color transparent, Color c, double center)
    {
        double a = Math.Clamp(center - 0.12, 0, 1);
        double b = Math.Clamp(center,        0, 1);
        double d = Math.Clamp(center + 0.12, 0, 1);
        return new LinearGradientBrush
        {
            StartPoint = new RelativePoint(0, 0.5, RelativeUnit.Relative),
            EndPoint   = new RelativePoint(1, 0.5, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(transparent, 0.0),
                new GradientStop(transparent, a),
                new GradientStop(c,           b),
                new GradientStop(transparent, d),
                new GradientStop(transparent, 1.0),
            },
        };
    }
}
