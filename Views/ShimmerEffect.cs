using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views;

/// <summary>
/// 마우스 호버 시 텍스트 색상이 그라데이션으로 흐르는 효과.
/// </summary>
public static class TextShimmer
{
    private static readonly Dictionary<Control, (DispatcherTimer timer, TextBlock[] targets)> _active = new();

    public static void AttachHover(Control control)
    {
        // Loaded 후 TextBlock을 찾아야 하므로 이벤트에서 처리
        control.PointerEntered += (_, _) => StartHover(control);
        control.PointerExited  += (_, _) => StopHover(control);
    }

    private static void StartHover(Control control)
    {
        if (_active.ContainsKey(control)) return;

        var textBlocks = control.GetVisualDescendants().OfType<TextBlock>().ToArray();
        if (textBlocks.Length == 0 && control is TextBlock tb)
            textBlocks = [tb];
        if (textBlocks.Length == 0) return;

        // 현재 실제 색상을 기준으로 그라데이션 생성
        var brushes = new LinearGradientBrush[textBlocks.Length];
        for (int i = 0; i < textBlocks.Length; i++)
        {
            var cur = (textBlocks[i].Foreground as ISolidColorBrush)?.Color ?? Color.Parse("#d8d8e0");
            brushes[i] = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0.5, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(1, 0.5, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(Colors.White, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 1.0),
                }
            };
            textBlocks[i].Foreground = brushes[i];
        }

        double offset = -0.3;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        timer.Tick += (_, _) =>
        {
            offset += 0.023;
            if (offset > 1.3) offset = -0.3;

            double a = Math.Clamp(offset - 0.15, 0, 1);
            double b = Math.Clamp(offset, 0, 1);
            double c = Math.Clamp(offset + 0.15, 0, 1);

            for (int i = 0; i < brushes.Length; i++)
            {
                var cur = brushes[i].GradientStops[0].Color;
                brushes[i].GradientStops[1] = new GradientStop(cur, a);
                brushes[i].GradientStops[2] = new GradientStop(Colors.White, b);
                brushes[i].GradientStops[3] = new GradientStop(cur, c);
            }
        };
        timer.Start();
        _active[control] = (timer, textBlocks);
    }

    private static void StopHover(Control control)
    {
        if (!_active.TryGetValue(control, out var state)) return;
        state.timer.Stop();

        // ClearValue로 스타일 상속 복원
        foreach (var tb in state.targets)
            tb.ClearValue(TextBlock.ForegroundProperty);

        _active.Remove(control);
    }

    // ── 패널 shimmer ──────────────────────────────────────────────────────
    // Show1/Show4처럼 내부 컨텐츠가 동적으로 교체되는 패널에 사용.
    // 호버 진입 시 현재 자식 TextBlock을 즉시 스캔해서 shimmer를 붙임.
    // 단, 너무 많은 TextBlock(DataGrid 셀 등)은 제외하고 상위 헤더 레벨만 적용.

    private static readonly Dictionary<Control, DispatcherTimer> _panelActive = new();

    public static void AttachPanelHover(Control panel)
    {
        panel.PointerEntered += (_, _) => StartPanelShimmer(panel);
        panel.PointerExited  += (_, _) => StopPanelShimmer(panel);
    }

    private static void StartPanelShimmer(Control panel)
    {
        if (_panelActive.ContainsKey(panel)) return;

        // 현재 시점의 TextBlock 최대 20개만 (헤더·레이블 수준)
        var textBlocks = panel.GetVisualDescendants()
            .OfType<TextBlock>()
            .Where(tb => tb.IsVisible && tb.Text?.Length > 0)
            .Take(20)
            .ToArray();
        if (textBlocks.Length == 0) return;

        var brushes = new LinearGradientBrush[textBlocks.Length];
        for (int i = 0; i < textBlocks.Length; i++)
        {
            var cur = (textBlocks[i].Foreground as ISolidColorBrush)?.Color ?? Color.Parse("#d8d8e0");
            brushes[i] = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0.5, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(1, 0.5, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(Colors.White, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 1.0),
                }
            };
            textBlocks[i].Foreground = brushes[i];
        }

        double offset = -0.3;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        timer.Tick += (_, _) =>
        {
            offset += 0.018; // 서브메뉴보다 약간 느리게
            if (offset > 1.3) offset = -0.3;

            double a = Math.Clamp(offset - 0.15, 0, 1);
            double b = Math.Clamp(offset, 0, 1);
            double c = Math.Clamp(offset + 0.15, 0, 1);

            for (int i = 0; i < brushes.Length; i++)
            {
                var cur = brushes[i].GradientStops[0].Color;
                brushes[i].GradientStops[1] = new GradientStop(cur, a);
                brushes[i].GradientStops[2] = new GradientStop(Colors.White, b);
                brushes[i].GradientStops[3] = new GradientStop(cur, c);
            }
        };
        timer.Start();
        _panelActive[panel] = timer;

        // 패널 호버 시 보더 글로우 효과 (TranslateTransform 없이 순수 색상만)
        _active[panel] = (timer, textBlocks);
    }

    private static void StopPanelShimmer(Control panel)
    {
        if (!_panelActive.TryGetValue(panel, out var timer)) return;
        timer.Stop();
        _panelActive.Remove(panel);

        if (_active.TryGetValue(panel, out var state))
        {
            foreach (var tb in state.targets)
                tb.ClearValue(TextBlock.ForegroundProperty);
            _active.Remove(panel);
        }
    }
}
