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
    private static readonly Dictionary<Control, (DispatcherTimer timer, TextBlock[] targets, IBrush?[] origBrushes)> _active = new();

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
        var origBrushes = new IBrush?[textBlocks.Length];
        for (int i = 0; i < textBlocks.Length; i++)
        {
            origBrushes[i] = textBlocks[i].Foreground;  // 원본 저장
            var cur = (textBlocks[i].Foreground as ISolidColorBrush)?.Color ?? Color.Parse("#d8d8e0");
            brushes[i] = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0.5, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(1, 0.5, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(Color.Parse("#a8d0f0"), 0.0),
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
                brushes[i].GradientStops[2] = new GradientStop(Color.Parse("#a8d0f0"), b);
                brushes[i].GradientStops[3] = new GradientStop(cur, c);
            }
        };
        timer.Start();
        _active[control] = (timer, textBlocks, origBrushes);
    }

    private static void StopHover(Control control)
    {
        if (!_active.TryGetValue(control, out var state)) return;
        state.timer.Stop();

        // 원본 색상으로 복원 — 새 SolidColorBrush 생성 (그라데이션 참조 완전 제거)
        for (int i = 0; i < state.targets.Length; i++)
        {
            if (state.origBrushes[i] is ISolidColorBrush scb)
                state.targets[i].Foreground = new SolidColorBrush(scb.Color);
            else if (state.origBrushes[i] != null)
                state.targets[i].Foreground = state.origBrushes[i];
            else
                state.targets[i].ClearValue(TextBlock.ForegroundProperty);
        }

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
        var origBrushes = new IBrush?[textBlocks.Length];
        for (int i = 0; i < textBlocks.Length; i++)
        {
            origBrushes[i] = textBlocks[i].Foreground;  // 원본 저장
            var cur = (textBlocks[i].Foreground as ISolidColorBrush)?.Color ?? Color.Parse("#d8d8e0");
            brushes[i] = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0.5, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(1, 0.5, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(cur, 0.0),
                    new GradientStop(cur, 0.0),
                    new GradientStop(Color.Parse("#a8d0f0"), 0.0),
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
                brushes[i].GradientStops[2] = new GradientStop(Color.Parse("#a8d0f0"), b);
                brushes[i].GradientStops[3] = new GradientStop(cur, c);
            }
        };
        timer.Start();
        _panelActive[panel] = timer;

        // 패널 호버 시 보더 글로우 효과 (TranslateTransform 없이 순수 색상만)
        _active[panel] = (timer, textBlocks, origBrushes);
    }

    private static void StopPanelShimmer(Control panel)
    {
        if (!_panelActive.TryGetValue(panel, out var timer)) return;
        timer.Stop();
        _panelActive.Remove(panel);

        if (_active.TryGetValue(panel, out var state))
        {
            // 원본 색상으로 복원 — 새 SolidColorBrush 생성 (그라데이션 참조 완전 제거)
            for (int i = 0; i < state.targets.Length; i++)
            {
                if (state.origBrushes[i] is ISolidColorBrush scb)
                    state.targets[i].Foreground = new SolidColorBrush(scb.Color);
                else if (state.origBrushes[i] != null)
                    state.targets[i].Foreground = state.origBrushes[i];
                else
                    state.targets[i].ClearValue(TextBlock.ForegroundProperty);
            }
            _active.Remove(panel);
        }
    }

    // ── 전역 shimmer 부착 ────────────────────────────────────────────────
    // Window 전체 비주얼 트리를 순회하여 주요 컨트롤에 호버 shimmer를 자동 부착.
    // 이미 부착된 컨트롤은 중복 방지(_attached HashSet).

    private static readonly HashSet<Control> _attached = new();

    /// <summary>
    /// 윈도우 내 모든 주요 인터랙티브 컨트롤에 호버 shimmer를 자동 부착합니다.
    /// 테마 변경 시 호출해도 안전 (중복 부착 방지).
    /// </summary>
    public static void AttachAll(Control root)
    {
        foreach (var ctrl in root.GetVisualDescendants().OfType<Control>())
        {
            // 이미 등록된 컨트롤 스킵
            if (_attached.Contains(ctrl)) continue;

            bool attach = ctrl is Button
                       || ctrl is TreeViewItem
                       || ctrl is ListBoxItem
                       || ctrl is ComboBox
                       || ctrl is ComboBoxItem
                       || ctrl is MenuItem
                       || ctrl is TabItem
                       || ctrl is Expander;

            // 독립 TextBlock (부모가 Button/ListBoxItem 등이 아닌 경우)
            if (!attach && ctrl is TextBlock tb && tb.Text?.Length > 0)
            {
                var parent = tb.Parent;
                if (parent is not Button && parent is not TreeViewItem
                    && parent is not ListBoxItem && parent is not ComboBoxItem
                    && parent is not MenuItem && parent is not TabItem)
                {
                    attach = true;
                }
            }

            if (attach)
            {
                _attached.Add(ctrl);
                AttachHover(ctrl);
            }
        }
    }

    /// <summary>
    /// 새로 생성된 컨트롤(동적 UI)에 shimmer를 부착합니다.
    /// </summary>
    public static void AttachIfNew(Control ctrl)
    {
        if (_attached.Contains(ctrl)) return;
        _attached.Add(ctrl);
        AttachHover(ctrl);
    }
}
