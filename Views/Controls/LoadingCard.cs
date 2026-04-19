using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Views;

namespace ETA.Views.Controls;

/// <summary>
/// wireframe v0.1 · 로딩/진행 카드.
///
/// 2가지 모드:
///   • Skeleton(title)              — 값 없음. shimmer sweep + skeleton bar.
///   • Live(title, value, progress) — 현재값 + ProgressBar (클래스 ok/warn/bad/info).
///
/// 사용:
///   var card = LoadingCard.Skeleton("SS");
///   var card = LoadingCard.Live("BOD", "4.2 mg/L", 1.0, BadgeStatus.Ok);
/// </summary>
public static class LoadingCard
{
    public static Border Skeleton(string title)
    {
        var titleBlock = new TextBlock
        {
            Text = title,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontLG,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        };

        var skelBar = new Border
        {
            Height = 22,
            CornerRadius = new CornerRadius(4),
            Background = AppTheme.BorderSubtle,
            Margin = new Thickness(0, 8, 0, 8),
        };

        var progress = new ProgressBar
        {
            Value = 60,
            IsIndeterminate = true,
            Height = 6,
            Classes = { "info" },
        };

        var stack = new StackPanel { Spacing = 0 };
        stack.Children.Add(titleBlock);
        stack.Children.Add(skelBar);
        stack.Children.Add(new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            Children =
            {
                StatusBadge.Info("입력중", AppTheme.FontXS),
                progress,
            },
        });

        var card = new Border
        {
            Padding = new Thickness(12),
            CornerRadius = new CornerRadius(6),
            Background = AppTheme.BgCard,
            BorderBrush = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            Child = stack,
        };

        // shimmer sweep 오버레이
        ShimmerOverlay.Attach(card);

        return card;
    }

    public static Border Live(string title, string value, double progress01, BadgeStatus status = BadgeStatus.Ok)
    {
        var titleBlock = new TextBlock
        {
            Text = title,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontLG,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        };

        var badge = status switch
        {
            BadgeStatus.Ok    => StatusBadge.Ok("✓",   AppTheme.FontXS, false),
            BadgeStatus.Warn  => StatusBadge.Warn("⚠", AppTheme.FontXS, false),
            BadgeStatus.Bad   => StatusBadge.Bad("✗",  AppTheme.FontXS, false),
            BadgeStatus.Info  => StatusBadge.Info("●●○", AppTheme.FontXS, false),
            _                 => StatusBadge.Muted("—", AppTheme.FontXS, false),
        };

        var head = new DockPanel();
        DockPanel.SetDock(badge, Dock.Right);
        head.Children.Add(badge);
        head.Children.Add(titleBlock);

        var valueBlock = new TextBlock
        {
            Text = value,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = 22,
            FontWeight = FontWeight.Bold,
            Margin = new Thickness(0, 4, 0, 6),
            Foreground = AppTheme.FgPrimary,
        };

        var progressCls = status switch
        {
            BadgeStatus.Ok   => "ok",
            BadgeStatus.Warn => "warn",
            BadgeStatus.Bad  => "bad",
            BadgeStatus.Info => "info",
            _                => "info",
        };

        var bar = new ProgressBar
        {
            Minimum = 0, Maximum = 1,
            Value = Math.Clamp(progress01, 0, 1),
            Height = 6,
        };
        bar.Classes.Add(progressCls);

        var stack = new StackPanel();
        stack.Children.Add(head);
        stack.Children.Add(valueBlock);
        stack.Children.Add(bar);

        return new Border
        {
            Padding = new Thickness(12),
            CornerRadius = new CornerRadius(6),
            Background = AppTheme.BgCard,
            BorderBrush = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            Child = stack,
        };
    }
}
