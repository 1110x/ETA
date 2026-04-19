using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Views;

namespace ETA.Views.Controls;

/// <summary>
/// 통일 상태 배지 (wireframe v0.1 디자인 시스템).
/// 5종 상태(ok/warn/bad/info/muted) + accent 공통 렌더.
/// 모든 화면에서 뱃지 생성은 이 클래스를 통해서만.
/// </summary>
public enum BadgeStatus { Ok, Warn, Bad, Info, Muted, Accent }

public static class StatusBadge
{
    public static Border Create(string text, BadgeStatus status, double? fontSize = null, bool withIcon = true)
    {
        var (bg, fg, border) = GetBrushes(status);
        var prefix = withIcon ? GetPrefix(status) : "";
        return new Border
        {
            Background = bg,
            BorderBrush = border,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(999),
            Padding = new Thickness(8, 2),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = prefix + text,
                FontSize = fontSize ?? AppTheme.FontXS,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = fg,
                VerticalAlignment = VerticalAlignment.Center,
            },
        };
    }

    public static Border Ok(string text, double? fontSize = null, bool withIcon = true)    => Create(text, BadgeStatus.Ok,    fontSize, withIcon);
    public static Border Warn(string text, double? fontSize = null, bool withIcon = true)  => Create(text, BadgeStatus.Warn,  fontSize, withIcon);
    public static Border Bad(string text, double? fontSize = null, bool withIcon = true)   => Create(text, BadgeStatus.Bad,   fontSize, withIcon);
    public static Border Info(string text, double? fontSize = null, bool withIcon = true)  => Create(text, BadgeStatus.Info,  fontSize, withIcon);
    public static Border Muted(string text, double? fontSize = null, bool withIcon = true) => Create(text, BadgeStatus.Muted, fontSize, withIcon);
    public static Border Accent(string text, double? fontSize = null, bool withIcon = true)=> Create(text, BadgeStatus.Accent,fontSize, withIcon);

    public static (IBrush bg, IBrush fg, IBrush border) GetBrushes(BadgeStatus s) => s switch
    {
        BadgeStatus.Ok     => (AppTheme.StatusOkBg,     AppTheme.StatusOkFg,     AppTheme.StatusOkBorder),
        BadgeStatus.Warn   => (AppTheme.StatusWarnBg,   AppTheme.StatusWarnFg,   AppTheme.StatusWarnBorder),
        BadgeStatus.Bad    => (AppTheme.StatusBadBg,    AppTheme.StatusBadFg,    AppTheme.StatusBadBorder),
        BadgeStatus.Info   => (AppTheme.StatusInfoBg,   AppTheme.StatusInfoFg,   AppTheme.StatusInfoBorder),
        BadgeStatus.Muted  => (AppTheme.StatusMutedBg,  AppTheme.StatusMutedFg,  AppTheme.StatusMutedBorder),
        BadgeStatus.Accent => (AppTheme.StatusAccentBg, AppTheme.StatusAccentFg, AppTheme.StatusAccentBorder),
        _                  => (AppTheme.StatusMutedBg,  AppTheme.StatusMutedFg,  AppTheme.StatusMutedBorder),
    };

    private static string GetPrefix(BadgeStatus s) => s switch
    {
        BadgeStatus.Ok    => "✓ ",
        BadgeStatus.Warn  => "! ",
        BadgeStatus.Bad   => "× ",
        BadgeStatus.Info  => "i ",
        BadgeStatus.Muted => "— ",
        _                 => "",
    };
}
