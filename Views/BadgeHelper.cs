using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.Common;

namespace ETA.Views;

/// <summary>
/// 아이콘 + 컬러 태그 배지 생성 헬퍼.
/// 분석항목, 업체, 사용자, 지역 등을 시각적으로 표시.
/// </summary>
internal static class Badge
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 약칭 (가장 빈번하게 사용 — BadgeColorHelper 초성 색상 연동) ──
    public static Border Abbr(string name)
    {
        var (bg, fg) = BadgeColorHelper.GetBadgeColor(name);
        return Make("🏢", name, bg, fg, fg); // border = fg color (subtle)
    }

    // ── 업체 ──
    public static Border Company(string name) => Make("🏭", name, "#1e3a5a", "#88aacc", "#336699");

    // ── 지역 (여수/율촌/세풍) ──
    public static Border Region(string region) => region switch
    {
        "여수" => Make("📍", region, "#1e3a5a", "#88aacc", "#336699"),
        "율촌" => Make("🏗", region, "#1a3a1a", "#88cc88", "#2a6a2a"),
        "세풍" => Make("🏭", region, "#3a2a1a", "#ccaa66", "#6a5a2a"),
        _      => Make("📍", region, "#2a2a3a", "#aaaacc", "#555577"),
    };

    // ── 분석항목 ──
    public static Border Analyte(string item) => item switch
    {
        "BOD"         => Make("🧪", item, "#3a1a1a", "#ff8866", "#6a3a2a"),
        "COD"         => Make("🧪", item, "#3a2a1a", "#ffaa66", "#6a4a2a"),
        "SS"          => Make("🧪", item, "#2a2a1a", "#ddcc44", "#5a5a2a"),
        "T-N" or "TN" => Make("🧪", item, "#1a3a1a", "#66cc66", "#2a6a2a"),
        "T-P" or "TP" => Make("🧪", item, "#2a1a3a", "#cc66cc", "#5a2a6a"),
        "pH"          => Make("🧪", item, "#1a2a3a", "#6699cc", "#2a4a6a"),
        _             => Make("🧪", item, "#1a2a2a", "#88bbcc", "#2a5a5a"),
    };

    // ── 사용자 ──
    public static Border User(string name) => Make("👤", name, "#1a1a3a", "#aaaaee", "#3a3a6a");

    // ── 부서/팀 ──
    public static Border Team(string name) => Make("📋", name, "#1a2a1a", "#88cc88", "#2a5a2a");

    // ── 상태 ──
    public static Border Status(string status) => status switch
    {
        "완료" => Make("✅", status, "#1a2a1a", "#88cc88", "#2a6a2a"),
        "진행" or "진행중" => Make("🔄", status, "#1a2a3a", "#88aaee", "#2a4a6a"),
        "대기" => Make("⏳", status, "#2a2a1a", "#ccaa66", "#5a5a2a"),
        "반려" => Make("❌", status, "#3a1a1a", "#ff6666", "#6a2a2a"),
        _      => Make("📌", status, "#2a2a3a", "#aaaacc", "#555577"),
    };

    // ── 범용: 아이콘 + 텍스트 ──
    public static Border Custom(string icon, string text, string bg, string fg, string border)
        => Make(icon, text, bg, fg, border);

    /// <summary>태그 배지 생성</summary>
    private static Border Make(string icon, string text, string bg, string fg, string bd)
    {
        return new Border
        {
            Background      = new SolidColorBrush(Color.Parse(bg)),
            BorderBrush     = new SolidColorBrush(Color.Parse(bd)),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(10),
            Padding         = new Thickness(6, 1, 8, 1),
            Margin          = new Thickness(0, 0, 4, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing     = 3,
                Children    =
                {
                    new TextBlock
                    {
                        Text = icon, FontSize = AppTheme.FontSM,
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new TextBlock
                    {
                        Text       = text,
                        FontFamily = Font, FontSize = AppTheme.FontSM, FontWeight = FontWeight.Medium,
                        Foreground = new SolidColorBrush(Color.Parse(fg)),
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                }
            }
        };
    }
}
