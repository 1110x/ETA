using Avalonia;
using Avalonia.Media;

namespace ETA.Views;

/// <summary>
/// 전역 테마 리소스 접근 헬퍼.
/// App.axaml 에 정의된 DynamicResource 를 C# 코드에서 사용할 때 이 클래스를 통해 접근.
/// 테마 변경 시 App.axaml 만 수정하면 전체 앱에 반영됨.
/// </summary>
internal static class AppTheme
{
    // ── 보더 (Border) ──
    public static IBrush BorderDefault  => Get("ThemeBorderDefault",  "#555577");
    public static IBrush BorderSubtle   => Get("ThemeBorderSubtle",   "#333344");
    public static IBrush BorderMuted    => Get("ThemeBorderMuted",    "#404050");
    public static IBrush BorderActive   => Get("ThemeBorderActive",   "#2a6a2a");
    public static IBrush BorderInfo     => Get("ThemeBorderInfo",     "#336699");
    public static IBrush BorderAccent   => Get("ThemeBorderAccent",   "#3a5a8a");
    public static IBrush BorderDanger   => Get("ThemeBorderDanger",   "#663333");
    public static IBrush BorderWarn     => Get("ThemeBorderWarn",     "#6a5a3a");
    public static IBrush BorderSeparator=> Get("ThemeBorderSeparator","#3a3a4a");
    public static IBrush BorderPanel    => Get("ThemeBorderPanel",    "#3a3a4a");

    // ── 전경색 (Foreground) ──
    public static IBrush FgPrimary      => Get("ThemeFgPrimary",     "#ffffff");
    public static IBrush FgSecondary    => Get("ThemeFgSecondary",   "#cccccc");
    public static IBrush FgMuted        => Get("ThemeFgMuted",       "#888899");
    public static IBrush FgDimmed       => Get("ThemeFgDimmed",      "#555566");
    public static IBrush FgSuccess      => Get("ThemeFgSuccess",     "#88cc88");
    public static IBrush FgInfo         => Get("ThemeFgInfo",        "#88aacc");
    public static IBrush FgDanger       => Get("ThemeFgDanger",      "#ff6666");
    public static IBrush FgWarn         => Get("ThemeFgWarn",        "#ffaa44");
    public static IBrush FgLink         => Get("ThemeFgLink",        "#88aaff");

    // ── 배경색 (Background) ──
    public static IBrush BgPrimary      => Get("ThemeBgPrimary",     "#1e1e26");
    public static IBrush BgSecondary    => Get("ThemeBgSecondary",   "#252535");
    public static IBrush BgCard         => Get("ThemeBgCard",        "#2a2a3a");
    public static IBrush BgInput        => Get("ThemeBgInput",       "#22223a");
    public static IBrush BgActiveGreen  => Get("ThemeBgActiveGreen", "#1a3a1a");
    public static IBrush BgActiveBlue   => Get("ThemeBgActiveBlue",  "#1e3a5a");
    public static IBrush BgDanger       => Get("ThemeBgDanger",      "#4a2a2a");

    // ── 폰트 크기 (DynamicResource 매핑) ──
    public static double FontXS   => GetDouble("FontSizeXS",    9);
    public static double FontSM   => GetDouble("FontSizeSM",   10);
    public static double FontBase => GetDouble("FontSizeBase",  11);
    public static double FontMD   => GetDouble("FontSizeMD",   12);
    public static double FontLG   => GetDouble("FontSizeLG",   13);
    public static double FontXL   => GetDouble("FontSizeXL",   14);

    /// <summary>리소스 키로 double 값 조회.</summary>
    public static double GetDouble(string key, double fallback)
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is double d)
            return d;
        return fallback;
    }

    /// <summary>리소스 키로 Brush 조회. 없으면 fallback 색상 반환.</summary>
    public static IBrush Get(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is IBrush b)
            return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }
}
