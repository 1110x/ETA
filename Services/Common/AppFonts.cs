using Avalonia;
using ETA.Services.Common;

namespace ETA.Services.Common;

/// <summary>
/// 글자 크기 DynamicResource 헬퍼 — 코드비하인드에서 사용.
/// 슬라이더가 ApplyFontScale() 을 호출할 때마다 Window.Resources 가 갱신되므로,
/// 컨트롤을 새로 생성하는 시점에 이 값을 읽으면 현재 스케일이 적용됨.
/// </summary>
public static class AppFonts
{
    // ── 기본 크기 (scale=1.0 기준) ────────────────────────────────────────
    public const double BaseXS    =  9.0;
    public const double BaseSM    = 10.0;
    public const double BaseBase  = 11.0;
    public const double BaseMD    = 12.0;
    public const double BaseLG    = 13.0;
    public const double BaseXL    = 14.0;
    public const double BaseTitle = 22.0;
    public const double BaseHuge  = 28.0;

    // ── DynamicResource 에서 현재 값 읽기 ────────────────────────────────
    public static double XS    => Get("FontSizeXS",    BaseXS);
    public static double SM    => Get("FontSizeSM",    BaseSM);
    public static double Base  => Get("FontSizeBase",  BaseBase);
    public static double MD    => Get("FontSizeMD",    BaseMD);
    public static double LG    => Get("FontSizeLG",    BaseLG);
    public static double XL    => Get("FontSizeXL",    BaseXL);
    public static double Title => Get("FontSizeTitle", BaseTitle);
    public static double Huge  => Get("FontSizeHuge",  BaseHuge);

    private static double Get(string key, double fallback)
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true
            && v is double d)
            return d;
        return fallback;
    }
}
