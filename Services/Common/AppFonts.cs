using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using System;
using System;

namespace ETA.Services.Common;

/// <summary>
/// 글자 크기 DynamicResource 헬퍼.
/// 슬라이더가 ApplyFontScale() 을 호출하면 Window.Resources 와
/// Application.Current.Resources 를 동시에 갱신하므로,
/// Obs() 로 반응형 바인딩을 사용할 수 있다.
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

    // ── 현재 값 스냅샷 (컨트롤 생성 시점에 한 번 읽는 용도) ────────────────
    public static double XS    => Get("FontSizeXS",    BaseXS);
    public static double SM    => Get("FontSizeSM",    BaseSM);
    public static double Base  => Get("FontSizeBase",  BaseBase);
    public static double MD    => Get("FontSizeMD",    BaseMD);
    public static double LG    => Get("FontSizeLG",    BaseLG);
    public static double XL    => Get("FontSizeXL",    BaseXL);
    public static double Title => Get("FontSizeTitle", BaseTitle);
    public static double Huge  => Get("FontSizeHuge",  BaseHuge);

    /// <summary>Application.Current.Resources 의 반응형 Observable 반환.
    /// 슬라이더를 움직이면 이 Observable 이 새 값을 방출한다.</summary>
    public static IObservable<object?> Obs(string key)
        => Application.Current?.GetResourceObservable(key)
           ?? new SingleValueObservable<object?>(Get(key, BaseBase));

    // ── TextBlock FontSize 반응형 바인딩 헬퍼 ────────────────────────────
    public static TextBlock BindXS(this TextBlock tb)
        => Fs(tb, "FontSizeXS");
    public static TextBlock BindSM(this TextBlock tb)
        => Fs(tb, "FontSizeSM");
    public static TextBlock BindBase(this TextBlock tb)
        => Fs(tb, "FontSizeBase");
    public static TextBlock BindMD(this TextBlock tb)
        => Fs(tb, "FontSizeMD");
    public static TextBlock BindLG(this TextBlock tb)
        => Fs(tb, "FontSizeLG");

    /// <summary>지정된 ResourceKey 로 FontSize 를 반응형 바인딩한다.</summary>
    public static TextBlock BindFs(this TextBlock tb, string key)
        => Fs(tb, key);

    // ── Button FontSize 반응형 바인딩 ────────────────────────────────────
    public static Button BindXS(this Button b) { b.Bind(Button.FontSizeProperty, Obs("FontSizeXS")); return b; }
    public static Button BindSM(this Button b) { b.Bind(Button.FontSizeProperty, Obs("FontSizeSM")); return b; }
    public static Button BindBase(this Button b) { b.Bind(Button.FontSizeProperty, Obs("FontSizeBase")); return b; }

    // ── TextBox FontSize 반응형 바인딩 ───────────────────────────────────
    public static TextBox BindSM(this TextBox t) { t.Bind(TextBox.FontSizeProperty, Obs("FontSizeSM")); return t; }
    public static TextBox BindBase(this TextBox t) { t.Bind(TextBox.FontSizeProperty, Obs("FontSizeBase")); return t; }

    private static TextBlock Fs(TextBlock tb, string key)
    {
        tb.Bind(TextBlock.FontSizeProperty, Obs(key));
        return tb;
    }

    private static double Get(string key, double fallback)
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true
            && v is double d)
            return d;
        return fallback;
    }

    private sealed class SingleValueObservable<T>(T value) : IObservable<T>
    {
        public IDisposable Subscribe(IObserver<T> observer)
        {
            observer.OnNext(value);
            observer.OnCompleted();
            return NullDisposable.Instance;
        }
    }

    private sealed class NullDisposable : IDisposable
    {
        public static readonly NullDisposable Instance = new();
        public void Dispose() { }
    }
}
