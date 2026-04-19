using System;
using System.IO;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;

namespace ETA.Views;

/// <summary>
/// wireframe v0.1 · Show1/2/3/4 등 윈도우 크기·위치 기억.
///
/// 사용:
///   WindowSizePersister.Persist(myWindow, "show1");
///   WindowSizePersister.Persist(myWindow, "show2");
///   WindowSizePersister.Persist(myWindow, "processing_facility");
///
/// 저장 위치: %AppData%/ETA/window_sizes.json
///   { "show1": { "W": 980, "H": 640, "X": 120, "Y": 80 }, ... }
///
/// 같은 키를 여러 번 호출해도 안전 (OnOpened/OnClosed 중복 등록 방지).
/// </summary>
public static class WindowSizePersister
{
    private static readonly string StorePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ETA",
        "window_sizes.json");

    private sealed class Rect { public double W { get; set; } public double H { get; set; } public double X { get; set; } public double Y { get; set; } }

    private static readonly System.Collections.Generic.HashSet<string> s_registered = new();

    public static void Persist(Window window, string key)
    {
        // 중복 등록 가드 — 같은 (window, key) 조합은 한 번만.
        var guard = $"{window.GetHashCode()}:{key}";
        if (!s_registered.Add(guard)) return;

        // ── 복원 ──
        var store = Load();
        if (store.TryGetValue(key, out var saved))
        {
            if (saved.W > 100) window.Width = saved.W;
            if (saved.H > 100) window.Height = saved.H;
            if (!double.IsNaN(saved.X) && !double.IsNaN(saved.Y))
            {
                window.Position = new PixelPoint((int)saved.X, (int)saved.Y);
                window.WindowStartupLocation = WindowStartupLocation.Manual;
            }
        }

        // ── 저장 훅 ──
        EventHandler? save = null;
        save = (_, _) =>
        {
            try
            {
                var s = Load();
                s[key] = new Rect
                {
                    W = window.Width,
                    H = window.Height,
                    X = window.Position.X,
                    Y = window.Position.Y,
                };
                Save(s);
            }
            catch { /* 파일 IO 실패는 조용히 무시 — UX 블로킹 방지 */ }
        };

        window.Closing += (_, _) => save?.Invoke(null, EventArgs.Empty);
        // 리사이즈/이동 직후에도 한번 — 크래시 대비
        window.PositionChanged += (_, _) => save?.Invoke(null, EventArgs.Empty);
        window.SizeChanged     += (_, _) => save?.Invoke(null, EventArgs.Empty);
    }

    private static System.Collections.Generic.Dictionary<string, Rect> Load()
    {
        try
        {
            if (!File.Exists(StorePath))
                return new System.Collections.Generic.Dictionary<string, Rect>();
            var json = File.ReadAllText(StorePath);
            return JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, Rect>>(json)
                ?? new System.Collections.Generic.Dictionary<string, Rect>();
        }
        catch
        {
            return new System.Collections.Generic.Dictionary<string, Rect>();
        }
    }

    private static void Save(System.Collections.Generic.Dictionary<string, Rect> store)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(StorePath)!);
            var json = JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(StorePath, json);
        }
        catch { /* 무시 */ }
    }
}
