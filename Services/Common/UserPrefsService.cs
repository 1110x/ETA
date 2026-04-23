using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ETA.Services.Common;

/// <summary>
/// 사용자별 앱 환경설정 저장소 (경로/옵션 등 비-DB 설정).
/// 저장 위치: `{프로젝트}/Logs/Users/{UserId}/UserPrefs.json`
/// WindowPositionManager와 같은 사용자 분리 정책을 따른다.
/// </summary>
public static class UserPrefsService
{
    public const string KeyTestRecordBookTemplateFolder = "TestRecordBookTemplateFolder";

    private static readonly object _lock = new();
    private static Dictionary<string, string> _cache = new();
    private static bool _loaded;

    private static string GetPrefsPath()
    {
        string userId = Environment.UserName;
        if (string.IsNullOrWhiteSpace(userId)) userId = "DefaultUser";
        string dir = Path.Combine(Directory.GetCurrentDirectory(), "Logs", "Users", userId);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "UserPrefs.json");
    }

    private static void EnsureLoaded()
    {
        if (_loaded) return;
        lock (_lock)
        {
            if (_loaded) return;
            try
            {
                var path = GetPrefsPath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var parsed = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                    if (parsed != null) _cache = parsed;
                }
            }
            catch { _cache = new Dictionary<string, string>(); }
            _loaded = true;
        }
    }

    public static string Get(string key, string fallback = "")
    {
        EnsureLoaded();
        lock (_lock)
            return _cache.TryGetValue(key, out var v) ? v : fallback;
    }

    public static void Set(string key, string value)
    {
        EnsureLoaded();
        lock (_lock)
        {
            _cache[key] = value ?? "";
            try
            {
                var json = JsonSerializer.Serialize(_cache,
                    new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(GetPrefsPath(), json);
            }
            catch { }
        }
    }

    // 편의 프로퍼티
    public static string TestRecordBookTemplateFolder
    {
        get => Get(KeyTestRecordBookTemplateFolder);
        set => Set(KeyTestRecordBookTemplateFolder, value);
    }
}
