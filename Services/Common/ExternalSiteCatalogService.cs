using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace ETA.Services.Common;

/// <summary>
/// 외부 시스템 (WAYBLE) 의 사이트·시료 카탈로그 로더.
/// Python 스크립트 Scripts/extract_wayble_catalog.py 가 xlsx 에서 뽑아둔
/// Data/wayble_catalog.json 을 읽어 메모리 캐시.
/// </summary>
public static class ExternalSiteCatalogService
{
    public record Entry(
        string OfficeCd,
        string SiteCd,
        string SiteNm,
        string SampleCategory,
        string SampleCategoryNm);

    private static readonly string JsonPath = Path.Combine("Data", "wayble_catalog.json");
    private static List<Entry>? _cache;
    private static string _cacheSource = "";

    public static string JsonFilePath => JsonPath;
    public static string LastSource => _cacheSource;

    public static List<Entry> GetWayble(bool forceReload = false)
    {
        if (!forceReload && _cache != null) return _cache;

        var list = new List<Entry>();
        try
        {
            if (!File.Exists(JsonPath))
            {
                _cache = list;
                return _cache;
            }

            using var doc = JsonDocument.Parse(File.ReadAllText(JsonPath));
            var root = doc.RootElement;

            _cacheSource = root.TryGetProperty("source", out var src) ? src.GetString() ?? "" : "";

            if (root.TryGetProperty("entries", out var entries) && entries.ValueKind == JsonValueKind.Array)
            {
                foreach (var e in entries.EnumerateArray())
                {
                    list.Add(new Entry(
                        e.GetProperty("officeCd").GetString() ?? "",
                        e.GetProperty("siteCd").GetString() ?? "",
                        e.GetProperty("siteNm").GetString() ?? "",
                        e.GetProperty("sampleCategory").GetString() ?? "",
                        e.GetProperty("sampleCategoryNm").GetString() ?? ""));
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ExternalSiteCatalogService.GetWayble] {ex.Message}");
        }

        _cache = list;
        return _cache;
    }

    /// <summary>Python 스크립트 재실행 → 카탈로그 갱신. 성공 시 개수, 실패 시 null 반환.</summary>
    public static (int count, string? error) RefreshFromXlsx()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName               = "python3",
                Arguments              = "Scripts/extract_wayble_catalog.py",
                WorkingDirectory       = Directory.GetCurrentDirectory(),
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                UseShellExecute        = false,
                CreateNoWindow         = true,
            };
            using var p = Process.Start(psi);
            if (p == null) return (0, "프로세스 시작 실패");
            p.WaitForExit(60_000);
            if (p.ExitCode != 0)
            {
                var err = p.StandardError.ReadToEnd();
                return (0, $"exit={p.ExitCode}: {err}");
            }
            var list = GetWayble(forceReload: true);
            return (list.Count, null);
        }
        catch (Exception ex)
        {
            return (0, ex.Message);
        }
    }
}
