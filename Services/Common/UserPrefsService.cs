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
    public const string KeyTradeStatementFolder = "TradeStatementFolder";
    public const string KeyQuotationExportFolder = "QuotationExportFolder";

    // 시험성적서 / 시험기록부 결재정보 (서명란에 들어갈 사람 이름)
    public const string KeyTestReportSignerQualityMgr = "TestReportSignerQualityMgr";
    public const string KeyTestRecordBookSignerAuthor = "TestRecordBookSignerAuthor";
    public const string KeyTestRecordBookSignerReviewer = "TestRecordBookSignerReviewer";
    public const string KeyTestRecordBookSignerApprover = "TestRecordBookSignerApprover";

    // 시험성적서 비고문구 (정도보증 적용 / 미적용 분기)
    public const string KeyTestReportRemarkQc  = "TestReportRemarkQc";
    public const string KeyTestReportRemarkRef = "TestReportRemarkRef";

    public const string DefaultTestReportRemarkQc  =
        "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분석결과 입니다.";
    public const string DefaultTestReportRemarkRef =
        "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분석결과입니다.";

    private static readonly object _lock = new();
    private static Dictionary<string, string> _cache = new();
    private static bool _loaded;

    private static string GetPrefsPath()
    {
        string userId = Environment.UserName;
        if (string.IsNullOrWhiteSpace(userId)) userId = "DefaultUser";
        string dir = Path.Combine(AppPaths.WritableDataRoot, "Logs", "Users", userId);
        try { Directory.CreateDirectory(dir); }
        catch { dir = Path.Combine(Path.GetTempPath(), "ETA", "Logs", "Users", userId);
                try { Directory.CreateDirectory(dir); } catch { } }
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

    public static string TradeStatementFolder
    {
        get => Get(KeyTradeStatementFolder);
        set => Set(KeyTradeStatementFolder, value);
    }

    public static string QuotationExportFolder
    {
        get => Get(KeyQuotationExportFolder);
        set => Set(KeyQuotationExportFolder, value);
    }

    public static string TestReportSignerQualityMgr
    {
        get => Get(KeyTestReportSignerQualityMgr);
        set => Set(KeyTestReportSignerQualityMgr, value);
    }

    public static string TestRecordBookSignerAuthor
    {
        get => Get(KeyTestRecordBookSignerAuthor);
        set => Set(KeyTestRecordBookSignerAuthor, value);
    }

    public static string TestRecordBookSignerReviewer
    {
        get => Get(KeyTestRecordBookSignerReviewer);
        set => Set(KeyTestRecordBookSignerReviewer, value);
    }

    public static string TestRecordBookSignerApprover
    {
        get => Get(KeyTestRecordBookSignerApprover);
        set => Set(KeyTestRecordBookSignerApprover, value);
    }

    /// <summary>정도보증 적용 시험성적서 비고문구. 미설정이면 기본값.</summary>
    public static string TestReportRemarkQc
    {
        get
        {
            var v = Get(KeyTestReportRemarkQc);
            return string.IsNullOrWhiteSpace(v) ? DefaultTestReportRemarkQc : v;
        }
        set => Set(KeyTestReportRemarkQc, value);
    }

    /// <summary>참고용(정도보증 미적용) 시험성적서 비고문구. 미설정이면 기본값.</summary>
    public static string TestReportRemarkRef
    {
        get
        {
            var v = Get(KeyTestReportRemarkRef);
            return string.IsNullOrWhiteSpace(v) ? DefaultTestReportRemarkRef : v;
        }
        set => Set(KeyTestReportRemarkRef, value);
    }
}
