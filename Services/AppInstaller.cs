using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Services;

// ── 설치 항목 정의 ────────────────────────────────────────────────────────────
public class InstallItem
{
    public string Name        { get; init; } = "";   // 표시 이름
    public string WingetId    { get; init; } = "";   // winget 패키지 ID
    public string CheckPath   { get; init; } = "";   // 설치 확인용 파일 경로
    public string DownloadUrl { get; init; } = "";   // winget 없을 때 수동 다운로드 URL
    public bool   Required    { get; init; } = true; // 필수 여부
    public string Description { get; init; } = "";   // 용도 설명
}

// ── 설치 결과 ─────────────────────────────────────────────────────────────────
public class InstallResult
{
    public string Name       { get; init; } = "";
    public bool   Success    { get; init; }
    public bool   Skipped    { get; init; }   // 이미 설치됨
    public string Message    { get; init; } = "";
}

// ── 진행 상태 콜백 ────────────────────────────────────────────────────────────
public class InstallProgress
{
    public string CurrentItem  { get; init; } = "";
    public int    Current      { get; init; }
    public int    Total        { get; init; }
    public string Status       { get; init; } = "";
}

// ═════════════════════════════════════════════════════════════════════════════
//  AppInstaller — 최초 실행 설치 관리자
// ═════════════════════════════════════════════════════════════════════════════
public static class AppInstaller
{
    // ── 설치 필요 패키지 목록 ─────────────────────────────────────────────────
    public static readonly List<InstallItem> RequiredPackages = new()
    {
        new InstallItem
        {
            Name        = "LibreOffice",
            WingetId    = "TheDocumentFoundation.LibreOffice",
            CheckPath   = @"C:\Program Files\LibreOffice\program\soffice.exe",
            DownloadUrl = "https://www.libreoffice.org/download/libreoffice-fresh/",
            Required    = false,   // PDF 출력에만 필요 (선택)
            Description = "PDF 출력 기능에 필요합니다.",
        },

        // 향후 추가 항목 예시:
        // new InstallItem
        // {
        //     Name        = "Git",
        //     WingetId    = "Git.Git",
        //     CheckPath   = @"C:\Program Files\Git\bin\git.exe",
        //     DownloadUrl = "https://git-scm.com/download/win",
        //     Required    = false,
        //     Description = "소스 코드 버전 관리에 필요합니다.",
        // },
    };

    // ── 설치 완료 마커 파일 경로 ──────────────────────────────────────────────
    private static string MarkerPath =>
        Path.Combine(AppContext.BaseDirectory, ".install_done");

    // ── 최초 실행 여부 확인 ───────────────────────────────────────────────────
    public static bool IsFirstRun() => !File.Exists(MarkerPath);

    // ── 최초 실행 완료 마킹 ───────────────────────────────────────────────────
    public static void MarkInstallDone()
    {
        try { File.WriteAllText(MarkerPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")); }
        catch { }
    }

    // ── 설치 필요한 항목만 필터링 ─────────────────────────────────────────────
    public static List<InstallItem> GetMissingPackages()
        => RequiredPackages
            .Where(p => !IsInstalled(p))
            .ToList();

    // ── 설치 여부 확인 ────────────────────────────────────────────────────────
    public static bool IsInstalled(InstallItem item)
    {
        if (!string.IsNullOrEmpty(item.CheckPath) && File.Exists(item.CheckPath))
            return true;

        // PATH 에서도 탐색
        if (!string.IsNullOrEmpty(item.CheckPath))
        {
            var fileName = Path.GetFileName(item.CheckPath);
            var pathEnv  = Environment.GetEnvironmentVariable("PATH") ?? "";
            foreach (var dir in pathEnv.Split(Path.PathSeparator))
            {
                if (File.Exists(Path.Combine(dir, fileName))) return true;
            }
        }
        return false;
    }

    // ── 전체 설치 실행 ────────────────────────────────────────────────────────
    /// <summary>
    /// missing 패키지를 순서대로 설치한다.
    /// onProgress 콜백으로 UI 진행 상황을 업데이트할 수 있다.
    /// </summary>
    public static async Task<List<InstallResult>> InstallAllAsync(
        List<InstallItem>?             items      = null,
        Action<InstallProgress>?       onProgress = null,
        bool                           silent     = true)
    {
        items ??= GetMissingPackages();
        var results = new List<InstallResult>();

        bool wingetOk = IsWingetAvailable();
        Log($"winget 사용가능={wingetOk}  설치항목={items.Count}개");

        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];

            onProgress?.Invoke(new InstallProgress
            {
                CurrentItem = item.Name,
                Current     = i + 1,
                Total       = items.Count,
                Status      = "설치 중...",
            });

            // 이미 설치됨
            if (IsInstalled(item))
            {
                Log($"[{item.Name}] 이미 설치됨 → 스킵");
                results.Add(new InstallResult
                {
                    Name    = item.Name,
                    Success = true,
                    Skipped = true,
                    Message = "이미 설치되어 있습니다.",
                });
                continue;
            }

            // winget 으로 설치
            if (wingetOk && !string.IsNullOrEmpty(item.WingetId))
            {
                Log($"[{item.Name}] winget 설치 시작...");
                var (ok, msg) = await WingetInstallAsync(item.WingetId, silent);

                results.Add(new InstallResult
                {
                    Name    = item.Name,
                    Success = ok,
                    Skipped = false,
                    Message = msg,
                });

                if (ok) Log($"[{item.Name}] ✅ 설치 완료");
                else
                {
                    Log($"[{item.Name}] ❌ 설치 실패: {msg}");
                    // winget 실패 시 브라우저 열기
                    if (!string.IsNullOrEmpty(item.DownloadUrl))
                        OpenUrl(item.DownloadUrl);
                }
            }
            else
            {
                // winget 없으면 브라우저로 다운로드 페이지 열기
                Log($"[{item.Name}] winget 없음 → 브라우저 열기");
                if (!string.IsNullOrEmpty(item.DownloadUrl))
                    OpenUrl(item.DownloadUrl);

                results.Add(new InstallResult
                {
                    Name    = item.Name,
                    Success = false,
                    Skipped = false,
                    Message = $"수동 설치 필요: {item.DownloadUrl}",
                });
            }

            onProgress?.Invoke(new InstallProgress
            {
                CurrentItem = item.Name,
                Current     = i + 1,
                Total       = items.Count,
                Status      = results.Last().Success ? "완료" : "실패",
            });
        }

        // 모든 필수 항목이 설치됐으면 마커 기록
        bool allRequired = RequiredPackages
            .Where(p => p.Required)
            .All(p => IsInstalled(p));

        if (allRequired) MarkInstallDone();

        return results;
    }

    // ── winget 설치 실행 ─────────────────────────────────────────────────────
    private static async Task<(bool ok, string msg)> WingetInstallAsync(
        string packageId, bool silent)
    {
        try
        {
            var args = $"install --id {packageId} -e " +
                       (silent ? "--silent " : "") +
                       "--accept-package-agreements --accept-source-agreements";

            using var proc = Process.Start(new ProcessStartInfo
            {
                FileName               = "winget",
                Arguments              = args,
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
            })!;

            // 비동기 대기 (최대 5분)
            var completed = await Task.Run(() => proc.WaitForExit(300_000));
            int exitCode  = completed ? proc.ExitCode : -1;

            return exitCode == 0
                ? (true,  "설치 완료")
                : (false, $"winget 종료코드 {exitCode}");
        }
        catch (Exception ex)
        {
            return (false, $"실행 오류: {ex.Message}");
        }
    }

    // ── winget 사용 가능 여부 ─────────────────────────────────────────────────
    public static bool IsWingetAvailable()
    {
        try
        {
            using var proc = Process.Start(new ProcessStartInfo
            {
                FileName               = "winget",
                Arguments              = "--version",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                CreateNoWindow         = true,
            })!;
            proc.WaitForExit(5_000);
            return proc.ExitCode == 0;
        }
        catch { return false; }
    }

    private static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { }
    }

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [Installer] {msg}";
        Debug.WriteLine(line);
        try
        {
            var logPath = Path.Combine(AppContext.BaseDirectory, "install.log");
            File.AppendAllText(logPath, line + Environment.NewLine);
        }
        catch { }
    }
}
