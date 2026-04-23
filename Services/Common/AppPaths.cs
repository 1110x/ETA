using System;
using System.IO;

namespace ETA.Services.Common;

/// <summary>
/// 개발(bin/Debug|Release/...)과 설치(C:\Program Files\ETA\...) 양쪽에서 일관된 경로 제공.
///
/// - <see cref="RootPath"/>: Data/Templates, Assets 가 들어있는 "리포지토리 루트" (read-only OK).
/// - <see cref="WritableDataRoot"/>: 로그·사용자 설정 등 쓰기 가능한 위치.
///     설치 환경에선 %LOCALAPPDATA%\ETA, 개발에선 리포 루트 사용.
/// </summary>
public static class AppPaths
{
    private static readonly Lazy<string> _rootPath = new(ResolveRootPath);
    private static readonly Lazy<string> _writablePath = new(ResolveWritablePath);

    /// <summary>Data/Templates, Assets 등 정적 리소스 폴더 기준 루트.</summary>
    public static string RootPath => _rootPath.Value;

    /// <summary>로그·사용자 설정 등 쓰기가 필요한 파일의 기준 루트.</summary>
    public static string WritableDataRoot => _writablePath.Value;

    /// <summary>Logs 폴더 (쓰기 가능).</summary>
    public static string LogsDir
    {
        get
        {
            var d = Path.Combine(WritableDataRoot, "Logs");
            try { Directory.CreateDirectory(d); } catch { }
            return d;
        }
    }

    private static string ResolveRootPath()
    {
        var baseDir = AppContext.BaseDirectory;

        // 개발 실행: bin/Debug/net8.0/ 또는 bin/Release/net8.0/  → 3단계 위가 리포 루트
        if (baseDir.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
            || baseDir.Contains("/bin/", StringComparison.OrdinalIgnoreCase))
        {
            var devRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", ".."));
            if (Directory.Exists(Path.Combine(devRoot, "Data")) || File.Exists(Path.Combine(devRoot, "ETA.csproj")))
                return devRoot;
        }

        // 설치 실행: 리소스가 exe 옆에 배포됨
        return baseDir.TrimEnd(Path.DirectorySeparatorChar);
    }

    private static string ResolveWritablePath()
    {
        // 개발 환경에선 리포 루트를 그대로 사용 (기존 Logs/ 유지)
        if (IsDevEnvironment())
            return RootPath;

        // 설치 환경: %LOCALAPPDATA%\ETA
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var dir = Path.Combine(appData, "ETA");
        try { Directory.CreateDirectory(dir); } catch { }
        return dir;
    }

    private static bool IsDevEnvironment()
    {
        var baseDir = AppContext.BaseDirectory;
        return baseDir.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
            || baseDir.Contains("/bin/", StringComparison.OrdinalIgnoreCase);
    }
}
