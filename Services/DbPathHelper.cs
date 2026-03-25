using System;
using System.IO;

namespace ETA.Services;

/// <summary>
/// 전체 서비스에서 공통으로 사용하는 DB 경로.
/// C:\Users\{user}\Documents\ETA\Data\eta.db 로 고정.
/// </summary>
public static class DbPathHelper
{
    public static string DbPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Data", "eta.db");

    public static string PhotoDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Data", "Photos");

    static DbPathHelper()
    {
        // 폴더 없으면 자동 생성
        Directory.CreateDirectory(Path.GetDirectoryName(DbPath)!);
        Directory.CreateDirectory(PhotoDirectory);
    }
}
