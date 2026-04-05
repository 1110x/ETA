using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using ETA.Services.Common;

namespace ETA.Services.Common;

/// <summary>
/// DB 연결 설정. SQLite(로컬) 또는 MariaDB(서버) 선택 가능.
/// MariaDB 연결 정보는 appsettings.json에서 읽습니다.
/// </summary>
public static class DbPathHelper
{
    // SQLite 로컬 경로 (로컬 모드 또는 폴백용)
    public static string DbPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Data", "eta.db");

    // MariaDB 연결 문자열 — appsettings.json의 "MariaDb" 섹션에서 로드
    // appsettings.json 예시:
    // {
    //   "MariaDb": {
    //     "Server": "1110s.synology.me",
    //     "Port": "3306",
    //     "Database": "eta_db",
    //     "User": "eta_user",
    //     "Password": "YOUR_PASSWORD_HERE"
    //   }
    // }
    public static string MariaDbConnectionString { get; private set; } = "";

    public static string PhotoDirectory { get; } = Path.Combine(
       Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
       "ETA", "Data", "Photos");

    public static string AttachmentDirectory { get; } = Path.Combine(
       Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
       "ETA", "Data", "Attachments");

    // NAS/파일서버 공유 경로(선택): 예) \\NAS\ETA\Data\Photos
    // 설정 키: Paths:SharedPhotoDirectory
    public static string SharedPhotoDirectory { get; private set; } = "";

    static DbPathHelper()
    {
        // 폴더 없으면 자동 생성
        Directory.CreateDirectory(Path.GetDirectoryName(DbPath)!);
        Directory.CreateDirectory(PhotoDirectory);
        Directory.CreateDirectory(AttachmentDirectory);

        // appsettings.json에서 MariaDB 설정 로드
        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true)
            .Build();

        var s = config["MariaDb:Server"];
        if (!string.IsNullOrEmpty(s))
        {
            MariaDbConnectionString =
                $"Server={s};" +
                $"Port={config["MariaDb:Port"] ?? "3306"};" +
                $"Database={config["MariaDb:Database"]};" +
                $"User={config["MariaDb:User"]};" +
                $"Password={config["MariaDb:Password"]};" +
                "CharSet=utf8mb4;SslMode=None;";
        }

        var sharedPhoto = config["Paths:SharedPhotoDirectory"]?.Trim();
        if (!string.IsNullOrEmpty(sharedPhoto))
        {
            SharedPhotoDirectory = sharedPhoto;
            // Windows UNC 경로는 macOS/Linux에서 생성하면 로컬에 잘못된 폴더가 만들어짐
            if (OperatingSystem.IsWindows())
            {
                try { Directory.CreateDirectory(SharedPhotoDirectory); } catch { }
            }
        }
    }
}
