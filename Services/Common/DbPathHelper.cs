using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace ETA.Services.Common;

/// <summary>
/// MariaDB 연결 설정. appsettings.json에서 서버/포트/계정 정보를 읽습니다.
/// DB 이름은 DbRotationService.CurrentDbName으로 자동 결정됩니다.
/// </summary>
public static class DbPathHelper
{
    /// <summary>현재 운영 기간 DB 연결 문자열 (Database= 포함)</summary>
    public static string MariaDbConnectionString { get; private set; } = "";

    /// <summary>DB 이름 없는 기본 연결 문자열 — DB 생성/삭제 등 관리 작업용</summary>
    public static string BaseConnectionString { get; private set; } = "";

    /// <summary>애플리케이션 DB 사용자 이름 (GRANT 시 사용)</summary>
    public static string AppUser { get; private set; } = "";

    /// <summary>appsettings.json에 설정된 원본 DB 이름 (기본값 eta_db)</summary>
    public static string ConfigDbName { get; private set; } = "eta_db";

    /// <summary>로그인 화면에서 사용자가 선택한 DB로 전환</summary>
    public static void SetActiveDb(string dbName)
    {
        if (!string.IsNullOrEmpty(BaseConnectionString) && !string.IsNullOrEmpty(dbName))
            MariaDbConnectionString = BaseConnectionString + $"Database={dbName};";
    }

    public static string PhotoDirectory { get; } = Path.Combine(
       Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
       "ETA", "Data", "Photos");

    public static string AttachmentDirectory { get; } = Path.Combine(
       Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
       "ETA", "Data", "Attachments");

    public static string SharedPhotoDirectory { get; private set; } = "";

    static DbPathHelper()
    {
        Directory.CreateDirectory(PhotoDirectory);
        Directory.CreateDirectory(AttachmentDirectory);

        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true)
            .Build();

        var server = config["MariaDb:Server"];
        if (!string.IsNullOrEmpty(server))
        {
            string port     = config["MariaDb:Port"]     ?? "3306";
            string user     = config["MariaDb:User"]     ?? "";
            string password = config["MariaDb:Password"] ?? "";

            AppUser = user;
            string dbName = config["MariaDb:Database"] ?? "eta_db";
            ConfigDbName = dbName;
            BaseConnectionString =
                $"Server={server};Port={port};" +
                $"User={user};Password={password};" +
                "CharSet=utf8mb4;SslMode=None;";

            // 기본 연결은 appsettings DB (eta_db) 사용 — 로그인에서 변경 가능
            MariaDbConnectionString = BaseConnectionString + $"Database={dbName};";
        }

        var sharedPhoto = config["Paths:SharedPhotoDirectory"]?.Trim();
        if (!string.IsNullOrEmpty(sharedPhoto))
        {
            SharedPhotoDirectory = sharedPhoto;
            // NAS 디렉토리 생성은 백그라운드에서 — 연결 지연 시 앱 구동 차단 방지
            if (OperatingSystem.IsWindows())
            {
                System.Threading.Tasks.Task.Run(() =>
                {
                    try { Directory.CreateDirectory(SharedPhotoDirectory); } catch { }
                });
            }
        }
    }
}
