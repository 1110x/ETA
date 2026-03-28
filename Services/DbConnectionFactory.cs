using System;
using System.Data;
using System.Data.Common;
using Microsoft.Data.Sqlite;
using MySqlConnector;

namespace ETA.Services;

/// <summary>
/// SQLite(로컬) 또는 MariaDB(서버) 연결을 생성하는 팩토리.
/// UseMariaDb 플래그로 런타임 전환 가능.
/// </summary>
public static class DbConnectionFactory
{
    /// <summary>true = MariaDB 사용, false = 로컬 SQLite 사용</summary>
    public static bool UseMariaDb { get; set; } = false;

    /// <summary>현재 선택된 DB 타입</summary>
    public static bool IsMariaDb => UseMariaDb;

    // ── SQL 방언 헬퍼 ────────────────────────────────────────────────────────

    /// <summary>AUTO_INCREMENT / AUTOINCREMENT</summary>
    public static string AutoIncrement => IsMariaDb ? "AUTO_INCREMENT" : "AUTOINCREMENT";

    /// <summary>마지막 삽입 ID 조회 함수</summary>
    public static string LastInsertId => IsMariaDb ? "LAST_INSERT_ID()" : "last_insert_rowid()";

    /// <summary>
    /// UPSERT 뒷부분 생성.
    /// SQLite : ON CONFLICT(conflictCols) DO UPDATE SET col = excluded.col, ...
    /// MariaDB: ON DUPLICATE KEY UPDATE col = VALUES(col), ...
    /// </summary>
    public static string UpsertSuffix(string[] conflictCols, string[] updateCols)
    {
        if (IsMariaDb)
        {
            var sets = string.Join(", ", System.Array.ConvertAll(updateCols, c => $"`{c}` = VALUES(`{c}`)"));
            return $"ON DUPLICATE KEY UPDATE {sets}";
        }
        else
        {
            var conflict = string.Join(", ", conflictCols);
            var sets = string.Join(",\n                        ",
                System.Array.ConvertAll(updateCols, c => $"{c} = excluded.{c}"));
            return $"ON CONFLICT({conflict}) DO UPDATE SET\n                        {sets}";
        }
    }

    // ── 연결 팩토리 ──────────────────────────────────────────────────────────

    /// <summary>
    /// 현재 설정에 따라 DbConnection 인스턴스를 반환합니다.
    /// 반환된 연결은 아직 Open()되지 않은 상태입니다.
    /// </summary>
    public static DbConnection CreateConnection()
    {
        if (IsMariaDb)
        {
            if (string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
                throw new InvalidOperationException(
                    "MariaDB 연결 문자열이 설정되지 않았습니다. appsettings.json을 확인하세요.");
            return new MySqlConnection(DbPathHelper.MariaDbConnectionString);
        }
        return new SqliteConnection($"Data Source={DbPathHelper.DbPath}");
    }
}

/// <summary>
/// System.Data.Common.DbParameterCollection에 AddWithValue 확장 메서드를 추가합니다.
/// MySqlParameterCollection / SqliteParameterCollection 모두 지원.
/// </summary>
public static class DbParameterCollectionExtensions
{
    public static DbParameter AddWithValue(this DbParameterCollection parms, string name, object? value)
    {
        dynamic d = parms;
        return (DbParameter)d.AddWithValue(name, value ?? (object)DBNull.Value);
    }
}
