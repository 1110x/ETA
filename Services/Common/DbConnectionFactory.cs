using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using MySqlConnector;
using ETA.Services.Common;

namespace ETA.Services.Common;

/// <summary>
/// MariaDB 연결을 생성하는 팩토리.
/// </summary>
public static class DbConnectionFactory
{
    // ── SQL 헬퍼 상수 ────────────────────────────────────────────────────────

    public const string AutoIncrement = "AUTO_INCREMENT";
    public const string LastInsertId  = "LAST_INSERT_ID()";
    public const string RowId         = "_id";
    public const string NowExpr       = "NOW()";

    public static string DateFmt(string col, string fmt)
        => $"DATE_FORMAT({col}, '{fmt}')";

    /// <summary>MariaDB UPSERT suffix: ON DUPLICATE KEY UPDATE col = VALUES(col), ...</summary>
    public static string UpsertSuffix(string[] conflictCols, string[] updateCols)
    {
        var sets = string.Join(", ", System.Array.ConvertAll(updateCols, c => $"`{c}` = VALUES(`{c}`)"));
        return $"ON DUPLICATE KEY UPDATE {sets}";
    }

    // ── 스키마 헬퍼 ──────────────────────────────────────────────────────────

    public static List<string> GetColumnNames(DbConnection conn, string tableName)
    {
        // SHOW COLUMNS FROM: INFORMATION_SCHEMA보다 훨씬 빠름 (수초 → 수ms)
        var cols = new List<string>();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SHOW COLUMNS FROM `{tableName}`";
        try
        {
            using var r = cmd.ExecuteReader();
            while (r.Read())
                cols.Add(r.GetString(0)); // Field 컬럼
        }
        catch { }
        return cols;
    }

    public static bool ColumnExists(DbConnection conn, string tableName, string columnName)
        => GetColumnNames(conn, tableName).Contains(columnName);

    public static bool TableExists(DbConnection conn, string tableName)
    {
        // SHOW TABLES LIKE: INFORMATION_SCHEMA보다 빠름
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SHOW TABLES LIKE '{tableName.Replace("'", "\\'")}'";
        using var r = cmd.ExecuteReader();
        return r.Read();
    }

    // ── 연결 팩토리 ──────────────────────────────────────────────────────────

    public static DbConnection CreateConnection()
    {
        if (string.IsNullOrEmpty(DbPathHelper.MariaDbConnectionString))
            throw new InvalidOperationException(
                "MariaDB 연결 문자열이 설정되지 않았습니다. appsettings.json을 확인하세요.");
        return new MySqlConnection(DbPathHelper.MariaDbConnectionString);
    }
}

/// <summary>
/// System.Data.Common.DbParameterCollection에 AddWithValue 확장 메서드를 추가합니다.
/// </summary>
public static class DbParameterCollectionExtensions
{
    public static DbParameter AddWithValue(this DbParameterCollection parms, string name, object? value)
    {
        dynamic d = parms;
        return (DbParameter)d.AddWithValue(name, value ?? (object)DBNull.Value);
    }
}
