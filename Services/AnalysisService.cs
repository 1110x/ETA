using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;

namespace ETA.Services;

public static class AnalysisService
{
    public static string GetDatabasePath() => DbPathHelper.DbPath;

    public static List<AnalysisItem> GetAllItems()
    {
        var items = new List<AnalysisItem>();

        string dbPath = GetDatabasePath();
        Console.WriteLine($"DB 경로 확인: {dbPath}");  // 여기서 로그
        if (!File.Exists(dbPath))
        {
            Console.WriteLine($"Database not found: {dbPath}");
            return items;
        }

        using var connection = new SqliteConnection($"Data Source={dbPath}");
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = @"
            SELECT Category, 
            Analyte, 
            Parts, 
            DecimalPlaces, 
            unit, 
            ES, 
            Method, 
            instrument
            FROM ""분석정보""
            ORDER BY ES ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            items.Add(new AnalysisItem
            {
                Category      = reader.GetStringOrEmpty("Category"),
                Analyte       = reader.GetStringOrEmpty("Analyte"),
                Parts         = reader.GetStringOrEmpty("Parts"),
                DecimalPlaces = reader.IsDBNull(reader.GetOrdinal("DecimalPlaces")) ? 0 : reader.GetInt32(reader.GetOrdinal("DecimalPlaces")),
                unit          = reader.GetStringOrEmpty("unit"),
                ES            = reader.GetStringOrEmpty("ES"),
                Method        = reader.GetStringOrEmpty("Method"),
                instrument    = reader.GetStringOrEmpty("instrument")
            });
        }

        return items;
    }

    // 확장 메서드
    private static string GetStringOrEmpty(this SqliteDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        return !reader.IsDBNull(ordinal) ? reader.GetString(ordinal) : string.Empty;
    }
}