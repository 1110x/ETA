using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.Common;
using ETA.Models;

namespace ETA.Services;

public static class AnalysisService
{

    public static List<AnalysisItem> GetAllItems()
    {
        var items = new List<AnalysisItem>();

        Console.WriteLine($"DB 경로 확인: {DbPathHelper.DbPath}");  // 여기서 로그
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath))
        {
            Console.WriteLine($"Database not found: {DbPathHelper.DbPath}");
            return items;
        }

        using var connection = DbConnectionFactory.CreateConnection();
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
            FROM `분석정보`
            ORDER BY ES ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            var item = new AnalysisItem
            {
                Category      = reader.GetStringOrEmpty("Category"),
                Analyte       = reader.GetStringOrEmpty("Analyte"),
                Parts         = reader.GetStringOrEmpty("Parts"),
                DecimalPlaces = reader.IsDBNull(reader.GetOrdinal("DecimalPlaces")) ? 0 : reader.GetInt32(reader.GetOrdinal("DecimalPlaces")),
                unit          = reader.GetStringOrEmpty("unit"),
                ES            = reader.GetStringOrEmpty("ES"),
                Method        = reader.GetStringOrEmpty("Method"),
                instrument    = reader.GetStringOrEmpty("instrument")
            };
            items.Add(item);
            Console.WriteLine($"[GetAllItems] 아이템: {item.Analyte} ({item.Category})");
        }

        Console.WriteLine($"[GetAllItems] 총 아이템 수: {items.Count}");
        return items;
    }

    // 확장 메서드
    private static string GetStringOrEmpty(this DbDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        return !reader.IsDBNull(ordinal) ? reader.GetString(ordinal) : string.Empty;
    }
}