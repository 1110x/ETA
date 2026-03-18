using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;

namespace ETA.Services;

public static class ContractPriceService
{
    private static string GetDatabasePath()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Data", "eta.db");
        Console.WriteLine($"[ContractPriceService] DB 경로: {path}"); // ← 경로 출력 필수!
        return path;
    }

    public static List<ContractPrice> GetAllContractPrices()
    {
        Console.WriteLine("[ContractPriceService] GetAllContractPrices 호출됨");
        var prices = new List<ContractPrice>();

        string dbPath = GetDatabasePath();

        if (!File.Exists(dbPath))
        {
            Console.WriteLine($"[오류] DB 파일 없음: {dbPath}");
            return prices;
        }

        try
        {
            using var connection = new SqliteConnection($"Data Source={dbPath}");
            Console.WriteLine("[DB] 연결 시도...");
            connection.Open();
            Console.WriteLine("[DB] 연결 성공");

            using var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT 
                    ES,
                    Category,
                    Analyte,
                    FS100,
                    FS100Plus,
                    FS56,
                    NFS56,
                    FS55,
                    FS52,
                    FSHN52,
                    NFS50,
                    NFS45,
                    NFS39,
                    NFS36,
                    NFS36RE,
                    FS25
                FROM ""분석단가""
                ORDER BY ES ASC";

            Console.WriteLine("[DB] 쿼리 실행 중...");
            using var reader = command.ExecuteReader();

            int rowCount = 0;
            while (reader.Read())
            {
                rowCount++;
                prices.Add(new ContractPrice
                {
                    ES        = reader.GetStringOrEmpty("ES"),
                    Category  = reader.GetStringOrEmpty("Category"),
                    Analyte   = reader.GetStringOrEmpty("Analyte"),
                    FS100     = reader.SafeGetDecimal("FS100"),
                    FS100Plus = reader.SafeGetDecimal("FS100Plus"),
                    FS56      = reader.SafeGetDecimal("FS56"),
                    NFS56     = reader.SafeGetDecimal("NFS56"),
                    FS55      = reader.SafeGetDecimal("FS55"),
                    FS52      = reader.SafeGetDecimal("FS52"),
                    FSHN52    = reader.SafeGetDecimal("FSHN52"),
                    NFS50     = reader.SafeGetDecimal("NFS50"),
                    NFS45     = reader.SafeGetDecimal("NFS45"),
                    NFS39     = reader.SafeGetDecimal("NFS39"),
                    NFS36     = reader.SafeGetDecimal("NFS36"),
                    NFS36RE   = reader.SafeGetDecimal("NFS36RE"),
                    FS25      = reader.SafeGetDecimal("FS25")
                });
            }
            Console.WriteLine($"[DB] 읽은 행 수: {rowCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("★ [ContractPriceService] 심각한 오류 ★");
            Console.WriteLine($"메시지: {ex.Message}");
            Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            if (ex.InnerException != null)
                Console.WriteLine($"내부 오류: {ex.InnerException.Message}");
        }

        return prices;
    }

    // 확장 메서드 (로그 추가)
    private static string GetStringOrEmpty(this SqliteDataReader reader, string columnName)
    {
        try
        {
            int ordinal = reader.GetOrdinal(columnName);
            return !reader.IsDBNull(ordinal) ? reader.GetString(ordinal) : string.Empty;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[GetStringOrEmpty] 컬럼 오류: {columnName} - {ex.Message}");
            return string.Empty;
        }
    }

    private static decimal? SafeGetDecimal(this SqliteDataReader reader, string columnName)
    {
        try
        {
            int ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? null : reader.GetDecimal(ordinal);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[SafeGetDecimal] 컬럼 오류: {columnName} - {ex.Message}");
            return null;
        }
    }
}