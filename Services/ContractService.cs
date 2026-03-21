using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;

namespace ETA.Services;

public static class ContractService
{
    private static string GetDatabasePath()
    {
        return Path.Combine(AppContext.BaseDirectory, "Data", "eta.db");
    }

    public static List<Contract> GetAllContracts()
    {
        var contracts = new List<Contract>();

        string dbPath = GetDatabasePath();

        if (!File.Exists(dbPath))
        {
            Console.WriteLine($"Database not found: {dbPath}");
            return contracts;
        }

        using var connection = new SqliteConnection($"Data Source={dbPath}");
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = @"
            SELECT 
                C_CompanyName,
                C_ContractStart,
                C_ContractEnd,
                C_ContractDays,
                C_ContractAmountVATExcluded,
                C_Abbreviation,
                C_ContractType,
                C_Address,
                C_Representative,
                C_FacilityType,
                C_CategoryType,
                C_MainProduct,
                C_ContactPerson,
                C_PhoneNumber,
                C_Email
            FROM ""계약 DB""
            ORDER BY C_CompanyName ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            contracts.Add(new Contract
            {
                C_CompanyName               = reader.GetStringOrEmpty("C_CompanyName"),

                C_ContractStart = reader.IsDBNull(reader.GetOrdinal("C_ContractStart"))
                    ? null
                    : reader.GetDateTime(reader.GetOrdinal("C_ContractStart")),

                C_ContractEnd = reader.IsDBNull(reader.GetOrdinal("C_ContractEnd"))
                    ? null
                    : reader.GetDateTime(reader.GetOrdinal("C_ContractEnd")),

                C_ContractDays              = reader.SafeGetInt32("C_ContractDays"),
                C_ContractAmountVATExcluded = reader.SafeGetDecimal("C_ContractAmountVATExcluded"),
                C_Abbreviation              = reader.GetStringOrEmpty("C_Abbreviation"),
                C_ContractType              = reader.GetStringOrEmpty("C_ContractType"),
                C_Address                   = reader.GetStringOrEmpty("C_Address"),
                C_Representative            = reader.GetStringOrEmpty("C_Representative"),
                C_FacilityType              = reader.GetStringOrEmpty("C_FacilityType"),
                C_CategoryType              = reader.GetStringOrEmpty("C_CategoryType"),
                C_MainProduct               = reader.GetStringOrEmpty("C_MainProduct"),
                C_ContactPerson             = reader.GetStringOrEmpty("C_ContactPerson"),
                C_PhoneNumber               = reader.GetStringOrEmpty("C_PhoneNumber"),
                C_Email                     = reader.GetStringOrEmpty("C_Email")
            });
        }
//C_ContractAmountVATExcluded
        return contracts;
    }

    // 확장 메서드들
    private static string GetStringOrEmpty(this SqliteDataReader reader, string columnName)
    {
        try
        {
            int ordinal = reader.GetOrdinal(columnName);
            return !reader.IsDBNull(ordinal) ? reader.GetString(ordinal) : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int? SafeGetInt32(this SqliteDataReader reader, string columnName)
    {
        try
        {
            int ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? null : reader.GetInt32(ordinal);
        }
        catch
        {
            return null;
        }
    }

private static decimal? SafeGetDecimal(this SqliteDataReader reader, string columnName)
{
    try
    {
        int ordinal = reader.GetOrdinal(columnName);
        Console.WriteLine($"[SafeGetDecimal] {columnName} 컬럼 찾음 (ordinal: {ordinal})");
        return reader.IsDBNull(ordinal) ? null : reader.GetDecimal(ordinal);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[SafeGetDecimal] 컬럼 '{columnName}' 읽기 실패: {ex.Message}");
        return null;
    }
}
}