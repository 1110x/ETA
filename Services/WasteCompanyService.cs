using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class WasteCompanyService
{
    public static string GetDatabasePath() => DbPathHelper.DbPath;

    public static List<WasteCompany> GetAllItems()
    {
        var items = new List<WasteCompany>();
        string dbPath = GetDatabasePath();

        if (!File.Exists(dbPath))
        {
            Debug.WriteLine("❌ DB 파일 없음");
            return items;
        }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var command = conn.CreateCommand();
        command.CommandText = @"
            SELECT 
                프로젝트,
                프로젝트명,
                관리번호,
                업체명,
                사업자번호
            FROM ""폐수배출업소""
            ORDER BY 관리번호 ASC";

        Debug.WriteLine($"[DEBUG] 실행될 SQL → {command.CommandText.Replace("\r\n", " ")}");

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var company = new WasteCompany
            {
                프로젝트       = GetStringOrEmpty(reader, "프로젝트"),
                프로젝트명     = GetStringOrEmpty(reader, "프로젝트명"),
                관리번호       = GetStringOrEmpty(reader, "관리번호"),
                업체명         = GetStringOrEmpty(reader, "업체명"),           // ← 여기가 핵심!
                사업자번호     = GetStringOrEmpty(reader, "사업자번호"),
                Original업체명 = GetStringOrEmpty(reader, "업체명")
            };
            items.Add(company);
        }

        Debug.WriteLine($"📊 로드 완료: {items.Count}개 업체");
        return items;
    }

    private static string GetStringOrEmpty(SqliteDataReader reader, string col)
    {
        try
        {
            int i = reader.GetOrdinal(col);
            return !reader.IsDBNull(i) ? reader.GetString(i) ?? "" : "";
        }
        catch
        {
            return "";
        }
    }

    // 필요하면 이 메서드도 사용하세요 (Page에서 ExecuteDirectUpdate 대신 추천)
    public static bool Update(WasteCompany wasteCompany)
    {
        if (string.IsNullOrEmpty(wasteCompany.Original업체명))
        {
            Debug.WriteLine("❌ Original업체명 없음 → 저장 스킵");
            return false;
        }

        string dbPath = GetDatabasePath();

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE ""폐수배출업소"" 
            SET 
                프로젝트     = @프로젝트,
                프로젝트명   = @프로젝트명,
                관리번호     = @관리번호,
                사업자번호   = @사업자번호
            WHERE 업체명 = @Original업체명";

        cmd.Parameters.AddWithValue("@프로젝트",     wasteCompany.프로젝트 ?? "");
        cmd.Parameters.AddWithValue("@프로젝트명",   wasteCompany.프로젝트명 ?? "");
        cmd.Parameters.AddWithValue("@관리번호",     wasteCompany.관리번호 ?? "");
        cmd.Parameters.AddWithValue("@사업자번호",   wasteCompany.사업자번호 ?? "");
        cmd.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE] {rows}행 업데이트 (업체명: {wasteCompany.Original업체명})");

        return rows > 0;
    }
}