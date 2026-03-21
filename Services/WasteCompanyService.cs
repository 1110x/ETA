using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;
namespace ETA.Services;

public static class WasteCompanyService
{
    public static string GetDatabasePath()
    {
// ★★★ 프로젝트 루트(Data 폴더)로 영구 고정 ★★★
    var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
    var dataDir = Path.Combine(projectRoot, "Data");
    Directory.CreateDirectory(dataDir);

    var dbPath = Path.Combine(dataDir, "eta.db");

    Debug.WriteLine($"[DB 위치] 프로젝트 루트 사용 → {dbPath}");
    return dbPath;
    }

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
            SELECT 프로젝트, 프로젝트명, 관리번호, 업체명, 사업자번호,
            FROM ""폐수배출업소""
            ORDER BY 사번 ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            var 관리번호 = GetStringOrEmpty(reader, "관리번호");

            var WasteCompany = new WasteCompany
            {
                프로젝트 = GetStringOrEmpty(reader, "프로젝트"),
                프로젝트명 = GetStringOrEmpty(reader, "프로젝트명"),
                관리번호 = GetStringOrEmpty(reader, "업체명"),
                사업자번호 = GetStringOrEmpty(reader, "사업자번호"),
                Original업체명 = GetStringOrEmpty(reader, "업체명"),
            };

            //string dateStr = GetStringOrEmpty(reader, "입사일");
            //if (DateOnly.TryParse(dateStr, out var date))
            //    agent.입사일 = date;

            items.Add(WasteCompany);
        }

        Debug.WriteLine($"📊 로드: {items.Count}명");
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
    public static bool Update(WasteCompany WasteCompany)
    {
        if (string.IsNullOrEmpty(WasteCompany.Original업체명))
        {
            Debug.WriteLine("❌ Original사번 없음 → 저장 스킵");
            return false;
        }

        string dbPath = GetDatabasePath();
        Debug.WriteLine($"[UPDATE 시도] Original관리번호='{WasteCompany.Original업체명}'");

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
        UPDATE ""WasteCompany"" SET 
            프로젝트=@프로젝트, 프로젝트명=@프로젝트명, 관리번호=@관리번호, 사업자번호=@사업자번호,
        WHERE 업체명=@Original업체명";

        cmd.Parameters.AddWithValue("@프로젝트", WasteCompany.프로젝트 ?? "");
        cmd.Parameters.AddWithValue("@프로젝트명", WasteCompany.프로젝트명 ?? "");
        cmd.Parameters.AddWithValue("@관리번호", WasteCompany.관리번호 ?? "");
        cmd.Parameters.AddWithValue("@업체명", WasteCompany.업체명 ?? "");
        cmd.Parameters.AddWithValue("@사업자번호",  WasteCompany.사업자번호 ?? "");


        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE 결과] {rows}행 영향받음 (1이면 성공, 0이면 실패)");

        if (rows > 0)
        {
           //WasteCompay.관리번호 = WasteCompay.관리번호
            Debug.WriteLine($"✅ DB 저장 성공: {WasteCompany.관리번호}");
            return true;
        }
        else
        {
            Debug.WriteLine($"❌ 실패 원인: 성명 '{WasteCompany.Original업체명}'인 레코드가 DB에 없음");
            return false;
        }
    }

}