using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.Common;
using ETA.Models;
using System.Diagnostics;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

public static class WasteCompanyService
{

    private static void EnsureAbbrevColumn(DbConnection conn)
    {
        if (!DbConnectionFactory.ColumnExists(conn, "폐수배출업소", "약칭"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE `폐수배출업소` ADD COLUMN `약칭` TEXT DEFAULT ''";
            try { cmd.ExecuteNonQuery(); } catch { }
        }
    }

    public static List<WasteCompany> GetAllItems()
    {
        var items = new List<WasteCompany>();

        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath))
        {
            Debug.WriteLine("❌ DB 파일 없음");
            return items;
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureAbbrevColumn(conn);

        using var command = conn.CreateCommand();
        command.CommandText = @"
            SELECT
                프로젝트,
                프로젝트명,
                관리번호,
                업체명,
                사업자번호,
                COALESCE(약칭, '') AS 약칭
            FROM `폐수배출업소`
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
                약칭           = GetStringOrEmpty(reader, "약칭"),
                Original업체명 = GetStringOrEmpty(reader, "업체명")
            };
            items.Add(company);
        }

        Debug.WriteLine($"📊 로드 완료: {items.Count}개 업체");
        return items;
    }

    private static string GetStringOrEmpty(DbDataReader reader, string col)
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


        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        EnsureAbbrevColumn(conn);

        cmd.CommandText = @"
            UPDATE `폐수배출업소`
            SET
                업체명       = @업체명,
                프로젝트     = @프로젝트,
                프로젝트명   = @프로젝트명,
                사업자번호   = @사업자번호,
                약칭         = @약칭
            WHERE 업체명 = @Original업체명";

        cmd.Parameters.AddWithValue("@업체명",       wasteCompany.업체명 ?? "");
        cmd.Parameters.AddWithValue("@프로젝트",     wasteCompany.프로젝트 ?? "");
        cmd.Parameters.AddWithValue("@프로젝트명",   wasteCompany.프로젝트명 ?? "");
        // 사업자번호 컬럼이 INTEGER 타입인 경우 빈 문자열 거부 → NULL로 전달
        var 사업자번호Val = string.IsNullOrEmpty(wasteCompany.사업자번호)
            ? (object)DBNull.Value
            : wasteCompany.사업자번호;
        cmd.Parameters.AddWithValue("@사업자번호", 사업자번호Val);
        cmd.Parameters.AddWithValue("@약칭",         wasteCompany.약칭 ?? "");
        cmd.Parameters.AddWithValue("@Original업체명", wasteCompany.Original업체명);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE] {rows}행 업데이트 (업체명: {wasteCompany.Original업체명})");

        if (rows > 0)
            wasteCompany.Original업체명 = wasteCompany.업체명 ?? "";  // 키 갱신

        return rows > 0;
    }
}