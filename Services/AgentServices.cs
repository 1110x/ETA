using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;
namespace ETA.Services;

public static class AgentService
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

    public static List<Agent> GetAllItems()
    {
        var items = new List<Agent>();
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
            SELECT 성명, 직급, 직무, 입사일, 사번,
                   자격사항, Email, 기타, 측정인고유번호
            FROM ""Agent""
            ORDER BY 사번 ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            var 사번값 = GetStringOrEmpty(reader, "사번");

            var agent = new Agent
            {
                성명 = GetStringOrEmpty(reader, "성명"),
                직급 = GetStringOrEmpty(reader, "직급"),
                직무 = GetStringOrEmpty(reader, "직무"),
                사번 = GetStringOrEmpty(reader, "사번"),
                Original성명 = GetStringOrEmpty(reader, "성명"),   // ← 여기 추가
                자격사항 = GetStringOrEmpty(reader, "자격사항"),
                Email = GetStringOrEmpty(reader, "Email"),
                기타 = GetStringOrEmpty(reader, "기타"),
                측정인고유번호 = GetStringOrEmpty(reader, "측정인고유번호")
            };

            string dateStr = GetStringOrEmpty(reader, "입사일");
            if (DateOnly.TryParse(dateStr, out var date))
                agent.입사일 = date;

            items.Add(agent);
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
    public static bool Update(Agent agent)
    {
        if (string.IsNullOrEmpty(agent.Original성명))
        {
            Debug.WriteLine("❌ Original사번 없음 → 저장 스킵");
            return false;
        }

        string dbPath = GetDatabasePath();
        Debug.WriteLine($"[UPDATE 시도] Original사번='{agent.Original성명}' | 새사번='{agent.성명}' | 이름='{agent.성명}'");

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
        UPDATE ""Agent"" SET 
            성명=@성명, 직급=@직급, 직무=@직무, 입사일=@입사일,
            사번=@사번, 자격사항=@자격사항, Email=@Email,
            기타=@기타, 측정인고유번호=@측정인고유번호
        WHERE 사번=@Original사번";

        cmd.Parameters.AddWithValue("@성명", agent.성명 ?? "");
        cmd.Parameters.AddWithValue("@직급", agent.직급 ?? "");
        cmd.Parameters.AddWithValue("@직무", agent.직무 ?? "");
        cmd.Parameters.AddWithValue("@입사일", agent.입사일 == DateOnly.MinValue ? DBNull.Value : agent.입사일.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@사번", agent.사번 ?? "");
        cmd.Parameters.AddWithValue("@자격사항", agent.자격사항 ?? "");
        cmd.Parameters.AddWithValue("@Email", agent.Email ?? "");
        cmd.Parameters.AddWithValue("@기타", agent.기타 ?? "");
        cmd.Parameters.AddWithValue("@측정인고유번호", agent.측정인고유번호 ?? "");
        

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE 결과] {rows}행 영향받음 (1이면 성공, 0이면 실패)");

        if (rows > 0)
        {
            agent.Original성명 = agent.성명;
            Debug.WriteLine($"✅ DB 저장 성공: {agent.성명}");
            return true;
        }
        else
        {
            Debug.WriteLine($"❌ 실패 원인: 성명 '{agent.Original성명}'인 레코드가 DB에 없음");
            return false;
        }
    }

}