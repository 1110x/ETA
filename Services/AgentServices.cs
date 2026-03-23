using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class AgentService
{
    // ── DB / 사진 경로 ────────────────────────────────────────────────────────
    public static string GetDatabasePath()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dir  = Path.Combine(root, "Data");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "eta.db");
    }

    public static string GetPhotoDirectory()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dir  = Path.Combine(root, "Data", "Photos");
        Directory.CreateDirectory(dir);
        return dir;
    }

    // ── PhotoPath 컬럼 자동 마이그레이션 (기존 DB 호환) ──────────────────────
    private static void EnsurePhotoPathColumn(SqliteConnection conn)
    {
        using var check = conn.CreateCommand();
        check.CommandText = @"PRAGMA table_info(""Agent"")";
        using var r = check.ExecuteReader();
        bool exists = false;
        while (r.Read())
            if (r.GetString(1) == "PhotoPath") { exists = true; break; }

        if (!exists)
        {
            using var alter = conn.CreateCommand();
            alter.CommandText = @"ALTER TABLE ""Agent"" ADD COLUMN PhotoPath TEXT DEFAULT ''";
            alter.ExecuteNonQuery();
            Debug.WriteLine("[DB] PhotoPath 컬럼 추가");
        }
    }

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<Agent> GetAllItems()
    {
        var items = new List<Agent>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) { Debug.WriteLine("❌ DB 없음"); return items; }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명, 직급, 직무, 입사일, 사번,
                   자격사항, Email, 기타, 측정인고유번호, PhotoPath
            FROM ""Agent""
            ORDER BY 사번 ASC";

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var a = new Agent
            {
                성명           = S(reader, "성명"),
                Original성명   = S(reader, "성명"),
                직급           = S(reader, "직급"),
                직무           = S(reader, "직무"),
                사번           = S(reader, "사번"),
                자격사항       = S(reader, "자격사항"),
                Email          = S(reader, "Email"),
                기타           = S(reader, "기타"),
                측정인고유번호 = S(reader, "측정인고유번호"),
                PhotoPath      = S(reader, "PhotoPath"),
            };
            if (DateOnly.TryParse(S(reader, "입사일"), out var d)) a.입사일 = d;
            items.Add(a);
        }

        Debug.WriteLine($"[DB] 로드 {items.Count}명");
        return items;
    }

    // ── 수정 ─────────────────────────────────────────────────────────────────
    public static bool Update(Agent agent)
    {
        if (string.IsNullOrEmpty(agent.Original성명)) { Debug.WriteLine("❌ Original성명 없음"); return false; }

        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE ""Agent"" SET
                성명=@성명, 직급=@직급, 직무=@직무, 입사일=@입사일,
                사번=@사번, 자격사항=@자격사항, Email=@Email,
                기타=@기타, 측정인고유번호=@측정인고유번호, PhotoPath=@PhotoPath
            WHERE 성명=@Original성명";

        SetParams(cmd, agent);
        cmd.Parameters.AddWithValue("@Original성명", agent.Original성명);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE] {rows}행 → {agent.성명}");
        if (rows > 0) { agent.Original성명 = agent.성명; return true; }
        return false;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────────
    public static bool Insert(Agent agent)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO ""Agent""
                (성명, 직급, 직무, 입사일, 사번, 자격사항, Email, 기타, 측정인고유번호, PhotoPath)
            VALUES
                (@성명, @직급, @직무, @입사일, @사번, @자격사항, @Email, @기타, @측정인고유번호, @PhotoPath)";

        SetParams(cmd, agent);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[INSERT] {rows}행 → {agent.성명}");
        return rows > 0;
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────────
    public static bool Delete(Agent agent)
    {
        if (string.IsNullOrEmpty(agent.성명)) { Debug.WriteLine("❌ 삭제 대상 없음"); return false; }

        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM ""Agent"" WHERE 성명=@성명";
        cmd.Parameters.AddWithValue("@성명", agent.성명);

        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[DELETE] {rows}행 → {agent.성명}");

        // 사진 파일도 함께 삭제
        if (rows > 0 && !string.IsNullOrEmpty(agent.PhotoPath) && File.Exists(agent.PhotoPath))
            try { File.Delete(agent.PhotoPath); } catch { }

        return rows > 0;
    }

    // ── 공통 파라미터 세팅 ────────────────────────────────────────────────────
    private static void SetParams(SqliteCommand cmd, Agent a)
    {
        cmd.Parameters.AddWithValue("@성명",           a.성명           ?? "");
        cmd.Parameters.AddWithValue("@직급",           a.직급           ?? "");
        cmd.Parameters.AddWithValue("@직무",           a.직무           ?? "");
        cmd.Parameters.AddWithValue("@입사일",         a.입사일 == DateOnly.MinValue
                                                            ? DBNull.Value
                                                            : (object)a.입사일.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@사번",           a.사번           ?? "");
        cmd.Parameters.AddWithValue("@자격사항",       a.자격사항       ?? "");
        cmd.Parameters.AddWithValue("@Email",          a.Email          ?? "");
        cmd.Parameters.AddWithValue("@기타",           a.기타           ?? "");
        cmd.Parameters.AddWithValue("@측정인고유번호", a.측정인고유번호 ?? "");
        cmd.Parameters.AddWithValue("@PhotoPath",      a.PhotoPath      ?? "");
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────
    private static string S(SqliteDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) ?? "" : ""; }
        catch { return ""; }
    }
}