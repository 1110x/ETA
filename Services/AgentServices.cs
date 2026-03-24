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

    // ── PhotoPath 컬럼 자동 마이그레이션 ─────────────────────────────────────
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

    // ── 계정 관련 컬럼 마이그레이션 (로그인 기능용) ──────────────────────────
    public static void MigrateAccountColumns()
    {
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return;

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        var columns = new[]
        {
            ("비밀번호",     "TEXT DEFAULT ''"),
            ("상태",         "TEXT DEFAULT 'approved'"),  // 기존 직원은 바로 approved
            ("todo_task_id", "TEXT DEFAULT ''"),
        };

        foreach (var (col, def) in columns)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"ALTER TABLE ""Agent"" ADD COLUMN ""{col}"" {def}";
                cmd.ExecuteNonQuery();
                Debug.WriteLine($"[AgentService] 컬럼 추가: {col}");
            }
            catch { /* 이미 존재하면 무시 */ }
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

        if (rows > 0 && !string.IsNullOrEmpty(agent.PhotoPath) && File.Exists(agent.PhotoPath))
            try { File.Delete(agent.PhotoPath); } catch { }

        return rows > 0;
    }

    // =========================================================================
    // 로그인 / 회원가입 관련
    // =========================================================================

    // ── 회원가입 (신규 직원, 상태=pending) ───────────────────────────────────
    public static bool SignUp(string name, string employeeId, string password, string department)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        EnsurePhotoPathColumn(conn);

        // 중복 사번 확인
        using var chk = conn.CreateCommand();
        chk.CommandText = @"SELECT COUNT(*) FROM ""Agent"" WHERE 사번=@id";
        chk.Parameters.AddWithValue("@id", employeeId);
        if (Convert.ToInt32(chk.ExecuteScalar()) > 0) return false;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO ""Agent"" (성명, 사번, 비밀번호, 직무, 상태)
            VALUES (@name, @id, @pw, @dept, 'pending')";
        cmd.Parameters.AddWithValue("@name", name);
        cmd.Parameters.AddWithValue("@id",   employeeId);
        cmd.Parameters.AddWithValue("@pw",   HashPassword(password));
        cmd.Parameters.AddWithValue("@dept", department ?? "");
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 로그인 검증 ───────────────────────────────────────────────────────────
    public static (bool success, string message) ValidateLogin(string employeeId, string password)
    {
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return (false, "등록된 계정이 없습니다.");

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT 비밀번호, 상태 FROM ""Agent"" WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", employeeId);

        using var r = cmd.ExecuteReader();
        if (!r.Read()) return (false, "등록되지 않은 사번입니다.");

        var dbPw   = r.IsDBNull(0) ? "" : r.GetString(0);
        var status = r.IsDBNull(1) ? "" : r.GetString(1);

        if (status == "pending")
            return (false, "관리자 승인 대기 중입니다.\nMicrosoft To Do에서 승인 후 로그인 가능합니다.");
        if (status == "rejected")
            return (false, "승인이 거부된 계정입니다.");
        if (status != "approved")
            return (false, "비활성 계정입니다.");

        if (dbPw != HashPassword(password))
            return (false, "비밀번호가 올바르지 않습니다.");

        return (true, "로그인 성공");
    }

    // ── To Do 태스크 ID 저장 ──────────────────────────────────────────────────
    public static void SaveTodoTaskId(string employeeId, string taskId)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE ""Agent"" SET todo_task_id=@tid WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@tid", taskId);
        cmd.Parameters.AddWithValue("@id",  employeeId);
        cmd.ExecuteNonQuery();
    }

    // ── 계정 승인 (TodoService 에서 호출) ─────────────────────────────────────
    public static void ApproveAccount(string employeeId)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE ""Agent"" SET 상태='approved'
            WHERE 사번=@id AND 상태='pending'";
        cmd.Parameters.AddWithValue("@id", employeeId);
        cmd.ExecuteNonQuery();
        Debug.WriteLine($"[AgentService] 승인 완료: {employeeId}");
    }

    // ── 승인된 직원 이름 목록 (LoginWindow 용) ────────────────────────────────
    public static List<string> GetAllNames()
    {
        var list   = new List<string>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return list;

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명 FROM ""Agent""
            WHERE 상태='approved' OR 상태='' OR 상태 IS NULL
            ORDER BY 성명";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 비밀번호 해시 (SHA256) ─────────────────────────────────────────────────
    public static string HashPassword(string password)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));
        return Convert.ToHexString(bytes).ToLower();
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