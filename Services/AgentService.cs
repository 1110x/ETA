using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.Common;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class AgentService
{
    // ── DB / 사진 경로 ────────────────────────────────────────────────────────
    public static string GetPhotoDirectory() => DbPathHelper.PhotoDirectory;

    // ── PhotoPath 컬럼 자동 마이그레이션 ─────────────────────────────────────
    private static void EnsurePhotoPathColumn(DbConnection conn)
    {
        if (!DbConnectionFactory.ColumnExists(conn, "Agent", "PhotoPath"))
        {
            using var alter = conn.CreateCommand();
            alter.CommandText = @"ALTER TABLE `Agent` ADD COLUMN PhotoPath TEXT DEFAULT ''";
            alter.ExecuteNonQuery();
            Debug.WriteLine("[DB] PhotoPath 컬럼 추가");
        }
    }

    // ── 계정 관련 컬럼 마이그레이션 (로그인 기능용) ──────────────────────────
    public static void MigrateAccountColumns()
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var columns = new[]
        {
            ("비밀번호",        "TEXT DEFAULT ''"),
            ("상태",            "TEXT DEFAULT 'approved'"),   // 기존 직원은 바로 approved
            ("todo_task_id",    "TEXT DEFAULT ''"),
            ("must_change_pw",  "INTEGER DEFAULT 0"),         // 1=최초 비밀번호 강제 변경 필요
            ("측정인LoginId",   "TEXT DEFAULT ''"),           // 측정인.kr 로그인 아이디
            ("측정인LoginPw",   "TEXT DEFAULT ''"),           // 측정인.kr 로그인 비밀번호
        };

        foreach (var (col, def) in columns)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"ALTER TABLE `Agent` ADD COLUMN `{col}` {def}";
                cmd.ExecuteNonQuery();
                Debug.WriteLine($"[AgentService] 컬럼 추가: {col}");
            }
            catch { /* 이미 존재하면 무시 */ }
        }
    }

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<Agent> GetAllItems()
    {
        var items  = new List<Agent>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Debug.WriteLine("❌ DB 없음"); return items; }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명, 직급, 직무, 입사일, 사번,
                   자격사항, Email, 기타, 측정인고유번호, PhotoPath
            FROM `Agent`
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
                // 절대경로로 저장된 경우 파일명만 추출 (정규화)
                PhotoPath      = NormalizePhotoPath(S(reader, "PhotoPath")),
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

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `Agent` SET
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
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsurePhotoPathColumn(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `Agent`
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

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM `Agent` WHERE 성명=@성명";
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
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsurePhotoPathColumn(conn);

        // 중복 사번 확인
        using var chk = conn.CreateCommand();
        chk.CommandText = @"SELECT COUNT(*) FROM `Agent` WHERE 사번=@id";
        chk.Parameters.AddWithValue("@id", employeeId);
        if (Convert.ToInt32(chk.ExecuteScalar()) > 0) return false;

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `Agent` (성명, 사번, 비밀번호, 직무, 상태)
            VALUES (@name, @id, @pw, @dept, 'pending')";
        cmd.Parameters.AddWithValue("@name", name);
        cmd.Parameters.AddWithValue("@id",   employeeId);
        cmd.Parameters.AddWithValue("@pw",   HashPassword(password));
        cmd.Parameters.AddWithValue("@dept", department ?? "");
        return cmd.ExecuteNonQuery() > 0;
    }

    // ── 로그인 검증 ───────────────────────────────────────────────────────────
    /// <summary>
    /// 반환: (success, message, mustChangePw)
    ///   mustChangePw=true → 로그인 직후 강제 비밀번호 변경 팝업 표시
    /// </summary>
    public static (bool success, string message, bool mustChangePw) ValidateLogin(
        string employeeId, string password)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return (false, "등록된 계정이 없습니다.", false);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT 비밀번호, 상태, COALESCE(must_change_pw,0)
                            FROM `Agent` WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", employeeId);

        using var r = cmd.ExecuteReader();
        if (!r.Read()) return (false, "등록되지 않은 사번입니다.", false);

        var dbPw         = r.IsDBNull(0) ? "" : r.GetString(0);
        var status       = r.IsDBNull(1) ? "" : r.GetString(1);
        var mustChangePw = !r.IsDBNull(2) && r.GetInt64(2) != 0;

        Debug.WriteLine($"[ValidateLogin] 사번={employeeId} 상태={status} " +
                        $"비밀번호설정={!string.IsNullOrEmpty(dbPw)} mustChange={mustChangePw}");

        // ★ 승인 상태 차단
        // 빈 문자열/NULL → 기존 직원(마이그레이션 전)으로 간주, approved 처리
        switch (status)
        {
            case "pending":
                return (false, "관리자 승인 대기 중입니다.\nMicrosoft To Do에서 승인 후 로그인 가능합니다.", false);
            case "rejected":
                return (false, "승인이 거부된 계정입니다.", false);
            case "approved":
            case "":        // 기존 직원: 상태 컬럼 없거나 빈 값
            case null:
                break;      // 아래로 계속
            default:
                return (false, $"비활성 계정입니다. (상태: {status})", false);
        }

        // 비밀번호 미설정 → MigrateInitialPasswords()가 아직 안 돌았을 때 방어
        if (string.IsNullOrEmpty(dbPw))
        {
            // 이 시점에서 직접 123456 해시로 세팅 후 mustChangePw=true 반환
            string initialHash = HashPassword("123456");
            using var fix = conn.CreateCommand();
            fix.CommandText = @"UPDATE `Agent` SET 비밀번호=@pw, must_change_pw=1 WHERE 사번=@id";
            fix.Parameters.AddWithValue("@pw", initialHash);
            fix.Parameters.AddWithValue("@id", employeeId);
            fix.ExecuteNonQuery();
            dbPw         = initialHash;
            mustChangePw = true;
            Debug.WriteLine($"[ValidateLogin] 비밀번호 없음 → 123456 자동 설정: {employeeId}");
        }

        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "eta_sync.log");
        void Log(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
            Debug.WriteLine(line);
            try { File.AppendAllText(logPath, line + "\n"); } catch { }
        }

        var inputHash = HashPassword(password);
        Log($"[ValidateLogin] DB해시={dbPw[..8]}... / 입력해시={inputHash[..8]}... / 일치={dbPw == inputHash}");

        if (dbPw != inputHash)
            return (false, "비밀번호가 올바르지 않습니다.", false);

        return (true, "로그인 성공", mustChangePw);
    }

    // ── To Do 태스크 ID 저장 ──────────────────────────────────────────────────
    public static void SaveTodoTaskId(string employeeId, string taskId)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `Agent` SET todo_task_id=@tid WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@tid", taskId);
        cmd.Parameters.AddWithValue("@id",  employeeId);
        cmd.ExecuteNonQuery();

        Debug.WriteLine($"[AgentService] todo_task_id 저장: {employeeId} → '{taskId}'");
    }

    // ── 비밀번호 규칙 검사 ────────────────────────────────────────────────────
    /// <summary>
    /// 규칙: 최소 6자 이상. 대문자·특수문자는 선택사항.
    /// </summary>
    public static (bool ok, string message) ValidatePasswordRule(string password)
    {
        if (string.IsNullOrEmpty(password))
            return (false, "새 비밀번호를 입력해주세요.");
        if (password.Length < 6)
            return (false, "비밀번호는 최소 6자 이상이어야 합니다.");
        return (true, "");
    }

    // ── 비밀번호 변경 ─────────────────────────────────────────────────────────
    /// <summary>
    /// 현재 비밀번호 확인 후 새 비밀번호로 변경.
    /// 성공 시 must_change_pw=0 으로 해제.
    /// </summary>
    public static (bool success, string message) ChangePassword(
        string employeeId, string currentPassword, string newPassword)
    {
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "eta_sync.log");
        void Log(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
            Debug.WriteLine(line);
            try { File.AppendAllText(logPath, line + "\n"); } catch { }
        }

        try
        {
            Log($"[ChangePassword] 시작 - 사번={employeeId}");

            Log($"[ChangePassword] DB 경로={DbPathHelper.DbPath}");

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            Log("[ChangePassword] DB 연결 성공");

            using var chk = conn.CreateCommand();
            chk.CommandText = @"SELECT 비밀번호 FROM `Agent` WHERE 사번=@id";
            chk.Parameters.AddWithValue("@id", employeeId);
            var stored = chk.ExecuteScalar()?.ToString() ?? "";
            Log($"[ChangePassword] DB 저장 해시={stored[..Math.Min(8,stored.Length)]}..., 입력 해시={HashPassword(currentPassword)[..8]}...");

            if (stored != HashPassword(currentPassword))
            {
                Log("[ChangePassword] 현재 비밀번호 불일치");
                return (false, "현재 비밀번호가 올바르지 않습니다.");
            }

            var (ok, msg) = ValidatePasswordRule(newPassword);
            Log($"[ChangePassword] 새 비밀번호 규칙 검사: ok={ok}");
            if (!ok) return (false, msg);

            using var upd = conn.CreateCommand();
            upd.CommandText = @"
                UPDATE `Agent`
                SET 비밀번호=@pw, must_change_pw=0
                WHERE 사번=@id";
            upd.Parameters.AddWithValue("@pw", HashPassword(newPassword));
            upd.Parameters.AddWithValue("@id", employeeId);
            int rows = upd.ExecuteNonQuery();
            Log($"[ChangePassword] UPDATE 완료 - {rows}행 변경됨");

            return (true, "비밀번호가 변경되었습니다.");
        }
        catch (Exception ex)
        {
            Log($"[ChangePassword] 예외 발생: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            return (false, $"오류: {ex.Message}");
        }
    }

    // ── 기존 Agent 초기 비밀번호 일괄 설정 (마이그레이션) ────────────────────
    /// <summary>
    /// 비밀번호가 없는 기존 직원에게 초기값 '123456' 부여 + must_change_pw=1 설정.
    /// MigrateAccountColumns() 이후에 호출한다.
    /// </summary>
    public static void MigrateInitialPasswords()
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // must_change_pw 컬럼 없으면 추가 (이중 안전망)
        try
        {
            using var alt = conn.CreateCommand();
            alt.CommandText = @"ALTER TABLE `Agent` ADD COLUMN must_change_pw INTEGER DEFAULT 0";
            alt.ExecuteNonQuery();
        }
        catch { }

        string initialHash = HashPassword("123456");

        using var cmd = conn.CreateCommand();
        // 비밀번호가 비어있는 approved 직원에게만 초기값 세팅
        cmd.CommandText = @"
            UPDATE `Agent`
            SET 비밀번호 = @pw,
                must_change_pw = 1
            WHERE (비밀번호 IS NULL OR 비밀번호 = '')
              AND (상태 = 'approved' OR 상태 = '' OR 상태 IS NULL)";
        cmd.Parameters.AddWithValue("@pw", initialHash);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[AgentService] 초기 비밀번호 설정 완료: {rows}명");
    }

    // ── 최초 비밀번호 변경 필요 여부 조회 ────────────────────────────────────
    public static bool NeedsPasswordChange(string employeeId)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return false;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT COALESCE(must_change_pw,0) FROM `Agent` WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", employeeId);
        var val = cmd.ExecuteScalar();
        return val != null && Convert.ToInt64(val) != 0;
    }

    // ── 계정 승인 (TodoService에서 호출) ─────────────────────────────────────
    public static void ApproveAccount(string employeeId)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        // pending뿐 아니라 어떤 상태든 approved로 갱신 (재승인 허용)
        cmd.CommandText = @"UPDATE `Agent` SET 상태='approved' WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", employeeId);
        cmd.ExecuteNonQuery();
        Debug.WriteLine($"[AgentService] 승인 완료: {employeeId}");
    }

    // ── 알 수 없는 해시가 저장된 계정 전체를 123456으로 초기화 ───────────────
    /// <summary>
    /// 비밀번호가 비어있지도 않고, 123456 해시도 아닌 계정을 모두 초기화.
    /// 앱 배포 후 1회만 실행하고 이후 호출부를 제거할 것.
    /// </summary>
    public static void ResetAllUnknownPasswords()
    {
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "eta_sync.log");
        void Log(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
            Debug.WriteLine(line);
            try { File.AppendAllText(logPath, line + "\n"); } catch { }
        }

        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        string hash123456 = HashPassword("123456");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        // 비어있지도 않고, 123456 해시도 아닌 계정만 대상
        cmd.CommandText = @"
            UPDATE `Agent`
            SET 비밀번호 = @pw,
                must_change_pw = 1
            WHERE 비밀번호 IS NOT NULL
              AND 비밀번호 != ''
              AND 비밀번호 != @pw";
        cmd.Parameters.AddWithValue("@pw", hash123456);
        int rows = cmd.ExecuteNonQuery();
        Log($"[ResetAllUnknownPasswords] 알 수 없는 비밀번호 → 123456 초기화: {rows}명");
    }

    // ── 특정 사번 비밀번호를 123456으로 강제 초기화 ──────────────────────────
    public static void ResetPassword(string employeeId)
    {
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "eta_sync.log");
        void Log(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
            Debug.WriteLine(line);
            try { File.AppendAllText(logPath, line + "\n"); } catch { }
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `Agent` SET 비밀번호=@pw, must_change_pw=1 WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@pw", HashPassword("123456"));
        cmd.Parameters.AddWithValue("@id", employeeId);
        int rows = cmd.ExecuteNonQuery();
        Log($"[ResetPassword] 사번={employeeId} → 123456으로 초기화 완료 ({rows}행)");
    }

    // ── 승인 취소 (To Do 미완료/삭제 감지 시 TodoService에서 호출) ────────────
    /// <summary>
    /// approved → pending 으로 되돌린다. 로그인이 즉시 차단된다.
    /// </summary>
    public static void ResetToPending(string employeeId)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        // approved 뿐 아니라 모든 상태에서 pending으로 (방어적 처리)
        cmd.CommandText = @"UPDATE `Agent` SET 상태='pending' WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", employeeId);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[AgentService] 승인 취소 → pending: {employeeId} ({rows}행)");
    }

    // ── 승인 관련 모든 계정 조회 (DB가 Master) ───────────────────────────────
    /// <summary>
    /// 상태가 pending 또는 approved인 계정을 모두 반환한다.
    /// (rejected 계정은 동기화 대상 제외)
    /// </summary>
    public static List<(string 성명, string 사번, string 상태, string todo_task_id)> GetAllApprovalAccounts()
    {
        var list   = new List<(string, string, string, string)>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명, 사번, 상태, COALESCE(todo_task_id, '') AS todo_task_id
            FROM `Agent`
            WHERE 상태 IN ('approved', 'pending')";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add((
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3)
            ));
        }
        return list;
    }

    // ── pending 계정 중 todo_task_id 없는 것 조회 ────────────────────────────
    public static List<Agent> GetPendingAccountsWithoutTask()
    {
        var list   = new List<Agent>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명, 사번 FROM `Agent`
            WHERE 상태='pending' AND (todo_task_id IS NULL OR todo_task_id = '')";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new Agent { 성명 = r.GetString(0), 사번 = r.GetString(1) });
        return list;
    }

    // ── 승인된 직원 이름 목록 (LoginWindow 용) ────────────────────────────────
    public static List<string> GetAllNames()
    {
        var list   = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 성명 FROM `Agent`
            WHERE 상태='approved' OR 상태='' OR 상태 IS NULL
            ORDER BY 성명";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 비밀번호 해시 (SHA256) ─────────────────────────────────────────────────
    public static string HashPassword(string password)
    {
        using var sha   = System.Security.Cryptography.SHA256.Create();
        var       bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));
        return Convert.ToHexString(bytes).ToLower();
    }

    // ── 디버그: 계정 상태 확인 ───────────────────────────────────────────────
    public static void DebugCheckAccount(string empId)
    {
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "eta_sync.log");
        void Log(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";
            Debug.WriteLine(line);
            try { File.AppendAllText(logPath, line + "\n"); } catch { }
        }

        Log($"[DebugCheck] DB 경로: {DbPathHelper.DbPath}");
        Log($"[DebugCheck] DB 존재: {File.Exists(DbPathHelper.DbPath)}");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT 상태, todo_task_id, 비밀번호, COALESCE(must_change_pw,0)
                            FROM `Agent` WHERE 사번=@id";
        cmd.Parameters.AddWithValue("@id", empId);

        using var r = cmd.ExecuteReader();
        if (r.Read())
        {
            var state    = r.IsDBNull(0) ? "없음" : r.GetString(0);
            var taskId   = r.IsDBNull(1) ? ""     : r.GetString(1);
            var pw       = r.IsDBNull(2) ? ""     : r.GetString(2);
            var mustChg  = r.IsDBNull(3) ? 0      : r.GetInt64(3);
            var pwInfo   = string.IsNullOrEmpty(pw) ? "비밀번호없음" : $"해시={pw[..Math.Min(8,pw.Length)]}...";
            Log($"[DebugCheck] 사번={empId} 상태={state} must_change={mustChg} {pwInfo} taskId={taskId}");
            Log($"[DebugCheck] 123456 해시={HashPassword("123456")[..8]}...");
        }
        else
        {
            Log($"[DebugCheck] {empId} → DB에 레코드 없음!");
        }
    }

    // ── 공통 파라미터 세팅 ────────────────────────────────────────────────────
    private static void SetParams(DbCommand cmd, Agent a)
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
    private static string S(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) ?? "" : ""; }
        catch { return ""; }
    }

    // 절대경로로 저장된 PhotoPath를 파일명만으로 정규화
    // "C:\...\Data\Photos\201000308.jpg" → "201000308.jpg"
    private static string NormalizePhotoPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return "";
        // 이미 파일명만 있으면 그대로
        return Path.IsPathRooted(path) ? Path.GetFileName(path) : path;
    }

    // ── 측정인.kr 자격증명 조회 / 저장 ──────────────────────────────────────
    /// <summary>사번으로 측정인 로그인 ID/PW 조회. 없으면 ("","") 반환.</summary>
    public static (string Id, string Pw) GetMeasurerCredentials(string 사번)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return ("", "");

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT 측정인LoginId, 측정인LoginPw FROM `Agent` WHERE 사번=@사번 LIMIT 1";
        cmd.Parameters.AddWithValue("@사번", 사번);

        using var r = cmd.ExecuteReader();
        if (!r.Read()) return ("", "");

        string id = S(r, "측정인LoginId");
        string pw = S(r, "측정인LoginPw");
        return (id, pw);
    }

    /// <summary>사번에 해당하는 행에 측정인 로그인 ID/PW 저장.</summary>
    public static void SaveMeasurerCredentials(string 사번, string id, string pw)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "UPDATE `Agent` SET 측정인LoginId=@id, 측정인LoginPw=@pw WHERE 사번=@사번";
        cmd.Parameters.AddWithValue("@id",  id  ?? "");
        cmd.Parameters.AddWithValue("@pw",  pw  ?? "");
        cmd.Parameters.AddWithValue("@사번", 사번 ?? "");
        cmd.ExecuteNonQuery();
    }

    // ── 관리자 권한 확인 ─────────────────────────────────────────────────────
    public static bool IsAdmin(string employeeId)
    {
        // 간단히 사번이 "admin"이거나 특정 사번이면 관리자
        return employeeId == "admin" || employeeId == "202912345"; // 예시
    }

    // ── 측정인.kr 인력 고유번호 일괄 업데이트 ───────────────────────────────
    /// <summary>
    /// 측정인.kr add_emp_id 드롭다운에서 수집한 (이름, 고유번호) 쌍으로
    /// Agent 테이블의 측정인고유번호를 업데이트.
    /// 이름 완전일치 → 앞 3글자+길이 유사 유일 후보 순으로 매칭.
    /// </summary>
    /// <returns>업데이트된 행 수</returns>
    public static int UpdateMeasurerEmployeeIds(List<(string Name, string Id)> pairs)
    {
        int count = 0;
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return 0;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // Agent 이름 전체 로드 (메모리에서 매칭 — 수십 명 규모이므로 충분)
        var agentNames = new List<string>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"SELECT 성명 FROM `Agent`";
            using var r = cmd.ExecuteReader();
            while (r.Read()) agentNames.Add(r.GetString(0));
        }

        foreach (var (name, id) in pairs)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(id)) continue;

            // 1) 완전 일치
            string? matched = agentNames.FirstOrDefault(a =>
                string.Equals(a, name, StringComparison.OrdinalIgnoreCase));

            // 2) 앞 3글자 포함 + 길이 유사 (오차 ±1) → 후보가 1명일 때만
            if (matched == null && name.Length >= 2)
            {
                string prefix = name[..Math.Min(3, name.Length)];
                var candidates = agentNames
                    .Where(a => a.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                             && Math.Abs(a.Length - name.Length) <= 1)
                    .ToList();
                if (candidates.Count == 1) matched = candidates[0];
            }

            if (matched == null) continue;

            using var upd = conn.CreateCommand();
            upd.CommandText = @"UPDATE `Agent` SET 측정인고유번호=@id WHERE 성명=@name";
            upd.Parameters.AddWithValue("@id",   id);
            upd.Parameters.AddWithValue("@name", matched);
            count += upd.ExecuteNonQuery();
        }

        Debug.WriteLine($"[AgentService] UpdateMeasurerEmployeeIds: {pairs.Count}개 수집 → {count}명 업데이트");
        return count;
    }
}
