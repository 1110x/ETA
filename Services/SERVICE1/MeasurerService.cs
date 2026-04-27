using System.Data;
using System.Data.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class MeasurerService
{
    // ── 메모리 캐시 ──────────────────────────────────────────────────────────
    private static List<string>? _companiesCache;
    private static List<(string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명)>? _allDataCache;
    private static List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)>? _analysisItemsCache;
    public static void InvalidateCache() { _companiesCache = null; _allDataCache = null; }
    public static void InvalidateAnalysisItemsCache() => _analysisItemsCache = null;

    // ── 테이블 생성 + 컬럼 마이그레이션 ─────────────────────────────────────
    public static void EnsureMeasurerTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 기본 테이블 생성
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS 측정인_채취지점 (
                    Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    계약번호   TEXT,
                    약칭       TEXT,
                    계약기간   TEXT,
                    업체명     TEXT NOT NULL,
                    채취지점명 TEXT NOT NULL,
                    UNIQUE(업체명, 채취지점명)
                );";
            cmd.ExecuteNonQuery();
        }

        // 기존 DB 마이그레이션 — 컬럼이 없으면 추가
        foreach (var col in new[] { ("계약번호", "TEXT"), ("약칭", "TEXT"), ("계약기간", "TEXT"),
                                    ("계약금액", "TEXT"), ("견적구분", "TEXT") })
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE 측정인_채취지점 ADD COLUMN {col.Item1} {col.Item2}";
                cmd.ExecuteNonQuery();
            }
            catch { /* 이미 존재하면 무시 */ }
        }
    }

    // ── 분석장비 테이블 생성 / 마이그레이션 ─────────────────────────────────
    public static void EnsureEquipmentTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS 측정인_분석장비 (
                    Id       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    장비명   TEXT NOT NULL,
                    코드값   VARCHAR(191) NOT NULL,
                    UNIQUE(코드값)
                );";
            cmd.ExecuteNonQuery();
        }
        // 마이그레이션: 현장측정 플래그 (0/1). 이미 있으면 무시
        try
        {
            using var alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE `측정인_분석장비` ADD COLUMN `현장측정` INTEGER NOT NULL DEFAULT 0";
            alter.ExecuteNonQuery();
        }
        catch { /* 컬럼이 이미 존재 */ }
    }

    // ── 분석장비 일괄 저장 (UPSERT) ─────────────────────────────────────────
    public static int SaveEquipments(List<(string 장비명, string 코드값)> items)
    {
        EnsureEquipmentTable();
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var (name, code) in items)
            {
                if (string.IsNullOrWhiteSpace(code)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(new[] { "코드값" }, new[] { "장비명" });
                cmd.CommandText = $"INSERT INTO 측정인_분석장비 (장비명, 코드값) VALUES (@name, @code) {upsert};";
                cmd.Parameters.AddWithValue("@name", name.Trim());
                cmd.Parameters.AddWithValue("@code", code.Trim());
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch (Exception ex) { txn.Rollback(); throw; }
        return count;
    }

    // ── 분석장비 전체 조회 ───────────────────────────────────────────────────
    public static List<(string 장비명, string 코드값)> GetAllEquipments()
    {
        EnsureEquipmentTable();
        var list = new List<(string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 장비명, 코드값 FROM 측정인_분석장비 ORDER BY 장비명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1)));
        return list;
    }

    // -- 분석장비 단건 저장 (UPSERT)
    public static void SaveEquipment(string 장비명, string 코드값, string? old코드값 = null)
    {
        EnsureEquipmentTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!string.IsNullOrWhiteSpace(old코드값) && old코드값 != 코드값)
        {
            using var del = conn.CreateCommand();
            del.CommandText = "DELETE FROM `측정인_분석장비` WHERE `코드값` = @old";
            del.Parameters.AddWithValue("@old", old코드값);
            del.ExecuteNonQuery();
        }
        var upsert = DbConnectionFactory.UpsertSuffix(new[] { "코드값" }, new[] { "장비명" });
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"INSERT INTO `측정인_분석장비` (장비명, 코드값) VALUES (@name, @code) {upsert}";
        cmd.Parameters.AddWithValue("@name", 장비명.Trim());
        cmd.Parameters.AddWithValue("@code", 코드값.Trim());
        cmd.ExecuteNonQuery();
    }

    // -- 분석장비 삭제
    public static void DeleteEquipment(string 코드값)
    {
        EnsureEquipmentTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `측정인_분석장비` WHERE `코드값` = @code";
        cmd.Parameters.AddWithValue("@code", 코드값);
        cmd.ExecuteNonQuery();
    }

    // ── 현장측정장비 플래그 ─────────────────────────────────────────────────
    public static bool GetFieldMeasFlag(string 코드값)
    {
        EnsureEquipmentTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `현장측정` FROM `측정인_분석장비` WHERE `코드값` = @c";
        cmd.Parameters.AddWithValue("@c", 코드값.Trim());
        var r = cmd.ExecuteScalar();
        return r != null && r != DBNull.Value && Convert.ToInt32(r) == 1;
    }

    public static void SetFieldMeasFlag(string 코드값, bool on)
    {
        EnsureEquipmentTable();
        var code = 코드값.Trim();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // UPSERT: 존재하면 현장측정만 갱신, 없으면 장비명=코드값으로 INSERT
        var upsert = DbConnectionFactory.UpsertSuffix(new[] { "코드값" }, new[] { "현장측정" });
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"INSERT INTO `측정인_분석장비` (`장비명`, `코드값`, `현장측정`) VALUES (@n, @c, @v) {upsert}";
        cmd.Parameters.AddWithValue("@n", code);
        cmd.Parameters.AddWithValue("@c", code);
        cmd.Parameters.AddWithValue("@v", on ? 1 : 0);
        cmd.ExecuteNonQuery();

        // 저장 결과 재조회 → 로그
        int saved = -1;
        using (var chk = conn.CreateCommand())
        {
            chk.CommandText = "SELECT `현장측정` FROM `측정인_분석장비` WHERE `코드값` = @c";
            chk.Parameters.AddWithValue("@c", code);
            var o = chk.ExecuteScalar();
            if (o != null && o != DBNull.Value) saved = Convert.ToInt32(o);
        }
        try
        {
            Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SetFieldMeasFlag] code='{code}' set={(on?1:0)} → saved={saved}\n");
        }
        catch { }
    }

    // ── 현장측정장비 (edit_meas_equip_no 드롭다운 전용 테이블) ─────────────
    public static void EnsureFieldMeasEquipmentTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `현장측정장비` (
                Id       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                `장비명` TEXT NOT NULL,
                `코드값` VARCHAR(191) NOT NULL,
                UNIQUE(`코드값`)
            );";
        cmd.ExecuteNonQuery();
    }

    public static int SaveFieldMeasEquipments(List<(string 장비명, string 코드값)> items)
    {
        EnsureFieldMeasEquipmentTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using (var del = conn.CreateCommand()) { del.CommandText = "DELETE FROM `현장측정장비`"; del.ExecuteNonQuery(); }
        int count = 0;
        foreach (var (n, c) in items)
        {
            if (string.IsNullOrWhiteSpace(c)) continue;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO `현장측정장비` (`장비명`, `코드값`) VALUES (@n, @c)";
            cmd.Parameters.AddWithValue("@n", n.Trim());
            cmd.Parameters.AddWithValue("@c", c.Trim());
            count += cmd.ExecuteNonQuery();
        }
        try
        {
            Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SaveFieldMeasEquipments] {count}건 저장\n");
        }
        catch { }
        return count;
    }

    // 전송 시: 현장측정장비 테이블의 (장비명, 코드값) 전체 조회
    public static List<(string 장비명, string 코드값)> GetFieldMeasEquipment()
    {
        EnsureFieldMeasEquipmentTable();
        var list = new List<(string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `장비명`, `코드값` FROM `현장측정장비` ORDER BY `장비명`";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add((r.GetString(0), r.GetString(1)));
        try
        {
            Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [GetFieldMeasEquipment] {list.Count}건 → [{string.Join(";", list.Select(x=>x.Item1))}]\n");
        }
        catch { }
        return list;
    }

    // 전송 시: 현장측정=1 장비 코드 전체 조회
    public static List<string> GetFieldMeasEquipmentCodes()
    {
        EnsureEquipmentTable();
        var list = new List<string>();
        int total = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT `코드값` FROM `측정인_분석장비` WHERE `현장측정` = 1";
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
        }
        using (var cmd2 = conn.CreateCommand())
        {
            cmd2.CommandText = "SELECT COUNT(*) FROM `측정인_분석장비`";
            var o = cmd2.ExecuteScalar();
            total = (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
        }
        try
        {
            Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [GetFieldMeasEquipmentCodes] 현장측정=1 {list.Count}건 / 전체 {total}건 → [{string.Join(",", list)}]\n");
        }
        catch { }
        return list;
    }

    // ── 분석항목 ↔ 장비 매핑 (다대다) ──────────────────────────────────────
    public static void EnsureAnalyteEquipmentTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `분석항목_장비` (
                Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                `분석항목` VARCHAR(191) NOT NULL,
                `장비코드` VARCHAR(191) NOT NULL,
                UNIQUE(`분석항목`, `장비코드`)
            );";
        cmd.ExecuteNonQuery();
    }

    public static List<string> GetEquipmentsForAnalyte(string 분석항목)
    {
        EnsureAnalyteEquipmentTable();
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `장비코드` FROM `분석항목_장비` WHERE `분석항목` = @a";
        cmd.Parameters.AddWithValue("@a", 분석항목.Trim());
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public static void SaveEquipmentsForAnalyte(string 분석항목, IEnumerable<string> 장비코드목록)
    {
        EnsureAnalyteEquipmentTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            using (var del = conn.CreateCommand())
            {
                del.Transaction = txn;
                del.CommandText = "DELETE FROM `분석항목_장비` WHERE `분석항목` = @a";
                del.Parameters.AddWithValue("@a", 분석항목.Trim());
                del.ExecuteNonQuery();
            }
            foreach (var code in 장비코드목록.Where(c => !string.IsNullOrWhiteSpace(c)).Distinct())
            {
                using var ins = conn.CreateCommand();
                ins.Transaction = txn;
                ins.CommandText = "INSERT INTO `분석항목_장비` (`분석항목`, `장비코드`) VALUES (@a, @c)";
                ins.Parameters.AddWithValue("@a", 분석항목.Trim());
                ins.Parameters.AddWithValue("@c", code.Trim());
                ins.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch { txn.Rollback(); throw; }
    }

    // ── 분석항목 ↔ 동시분석그룹 번호 (VBA 측정DB.AI 컬럼 모사) ───────────────
    // 값: 1~N (동일 번호 = 같이 분석되는 묶음). 0/NULL = 그룹 없음
    //  예: 금속류(구리, 납, 니켈, 망간, 바륨) → 그룹 1
    //       음이온(F, Cl, NO3, SO4) → 그룹 2
    //  측정인 결과제출 시 동일 그룹 묶어서 conc_group 지정에 사용
    public static void EnsureAnalyteConcurrentGroupTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `분석항목_동시그룹` (
                Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                `분석항목` VARCHAR(191) NOT NULL,
                `그룹번호` INTEGER NOT NULL,
                UNIQUE(`분석항목`)
            );";
        cmd.ExecuteNonQuery();
    }

    public static int GetConcurrentGroup(string 분석항목)
    {
        EnsureAnalyteConcurrentGroupTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `그룹번호` FROM `분석항목_동시그룹` WHERE `분석항목` = @a";
        cmd.Parameters.AddWithValue("@a", 분석항목.Trim());
        var r = cmd.ExecuteScalar();
        return (r == null || r == DBNull.Value) ? 0 : Convert.ToInt32(r);
    }

    public static void SaveConcurrentGroup(string 분석항목, int 그룹번호)
    {
        EnsureAnalyteConcurrentGroupTable();
        var self = 분석항목.Trim();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (그룹번호 <= 0)
        {
            using var del = conn.CreateCommand();
            del.CommandText = "DELETE FROM `분석항목_동시그룹` WHERE `분석항목` = @a";
            del.Parameters.AddWithValue("@a", self);
            del.ExecuteNonQuery();
            return;
        }
        var upsert = DbConnectionFactory.UpsertSuffix(new[] { "분석항목" }, new[] { "그룹번호" });
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"INSERT INTO `분석항목_동시그룹` (`분석항목`, `그룹번호`) VALUES (@a, @g) {upsert}";
        cmd.Parameters.AddWithValue("@a", self);
        cmd.Parameters.AddWithValue("@g", 그룹번호);
        cmd.ExecuteNonQuery();
    }

    public static Dictionary<string, int> GetAllConcurrentGroups()
    {
        EnsureAnalyteConcurrentGroupTable();
        var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `분석항목`, `그룹번호` FROM `분석항목_동시그룹`";
        using var r = cmd.ExecuteReader();
        while (r.Read()) dict[r.GetString(0)] = r.GetInt32(1);
        return dict;
    }

    // =========================================================================
    // 측정인계약 / 측정인처리시설 (add_meas_cont_no + cmb_emis_cmpy_plc_no 스크랩)
    // =========================================================================

    public static void EnsureMeasurerContractTables()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS `측정인계약` (
                    Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    `계약번호` VARCHAR(40) NOT NULL,
                    `업체명`   TEXT,
                    `대표사업장` TEXT,
                    `계약기간` TEXT,
                    `전체텍스트` TEXT,
                    `수집일시` DATETIME,
                    UNIQUE(`계약번호`)
                );";
            cmd.ExecuteNonQuery();
        }
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS `측정인처리시설` (
                    Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    `계약번호`   VARCHAR(40) NOT NULL,
                    `처리시설명` VARCHAR(200) NOT NULL,
                    `처리시설코드` TEXT,
                    `수집일시`   DATETIME,
                    UNIQUE(`계약번호`, `처리시설명`)
                );";
            cmd.ExecuteNonQuery();
        }
        using (var cmd = conn.CreateCommand())
        {
            // 측정시설 = edit_emis_fac_no (처리시설 하위: CIP 폐수, 방류수, 저류조 등)
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS `측정인측정시설` (
                    Id           INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    `계약번호`     VARCHAR(40)  NOT NULL,
                    `처리시설코드` VARCHAR(40)  NOT NULL,
                    `처리시설명`   VARCHAR(200),
                    `측정시설명`   VARCHAR(200) NOT NULL,
                    `측정시설코드` TEXT,
                    `수집일시`     DATETIME,
                    UNIQUE(`계약번호`, `처리시설코드`, `측정시설명`)
                );";
            cmd.ExecuteNonQuery();
        }
    }

    public readonly record struct MeasurerContract(
        string 계약번호, string 업체명, string 대표사업장, string 계약기간, string 전체텍스트);

    public readonly record struct MeasurerPlace(
        string 계약번호, string 처리시설명, string 처리시설코드);

    public readonly record struct MeasurerFacility(
        string 계약번호, string 처리시설코드, string 처리시설명, string 측정시설명, string 측정시설코드);

    public static int SaveMeasurerContracts(List<MeasurerContract> items)
    {
        EnsureMeasurerContractTables();
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            var upsert = DbConnectionFactory.UpsertSuffix(
                new[] { "계약번호" },
                new[] { "업체명", "대표사업장", "계약기간", "전체텍스트", "수집일시" });
            foreach (var c in items)
            {
                if (string.IsNullOrWhiteSpace(c.계약번호)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                cmd.CommandText = $@"INSERT INTO `측정인계약`
                    (`계약번호`, `업체명`, `대표사업장`, `계약기간`, `전체텍스트`, `수집일시`)
                    VALUES (@no, @co, @pl, @pd, @tx, @ts) {upsert}";
                cmd.Parameters.AddWithValue("@no", c.계약번호.Trim());
                cmd.Parameters.AddWithValue("@co", c.업체명 ?? "");
                cmd.Parameters.AddWithValue("@pl", c.대표사업장 ?? "");
                cmd.Parameters.AddWithValue("@pd", c.계약기간 ?? "");
                cmd.Parameters.AddWithValue("@tx", c.전체텍스트 ?? "");
                cmd.Parameters.AddWithValue("@ts", DateTime.Now);
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch { txn.Rollback(); throw; }
        return count;
    }

    public static int SaveMeasurerPlaces(string 계약번호, List<MeasurerPlace> items)
    {
        EnsureMeasurerContractTables();
        if (string.IsNullOrWhiteSpace(계약번호)) return 0;
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            // 해당 계약 기존 행 삭제 후 재삽입 — 측정인 사이트 변경 반영
            using (var del = conn.CreateCommand())
            {
                del.Transaction = txn;
                del.CommandText = "DELETE FROM `측정인처리시설` WHERE `계약번호` = @no";
                del.Parameters.AddWithValue("@no", 계약번호);
                del.ExecuteNonQuery();
            }
            foreach (var p in items)
            {
                if (string.IsNullOrWhiteSpace(p.처리시설명)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                cmd.CommandText = @"INSERT INTO `측정인처리시설`
                    (`계약번호`, `처리시설명`, `처리시설코드`, `수집일시`)
                    VALUES (@no, @nm, @cd, @ts)";
                cmd.Parameters.AddWithValue("@no", 계약번호);
                cmd.Parameters.AddWithValue("@nm", p.처리시설명.Trim());
                cmd.Parameters.AddWithValue("@cd", p.처리시설코드 ?? "");
                cmd.Parameters.AddWithValue("@ts", DateTime.Now);
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch { txn.Rollback(); throw; }
        return count;
    }

    public static List<MeasurerContract> GetAllMeasurerContracts()
    {
        EnsureMeasurerContractTables();
        var list = new List<MeasurerContract>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT `계약번호`, `업체명`, `대표사업장`, `계약기간`, `전체텍스트`
                            FROM `측정인계약` ORDER BY `계약번호` ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new MeasurerContract(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4)));
        return list;
    }

    public static List<MeasurerPlace> GetMeasurerPlaces(string 계약번호)
    {
        EnsureMeasurerContractTables();
        var list = new List<MeasurerPlace>();
        if (string.IsNullOrWhiteSpace(계약번호)) return list;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT `계약번호`, `처리시설명`, `처리시설코드`
                            FROM `측정인처리시설` WHERE `계약번호` = @no ORDER BY `처리시설명` ASC";
        cmd.Parameters.AddWithValue("@no", 계약번호);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new MeasurerPlace(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2)));
        return list;
    }

    public static void DeleteMeasurerContract(string 계약번호)
    {
        EnsureMeasurerContractTables();
        if (string.IsNullOrWhiteSpace(계약번호)) return;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            using (var del0 = conn.CreateCommand())
            {
                del0.Transaction = txn;
                del0.CommandText = "DELETE FROM `측정인측정시설` WHERE `계약번호` = @no";
                del0.Parameters.AddWithValue("@no", 계약번호);
                del0.ExecuteNonQuery();
            }
            using (var del1 = conn.CreateCommand())
            {
                del1.Transaction = txn;
                del1.CommandText = "DELETE FROM `측정인처리시설` WHERE `계약번호` = @no";
                del1.Parameters.AddWithValue("@no", 계약번호);
                del1.ExecuteNonQuery();
            }
            using (var del2 = conn.CreateCommand())
            {
                del2.Transaction = txn;
                del2.CommandText = "DELETE FROM `측정인계약` WHERE `계약번호` = @no";
                del2.Parameters.AddWithValue("@no", 계약번호);
                del2.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch { txn.Rollback(); throw; }
    }

    // 한 (계약, 처리시설코드) 묶음의 측정시설 목록 교체 삽입
    public static int SaveMeasurerFacilities(string 계약번호, string 처리시설코드, List<MeasurerFacility> items)
    {
        EnsureMeasurerContractTables();
        if (string.IsNullOrWhiteSpace(계약번호) || string.IsNullOrWhiteSpace(처리시설코드)) return 0;
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            using (var del = conn.CreateCommand())
            {
                del.Transaction = txn;
                del.CommandText = "DELETE FROM `측정인측정시설` WHERE `계약번호` = @no AND `처리시설코드` = @pc";
                del.Parameters.AddWithValue("@no", 계약번호);
                del.Parameters.AddWithValue("@pc", 처리시설코드);
                del.ExecuteNonQuery();
            }
            foreach (var f in items)
            {
                if (string.IsNullOrWhiteSpace(f.측정시설명)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                cmd.CommandText = @"INSERT INTO `측정인측정시설`
                    (`계약번호`, `처리시설코드`, `처리시설명`, `측정시설명`, `측정시설코드`, `수집일시`)
                    VALUES (@no, @pc, @pn, @fn, @fc, @ts)";
                cmd.Parameters.AddWithValue("@no", 계약번호);
                cmd.Parameters.AddWithValue("@pc", 처리시설코드);
                cmd.Parameters.AddWithValue("@pn", f.처리시설명 ?? "");
                cmd.Parameters.AddWithValue("@fn", f.측정시설명.Trim());
                cmd.Parameters.AddWithValue("@fc", f.측정시설코드 ?? "");
                cmd.Parameters.AddWithValue("@ts", DateTime.Now);
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch { txn.Rollback(); throw; }
        return count;
    }

    public static List<MeasurerFacility> GetMeasurerFacilities(string 계약번호, string? 처리시설코드 = null)
    {
        EnsureMeasurerContractTables();
        var list = new List<MeasurerFacility>();
        if (string.IsNullOrWhiteSpace(계약번호)) return list;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        if (string.IsNullOrWhiteSpace(처리시설코드))
        {
            cmd.CommandText = @"SELECT `계약번호`, `처리시설코드`, `처리시설명`, `측정시설명`, `측정시설코드`
                                FROM `측정인측정시설` WHERE `계약번호` = @no
                                ORDER BY `처리시설명` ASC, `측정시설명` ASC";
            cmd.Parameters.AddWithValue("@no", 계약번호);
        }
        else
        {
            cmd.CommandText = @"SELECT `계약번호`, `처리시설코드`, `처리시설명`, `측정시설명`, `측정시설코드`
                                FROM `측정인측정시설` WHERE `계약번호` = @no AND `처리시설코드` = @pc
                                ORDER BY `측정시설명` ASC";
            cmd.Parameters.AddWithValue("@no", 계약번호);
            cmd.Parameters.AddWithValue("@pc", 처리시설코드);
        }
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new MeasurerFacility(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4)));
        return list;
    }

    // ── 분석항목 테이블 생성 / 마이그레이션 ─────────────────────────────────
    public static void EnsureAnalysisItemTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS 측정인_분석항목 (
                Id         INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                분야       TEXT NOT NULL DEFAULT '',
                항목구분   TEXT NOT NULL DEFAULT '',
                항목명     TEXT NOT NULL,
                코드값     TEXT NOT NULL,
                select2id  TEXT,
                UNIQUE(코드값)
            );";
        cmd.ExecuteNonQuery();
        // 기존 테이블 마이그레이션
        try { cmd.CommandText = "ALTER TABLE 측정인_분석항목 ADD COLUMN 분야 TEXT NOT NULL DEFAULT ''"; cmd.ExecuteNonQuery(); } catch { }
        try { cmd.CommandText = "ALTER TABLE 측정인_분석항목 ADD COLUMN 항목구분 TEXT NOT NULL DEFAULT ''"; cmd.ExecuteNonQuery(); } catch { }
    }

    // ── 분석항목 일괄 저장 (UPSERT) ─────────────────────────────────────────
    /// <param name="items">(분야, 항목구분, 항목명, 코드값, select2id) 목록</param>
    public static int SaveAnalysisItems(List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)> items)
    {
        EnsureAnalysisItemTable();
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var (field, category, name, code, s2id) in items)
            {
                if (string.IsNullOrWhiteSpace(code)) continue;
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(
                    new[] { "코드값" },
                    new[] { "분야", "항목구분", "항목명", "select2id" });
                cmd.CommandText =
                    $"""
                    INSERT INTO 측정인_분석항목 (분야, 항목구분, 항목명, 코드값, select2id)
                    VALUES (@field, @category, @name, @code, @s2id)
                    {upsert};
                    """;
                cmd.Parameters.AddWithValue("@field",    field.Trim());
                cmd.Parameters.AddWithValue("@category", category.Trim());
                cmd.Parameters.AddWithValue("@name",     name.Trim());
                cmd.Parameters.AddWithValue("@code",     code.Trim());
                cmd.Parameters.AddWithValue("@s2id",     s2id.Trim());
                count += cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch (Exception ex)
        {
            txn.Rollback();
            throw;
        }
        return count;
    }

    // ── 분석항목 전체 조회 ───────────────────────────────────────────────────
    public static List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)> GetAllAnalysisItems()
    {
        if (_analysisItemsCache != null) return _analysisItemsCache;
        EnsureAnalysisItemTable();
        var list = new List<(string, string, string, string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 분야, 항목구분, 항목명, 코드값, select2id FROM 측정인_분석항목 ORDER BY 분야 ASC, 항목구분 ASC, 항목명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.IsDBNull(0) ? "" : r.GetString(0),
                      r.IsDBNull(1) ? "" : r.GetString(1),
                      r.GetString(2),
                      r.GetString(3),
                      r.IsDBNull(4) ? "" : r.GetString(4)));
        _analysisItemsCache = list;
        return list;
    }

    // ── 전체 데이터 수동 초기화 (필요 시 직접 호출 — 스크래핑 자동 흐름에서는 사용하지 않음) ──
    // SaveSamplingPoints가 UPSERT를 사용하므로 일반적으로 호출 불필요.
    public static void ClearAll()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM 측정인_채취지점";
        cmd.ExecuteNonQuery();
        InvalidateCache();
    }

    // ── 채취지점 데이터 저장 ─────────────────────────────────────────────────
    public static void SaveSamplingPoints(
        string companyName,
        List<string> points,
        string 계약번호 = "",
        string 계약기간 = "")
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var point in points.Where(p => !string.IsNullOrWhiteSpace(p)))
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                var upsert = DbConnectionFactory.UpsertSuffix(
                    new[] { "업체명", "채취지점명" },
                    new[] { "계약번호", "계약기간" });
                cmd.CommandText =
                    $"""
                    INSERT INTO 측정인_채취지점 (계약번호, 계약기간, 업체명, 채취지점명)
                    VALUES (@contractNo, @period, @company, @point)
                    {upsert};
                    """;
                cmd.Parameters.AddWithValue("@contractNo", 계약번호.Trim());
                cmd.Parameters.AddWithValue("@period",     계약기간.Trim());
                cmd.Parameters.AddWithValue("@company",    companyName.Trim());
                cmd.Parameters.AddWithValue("@point",      point.Trim());
                cmd.ExecuteNonQuery();
            }
            txn.Commit();
            InvalidateCache();
        }
        catch (Exception ex)
        {
            txn.Rollback();
            throw;
        }
    }

    // ── 업체별 관리 정보 조회 ────────────────────────────────────────────────
    public static (string 약칭, string 계약금액, string 견적구분) GetCompanyInfo(string companyName)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT IFNULL(약칭,''), IFNULL(계약금액,''), IFNULL(견적구분,'')
            FROM 측정인_채취지점 WHERE 업체명 = @company LIMIT 1";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        return r.Read()
            ? (r.GetString(0), r.GetString(1), r.GetString(2))
            : ("", "", "");
    }

    // ── 업체별 관리 정보 저장 ────────────────────────────────────────────────
    public static void UpdateCompanyInfo(string companyName, string 약칭, string 계약금액, string 견적구분)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE 측정인_채취지점
            SET 약칭 = @alias, 계약금액 = @amount, 견적구분 = @quotType
            WHERE 업체명 = @company";
        cmd.Parameters.AddWithValue("@alias",    약칭.Trim());
        cmd.Parameters.AddWithValue("@amount",   계약금액.Trim());
        cmd.Parameters.AddWithValue("@quotType", 견적구분.Trim());
        cmd.Parameters.AddWithValue("@company",  companyName.Trim());
        cmd.ExecuteNonQuery();
    }

    // ── 계약 DB에서 약칭을 측정인_채취지점으로 동기화 ──────────────────────────
    public static void SyncAbbrFromContractDb()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        // 약칭이 비어 있고 계약 DB에 같은 업체명이 있으면 약칭 복사
        cmd.CommandText = @"
            UPDATE 측정인_채취지점
            SET 약칭 = (
                SELECT C_Abbreviation
                FROM `계약 DB`
                WHERE `계약 DB`.C_CompanyName = 측정인_채취지점.업체명
                  AND C_Abbreviation IS NOT NULL
                  AND C_Abbreviation != ''
                LIMIT 1
            )
            WHERE (약칭 IS NULL OR 약칭 = '')
              AND EXISTS (
                SELECT 1 FROM `계약 DB`
                WHERE `계약 DB`.C_CompanyName = 측정인_채취지점.업체명
                  AND C_Abbreviation IS NOT NULL
                  AND C_Abbreviation != ''
              )";
        cmd.ExecuteNonQuery();
    }

    // ── 저장된 업체명 목록 조회 ──────────────────────────────────────────────
    public static List<string> GetCompanies()
    {
        if (_companiesCache != null) return _companiesCache;
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 업체명 FROM 측정인_채취지점 ORDER BY 업체명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        _companiesCache = list;
        return list;
    }

    // ── 특정 업체의 채취지점명 목록 조회 ─────────────────────────────────────
    public static List<string> GetSamplingPoints(string companyName)
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 채취지점명 FROM 측정인_채취지점 WHERE 업체명 = @company ORDER BY 채취지점명 ASC";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 업체 목록을 Contract 형태로 반환 (견적/의뢰서 Content4용) ─────────────
    public static List<ETA.Models.Contract> GetCompaniesAsContracts()
    {
        var map = new Dictionary<string, ETA.Models.Contract>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 업체명, IFNULL(약칭,''), IFNULL(계약기간,'')
            FROM 측정인_채취지점
            GROUP BY 업체명
            ORDER BY 업체명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            string name   = r.GetString(0);
            string abbr   = r.GetString(1);
            string period = r.GetString(2);

            DateTime? endDate = ParsePeriodEnd(period);
            map[name] = new ETA.Models.Contract
            {
                C_CompanyName  = name,
                C_Abbreviation = abbr,
                C_ContractEnd  = endDate,
            };
        }
        return [.. map.Values];
    }

    // "YYYY-MM-DD ~ YYYY-MM-DD" 또는 "YYYY.MM.DD~YYYY.MM.DD" 등 다양한 형식에서 종료일 파싱
    private static DateTime? ParsePeriodEnd(string period)
    {
        if (string.IsNullOrWhiteSpace(period)) return null;
        // 구분자 기준으로 뒷부분 추출
        var sep = new[] { " ~ ", "~", " - ", "–", "—" };
        string? endStr = null;
        foreach (var s in sep)
        {
            int idx = period.IndexOf(s, StringComparison.Ordinal);
            if (idx >= 0) { endStr = period[(idx + s.Length)..].Trim(); break; }
        }
        endStr ??= period.Trim();
        // 날짜 형식 파싱 시도
        var fmts = new[] { "yyyy-MM-dd", "yyyy.MM.dd", "yyyy/MM/dd", "yyyyMMdd" };
        foreach (var fmt in fmts)
            if (DateTime.TryParseExact(endStr, fmt, null,
                    System.Globalization.DateTimeStyles.None, out var dt))
                return dt;
        return null;
    }

    // ── 전체 데이터 조회 ──────────────────────────────────────────────────────
    public static List<(string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명)> GetAllData()
    {
        if (_allDataCache != null) return _allDataCache;
        var list = new List<(string, string, string, string, string)>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 계약번호, IFNULL(약칭,''), IFNULL(계약기간,''), 업체명, 채취지점명
            FROM 측정인_채취지점
            ORDER BY 업체명, 채취지점명";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4)));
        _allDataCache = list;
        return list;
    }

    /// <summary>
    /// 채취지점명을 기준으로 계약번호를 직접 조회합니다.
    /// sampleName이 DB의 채취지점명과 일치하는 행의 계약번호를 반환합니다.
    /// </summary>
    public static (string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명) FindContractBySamplingPoint(
        string sampleName,
        string abbr = "")
    {
        var all = GetAllData();
        if (all.Count == 0 || string.IsNullOrWhiteSpace(sampleName))
            return ("", "", "", "", "");

        string normSample = NormalizeCompany(sampleName);
        string normAbbr   = NormalizeCompany(abbr);

        // 약칭이 있으면 먼저 약칭으로 후보를 좁힘 → 다른 업체의 시료명 오매칭 방지
        var candidates = all.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(normAbbr))
        {
            var abbrMatched = all.Where(item =>
            {
                var ia = NormalizeCompany(item.약칭);
                if (string.IsNullOrWhiteSpace(ia)) return false;
                return ia == normAbbr
                    || ia.Contains(normAbbr, StringComparison.OrdinalIgnoreCase)
                    || normAbbr.Contains(ia, StringComparison.OrdinalIgnoreCase);
            }).ToList();
            if (abbrMatched.Count > 0) candidates = abbrMatched;
        }

        var best = candidates
            .Select(item =>
            {
                string itemSite = NormalizeCompany(item.채취지점명);
                string itemAbbr = NormalizeCompany(item.약칭);

                int score = 0;
                if (itemSite == normSample) score += 1000;
                else if (itemSite.Contains(normSample, StringComparison.OrdinalIgnoreCase)
                      || normSample.Contains(itemSite, StringComparison.OrdinalIgnoreCase))
                    score += 500;

                if (!string.IsNullOrWhiteSpace(normAbbr))
                {
                    if (itemAbbr == normAbbr) score += 200;
                    else if (itemAbbr.Contains(normAbbr, StringComparison.OrdinalIgnoreCase)
                          || normAbbr.Contains(itemAbbr, StringComparison.OrdinalIgnoreCase))
                        score += 80;
                }
                if (!string.IsNullOrWhiteSpace(item.계약번호)) score += 20;

                return (item, score);
            })
            .Where(x => x.score > 0)
            .OrderByDescending(x => x.score)
            .FirstOrDefault();

        return best.score > 0 ? best.item : ("", "", "", "", "");
    }

    public static (string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명) FindBestContract(
        string companyName,
        string workSite,
        string abbr = "")
    {
        var all = GetAllData();
        if (all.Count == 0)
            return ("", "", "", "", "");

        string normCompany = NormalizeCompany(companyName);
        string normSite = NormalizeCompany(workSite);
        string normAbbr = NormalizeCompany(abbr);

        int Score(string source, string query, int exactScore, int containsScore)
        {
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(query)) return 0;
            if (source == query) return exactScore;
            if (source.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                query.Contains(source, StringComparison.OrdinalIgnoreCase))
                return containsScore;
            return 0;
        }

        var best = all
            .Select(item =>
            {
                string itemCompany = NormalizeCompany(item.업체명);
                string itemSite = NormalizeCompany(item.채취지점명);
                string itemAbbr = NormalizeCompany(item.약칭);

                int score = 0;
                score += Score(itemCompany, normCompany, 500, 280);
                score += Score(itemSite, normSite, 320, 180);
                score += Score(itemAbbr, normAbbr, 120, 60);

                if (!string.IsNullOrWhiteSpace(normCompany) && !string.IsNullOrWhiteSpace(normSite))
                {
                    if ((itemCompany + itemSite).Contains(normCompany + normSite, StringComparison.OrdinalIgnoreCase) ||
                        (normCompany + normSite).Contains(itemCompany + itemSite, StringComparison.OrdinalIgnoreCase))
                        score += 140;
                }

                if (!string.IsNullOrWhiteSpace(item.계약번호))
                    score += 20;

                return (item, score);
            })
            .OrderByDescending(x => x.score)
            .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.item.계약번호))
            .FirstOrDefault();

        return best.score > 0 ? best.item : ("", "", "", "", "");
    }

    private static string NormalizeCompany(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "";

        return name
            .Replace("㈜", "주")
            .Replace("(주)", "주")
            .Replace("（주）", "주")
            .Replace("주식회사", "주")
            .Replace("유한회사", "유")
            .Replace("㈔", "사")
            .Replace("(사)", "사")
            .Replace(" ", "")
            .Replace("-", "")
            .Replace("·", "")
            .Replace("(", "")
            .Replace(")", "")
            .Replace("（", "")
            .Replace("）", "")
            .ToLowerInvariant()
            .Trim();
    }

    // =========================================================================
    // 측정시설(지역) xlsx Import — 측정인.kr 에서 다운받은 원본 파일 → DB 갱신
    //   - 측정인계약 UPSERT (계약번호 기준)
    //   - 측정인_채취지점 UPSERT (업체명, 채취지점명 기준)
    // =========================================================================

    public sealed class FacilityImportResult
    {
        public int TotalRows       { get; set; }
        public int ContractsUpsert { get; set; }  // 계약 신규/갱신
        public int PointsUpsert    { get; set; }  // 채취지점 신규/갱신
        public int Skipped         { get; set; }
        public string? Error       { get; set; }
    }

    /// <summary>Data 폴더에서 가장 최근 측정시설(지역)관리_*.xlsx 경로 반환 (없으면 null).</summary>
    public static string? FindLatestFacilityXlsx(string dataDir)
    {
        try
        {
            if (!Directory.Exists(dataDir)) return null;
            return new DirectoryInfo(dataDir)
                .GetFiles("측정시설*관리_*.xlsx")
                .Where(f => !f.Name.StartsWith("~$"))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .FirstOrDefault()?.FullName;
        }
        catch { return null; }
    }

    public static FacilityImportResult ImportFacilitiesFromXlsx(string xlsxPath)
    {
        var r = new FacilityImportResult();
        EnsureMeasurerContractTables();
        EnsureMeasurerTable();

        if (!File.Exists(xlsxPath))
        {
            r.Error = "파일을 찾을 수 없습니다: " + xlsxPath;
            return r;
        }

        // 1) xlsx 파싱
        // 컬럼(1-based): 1.측정대행업체 2.코드 3.의뢰기관 4.의뢰기관 관리번호 5.의뢰기관 주소
        //               6.의뢰기관 사업자번호 7.측정대상 사업장 8.사업장 관리번호 9.사업장 주소
        //               10.사업장 사업자번호 11.측정시설 12.측정시설 주소 13.측정분야 ...
        //               25.등록일시 26.수정일시
        var rows = new List<(string 계약번호, string 업체명, string 대표사업장, string 채취지점명, string 수정일시)>();
        try
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(xlsxPath);
            var ws = wb.Worksheets.First();
            // 1행=제목, 2행=헤더, 3행~=데이터
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int i = 3; i <= lastRow; i++)
            {
                string 계약번호 = (ws.Cell(i, 4).GetString() ?? "").Trim();
                string 업체명   = (ws.Cell(i, 7).GetString() ?? "").Trim();
                string 주소     = (ws.Cell(i, 9).GetString() ?? "").Trim();
                string 시설     = (ws.Cell(i, 11).GetString() ?? "").Trim();
                string 수정     = (ws.Cell(i, 26).GetString() ?? "").Trim();
                if (string.IsNullOrWhiteSpace(계약번호) || string.IsNullOrWhiteSpace(업체명))
                { r.Skipped++; continue; }
                rows.Add((계약번호, 업체명, 주소, 시설, 수정));
            }
            r.TotalRows = rows.Count;
        }
        catch (Exception ex)
        {
            r.Error = "xlsx 파싱 실패: " + ex.Message;
            return r;
        }

        // 2) DB UPSERT
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var txn = conn.BeginTransaction();
        try
        {
            var upsertC = DbConnectionFactory.UpsertSuffix(
                new[] { "계약번호" },
                new[] { "업체명", "대표사업장", "수집일시" });
            var upsertP = DbConnectionFactory.UpsertSuffix(
                new[] { "업체명", "채취지점명" },
                new[] { "계약번호" });
            var upsertF = DbConnectionFactory.UpsertSuffix(
                new[] { "계약번호", "처리시설명" },
                new[] { "수집일시" });

            // 계약: 같은 계약번호는 한 번만
            foreach (var g in rows.GroupBy(x => x.계약번호))
            {
                var first = g.First();
                using var cmd = conn.CreateCommand();
                cmd.Transaction = txn;
                cmd.CommandText = $@"INSERT INTO `측정인계약`
                    (`계약번호`, `업체명`, `대표사업장`, `계약기간`, `전체텍스트`, `수집일시`)
                    VALUES (@no, @co, @pl, '', @co, @ts) {upsertC}";
                cmd.Parameters.AddWithValue("@no", first.계약번호);
                cmd.Parameters.AddWithValue("@co", first.업체명);
                cmd.Parameters.AddWithValue("@pl", first.대표사업장);
                cmd.Parameters.AddWithValue("@ts", string.IsNullOrWhiteSpace(first.수정일시)
                    ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    : first.수정일시);
                cmd.ExecuteNonQuery();
                r.ContractsUpsert++;
            }

            // 채취지점 + 처리시설: 측정시설 컬럼을 두 테이블에 모두 UPSERT
            var now = DateTime.Now;
            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.채취지점명)) continue;

                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = txn;
                    cmd.CommandText = $@"INSERT INTO `측정인_채취지점`
                        (`계약번호`, `업체명`, `채취지점명`)
                        VALUES (@no, @co, @pt) {upsertP}";
                    cmd.Parameters.AddWithValue("@no", row.계약번호);
                    cmd.Parameters.AddWithValue("@co", row.업체명);
                    cmd.Parameters.AddWithValue("@pt", row.채취지점명);
                    cmd.ExecuteNonQuery();
                }
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = txn;
                    cmd.CommandText = $@"INSERT INTO `측정인처리시설`
                        (`계약번호`, `처리시설명`, `처리시설코드`, `수집일시`)
                        VALUES (@no, @nm, '', @ts) {upsertF}";
                    cmd.Parameters.AddWithValue("@no", row.계약번호);
                    cmd.Parameters.AddWithValue("@nm", row.채취지점명);
                    cmd.Parameters.AddWithValue("@ts", now);
                    cmd.ExecuteNonQuery();
                }
                r.PointsUpsert++;
            }

            txn.Commit();
        }
        catch (Exception ex)
        {
            try { txn.Rollback(); } catch { }
            r.Error = "DB 저장 실패: " + ex.Message;
            return r;
        }

        InvalidateCache();
        try
        {
            Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ImportFacilitiesFromXlsx] file='{Path.GetFileName(xlsxPath)}' " +
                $"rows={r.TotalRows} contracts={r.ContractsUpsert} points={r.PointsUpsert} skipped={r.Skipped}{Environment.NewLine}");
        }
        catch { }

        return r;
    }
}
