using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 수질분석센터 원자료 테이블 일괄 생성
/// ────────────────────────────────────────────────────────
/// - 분석정보.Analyte DISTINCT 전체 대상
/// - Category / Method / Analyte 기반 자동 스키마 분류 (7종)
///     BOD형    : 시료량, D1, D2, 희석배수, 식종*  (BOD_DATA 와 동일)
///     SS형     : 시료량, 전무게, 후무게, 무게차, 희석배수 (중량법)
///     NHexan형 : 시료량 + 결과 (단순)
///     UvVis형  : 시료량, 흡광도, 희석배수, 검량선_a, 농도  (T-P 스키마)
///     GcLc형   : UvVis + ISTD
///     Metal형  : UvVis + ST1~ST6  (금속류 — Category 매칭)
///     VOC형    : UvVis + ISTD + ST1~ST7  (휘발성유기화합물/유기물질 — Category 매칭)
/// - 테이블명: 수질분석센터_{SafeName}_DATA
///   · TOC 는 NPOC/TCIC 2개로 분리 생성
///   · 공백/특수문자는 언더스코어로 치환
/// </summary>
public static class WaterCenterDbMigration
{
    private static readonly string LogPath = Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Logs", "FacilityMigration.log");

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [WaterCenterDb] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    private static void Exec(DbConnection conn, string sql)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    // ── Public API ────────────────────────────────────────────────
    public static void EnsureWaterCenterDataTables()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureWaterCenterDataTables(conn);
    }

    public static void EnsureWaterCenterDataTables(DbConnection conn)
    {
        Log("══ EnsureWaterCenterDataTables 진입 ══");
        if (!DbConnectionFactory.TableExists(conn, "분석정보"))
        {
            Log("⚠ 분석정보 테이블 없음 — 스킵");
            return;
        }
        Log("분석정보 테이블 확인됨");

        // 0-1. GcLc 스키마 v2 마이그레이션 (1회성)
        //      기존 "피크면적, 내부표준" 구 스키마 → "흡광도, 검량선_a, ISTD" 신 스키마
        //      구 스키마 테이블을 DROP 하면 이후 루프에서 신 스키마로 재생성됨
        DropLegacyGcLcTables(conn);

        // 0-2. 기존 하이픈 버전 테이블이 있다면 하이픈 없는 이름으로 RENAME
        //     (일관성: 처리시설_TN_DATA / 처리시설_TP_DATA 네이밍 규칙 적용)
        RenameHyphenTables(conn);

        // 0-3. Metal/VOC 스키마 변경 마이그레이션 (1회성)
        //      금속류는 ST6, VOC/유기물질은 ST7 컬럼이 없으면 DROP → 신 스키마로 재생성
        DropMismatchedSchemaTables(conn);

        // 1. Analyte + Category + Method + instrument 수집
        var items = new List<(string analyte, string category, string method, string instrument)>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                SELECT DISTINCT Analyte,
                       COALESCE(Category, ''),
                       COALESCE(Method, ''),
                       COALESCE(instrument, '')
                FROM `분석정보`
                WHERE Analyte IS NOT NULL AND Analyte <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var a = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                var c = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                var m = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                var i = r.IsDBNull(3) ? "" : r.GetString(3).Trim();
                if (!string.IsNullOrWhiteSpace(a))
                    items.Add((a, c, m, i));
            }
        }
        Log($"분석정보에서 Analyte {items.Count}개 수집");

        // 2. 각 Analyte → 테이블 생성
        int created = 0;
        foreach (var (analyte, category, method, instrument) in items)
        {
            // TOC 특수: NPOC / TCIC 2개 테이블 (UvVis형, T-P 스키마)
            if (analyte.Equals("TOC", StringComparison.OrdinalIgnoreCase))
            {
                created += EnsureTable(conn, "수질분석센터_TOC_NPOC_DATA", "UvVis");
                created += EnsureTable(conn, "수질분석센터_TOC_TCIC_DATA", "TocTcic");
                continue;
            }

            var schema    = DetermineSchema(analyte, category, method, instrument);
            var tableName = $"수질분석센터_{SafeName(analyte)}_DATA";
            created += EnsureTable(conn, tableName, schema);
        }
        Log($"수질분석센터 DATA 테이블 확인 완료 — {created}개 신규 생성");

        // 4. TOC TCIC 테이블 스키마 마이그레이션 (3개 테이블 공통)
        //    - 시료량 컬럼 DROP, TCAU/TCcon/ICAU/ICcon 4개 컬럼 ADD
        MigrateTocTcicTables(conn);

        // 5. VOC 테이블 ISTD 컬럼 마이그레이션 (수질분석센터 VOC 테이블들)
        //    - ST1_ISTD ~ ST7_ISTD 컬럼 ADD
        MigrateVocIstdColumns(conn);
    }

    // ── TOC TCIC 3개 테이블 스키마 동기화 ─────────────────────────
    //   대상: TOC_TCIC_DATA / 수질분석센터_TOC_TCIC_DATA / 처리시설_TOC_TCIC_DATA
    //   변경: DROP 시료량 / ADD TCAU, TCcon, ICAU, ICcon
    private static void MigrateTocTcicTables(DbConnection conn)
    {
        var tables = new[]
        {
            "TOC_TCIC_DATA",
            "수질분석센터_TOC_TCIC_DATA",
            "처리시설_TOC_TCIC_DATA",
        };
        var newCols = new[] { "TCAU", "TCcon", "ICAU", "ICcon" };

        foreach (var tbl in tables)
        {
            if (!DbConnectionFactory.TableExists(conn, tbl))
            {
                Log($"  {tbl} 없음 — 스킵");
                continue;
            }

            // 시료량 DROP
            if (DbConnectionFactory.ColumnExists(conn, tbl, "시료량"))
            {
                try
                {
                    Exec(conn, $"ALTER TABLE `{tbl}` DROP COLUMN `시료량`");
                    Log($"  {tbl}.시료량 DROP 완료");
                }
                catch (Exception ex) { Log($"  {tbl}.시료량 DROP 실패: {ex.Message}"); }
            }

            // 신규 컬럼 ADD
            foreach (var col in newCols)
            {
                if (DbConnectionFactory.ColumnExists(conn, tbl, col)) continue;
                try
                {
                    Exec(conn, $"ALTER TABLE `{tbl}` ADD COLUMN `{col}` TEXT DEFAULT ''");
                    Log($"  {tbl}.{col} ADD 완료");
                }
                catch (Exception ex) { Log($"  {tbl}.{col} ADD 실패: {ex.Message}"); }
            }
        }
    }

    // ── Analyte + Category + Method → 스키마 결정 ──────────────────
    internal static string DetermineSchema(string analyte, string category, string method, string instrument)
    {
        // -2. 생태독성 (물벼룩 독성시험)
        if (analyte.Contains("생태독성") || category.Contains("생물류") || method.Contains("독성"))
            return "Ecotox";

        // -1. 특수 항목 (Category 보다 우선)
        //     6가크롬: 자외선/가시선분광법 + 검량선 5점 → Cr6형
        if (analyte.Contains("6가크롬") || (analyte.Contains("6가") && analyte.Contains("크롬")))
            return "Cr6";

        // 0. Category 우선 매칭 (가장 신뢰)
        if (category.Contains("금속")) return "Metal";
        if (category.Contains("휘발성") || category.Contains("유기물질")) return "VOC";

        // 1. Analyte 이름 매칭
        if (analyte.Equals("BOD", StringComparison.OrdinalIgnoreCase)) return "BOD";
        if (analyte.Equals("SS", StringComparison.OrdinalIgnoreCase) ||
            analyte.Contains("부유물질")) return "SS";
        if (analyte.Equals("N-Hexan", StringComparison.OrdinalIgnoreCase) ||
            analyte.Contains("노말헥산") || analyte.Contains("노르말헥산")) return "NHexan";

        // 2. Method 매칭
        var m = method.ToUpperInvariant();
        if (method.Contains("자외선") || method.Contains("가시선") || method.Contains("분광법"))
            return "UvVis";
        if (method.Contains("중량법")) return "SS";
        if (method.Contains("BOD") || method.Contains("생물화학적")) return "BOD";
        if (m.Contains("GC") || method.Contains("기체크로마토그래피")) return "GcLc";
        if (m.Contains("LC") || method.Contains("액체크로마토그래피")) return "GcLc";
        if (m.Contains("ICP") || m.Contains("AAS") ||
            method.Contains("유도결합플라스마") || method.Contains("원자흡광")) return "Metal";
        if (m.Contains("IC") || method.Contains("이온크로마토그래피")) return "GcLc";
        if (method.Contains("TOC") || method.Contains("유기탄소")) return "UvVis";
        if (method.Contains("노말헥산")) return "NHexan";

        // 3. instrument 보조 매칭
        var inst = instrument.ToUpperInvariant();
        if (inst.Contains("ICP") || inst.Contains("AAS")) return "Metal";
        if (inst.Contains("GC") || inst.Contains("GCMS")) return "GcLc";
        if (inst.Contains("LC") || inst.Contains("LCMS")) return "GcLc";
        if (inst.Contains("UV") || inst.Contains("VIS") || inst.Contains("분광")) return "UvVis";

        // 4. 기본: NHexan형 (단순 결과값)
        return "NHexan";
    }

    // ── VOC 테이블 ISTD 컬럼 마이그레이션 ───────────────────────────
    //   대상: 수질분석센터_*_DATA 테이블 중 ST1_농도 컬럼이 있는 테이블들
    //   추가: ST1_ISTD ~ ST7_ISTD 컬럼 (VOC/유기화합물/휘발성 테이블)
    private static void MigrateVocIstdColumns(DbConnection conn)
    {
        var istdCols = new[] { "ST1_ISTD", "ST2_ISTD", "ST3_ISTD", "ST4_ISTD", "ST5_ISTD", "ST6_ISTD", "ST7_ISTD" };

        try
        {
            // 수질분석센터_*_DATA 테이블 중 VOC/유기화합물 관련 테이블만 찾기
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE()
                  AND TABLE_NAME LIKE '수질분석센터\_%\_DATA' ESCAPE '\\'
                  AND COLUMN_NAME = 'ST1_농도'
                  AND (TABLE_NAME LIKE '%휘발성유기화합물%' OR TABLE_NAME LIKE '%유기물질%')";

            var vocTables = new List<string>();
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                vocTables.Add(reader.GetString(0));
            }
            reader.Close();

            Log($"VOC ISTD 컬럼 마이그레이션 대상: {vocTables.Count}개 테이블");

            foreach (var tbl in vocTables)
            {
                // 이미 ST1_ISTD 컬럼이 있으면 스킵
                if (DbConnectionFactory.ColumnExists(conn, tbl, "ST1_ISTD"))
                {
                    Log($"  {tbl} — 이미 ISTD 컬럼 존재, 스킵");
                    continue;
                }

                // ST1_ISTD ~ ST7_ISTD 컬럼 추가
                foreach (var col in istdCols)
                {
                    try
                    {
                        Exec(conn, $"ALTER TABLE `{tbl}` ADD COLUMN `{col}` TEXT DEFAULT ''");
                        Log($"  {tbl}.{col} 추가 완료");
                    }
                    catch (Exception ex)
                    {
                        Log($"  {tbl}.{col} 추가 실패: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log($"VOC ISTD 마이그레이션 오류: {ex.Message}");
        }
    }

    // ── Metal/VOC/Cr6 스키마 변경에 따른 기존 테이블 DROP ─────────
    //   판단 기준: 새 스키마에는 `ST1_농도` 컬럼이 있는데,
    //              기존(구) 스키마는 `ST1` 단일 컬럼만 있음
    //              → ST1_농도 컬럼이 없으면 DROP
    //   대상:
    //     · 금속류 (Category 금속*)  → Metal형
    //     · VOC/유기물질 (Category 휘발성*/유기물질) → VOC형
    //     · 6가크롬 (Analyte명 매칭) → Cr6형
    private static void DropMismatchedSchemaTables(DbConnection conn)
    {
        var targets = new List<(string analyte, string label)>();
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT Analyte, COALESCE(Category, '')
                FROM `분석정보`
                WHERE Analyte IS NOT NULL AND Analyte <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var a = r.GetString(0).Trim();
                var c = r.GetString(1).Trim();
                if (string.IsNullOrEmpty(a)) continue;

                // 6가크롬 우선
                if (a.Contains("6가크롬") || (a.Contains("6가") && a.Contains("크롬")))
                    targets.Add((a, "Cr6형"));
                else if (c.Contains("금속"))
                    targets.Add((a, "Metal형"));
                else if (c.Contains("휘발성") || c.Contains("유기물질"))
                    targets.Add((a, "VOC형"));
            }
        }
        catch (Exception ex)
        {
            Log($"DropMismatchedSchemaTables 조회 실패: {ex.Message}");
            return;
        }

        int dropped = 0;
        foreach (var (analyte, label) in targets)
        {
            var tbl = $"수질분석센터_{SafeName(analyte)}_DATA";
            if (!DbConnectionFactory.TableExists(conn, tbl)) continue;

            // 새 스키마 마커 컬럼: ST1_농도 (모든 ST 페어 형식 공통)
            if (DbConnectionFactory.ColumnExists(conn, tbl, "ST1_농도")) continue;

            try
            {
                Exec(conn, $"DROP TABLE `{tbl}`");
                Log($"DROP {tbl} ({label} 스키마 변경 — ST 농도/값 페어 도입)");
                dropped++;
            }
            catch (Exception ex) { Log($"DROP 실패 ({tbl}): {ex.Message}"); }
        }
        Log($"스키마 불일치 테이블 DROP 완료: {dropped}개 / 대상 {targets.Count}개");
    }

    // ── GcLc 구 스키마 테이블 DROP (피크면적 컬럼 가진 것들) ────────
    private static void DropLegacyGcLcTables(DbConnection conn)
    {
        var toDrop = new List<string>();
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT TABLE_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE()
                  AND TABLE_NAME LIKE '수질분석센터\_%\_DATA' ESCAPE '\\'
                  AND COLUMN_NAME = '피크면적'";
            using var r = cmd.ExecuteReader();
            while (r.Read()) toDrop.Add(r.GetString(0));
        }
        catch (Exception ex)
        {
            Log($"DropLegacyGcLcTables 탐색 실패: {ex.Message}");
            return;
        }

        if (toDrop.Count == 0)
        {
            Log("구 GcLc 스키마 테이블 없음 — 스킵");
            return;
        }

        Log($"구 GcLc 스키마 테이블 {toDrop.Count}개 발견 — DROP 진행");
        foreach (var tbl in toDrop)
        {
            try
            {
                Exec(conn, $"DROP TABLE `{tbl}`");
                Log($"DROP {tbl}");
            }
            catch (Exception ex) { Log($"DROP 실패 ({tbl}): {ex.Message}"); }
        }
    }

    // ── 하이픈 포함 테이블명을 하이픈 없는 이름으로 RENAME (일관성) ─
    private static void RenameHyphenTables(DbConnection conn)
    {
        // INFORMATION_SCHEMA 로 수질분석센터_ prefix + 하이픈 포함 테이블 탐색
        var toRename = new List<string>();
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT TABLE_NAME
                FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_SCHEMA = DATABASE()
                  AND TABLE_NAME LIKE '수질분석센터\_%\_DATA' ESCAPE '\\'
                  AND TABLE_NAME LIKE '%-%'";
            using var r = cmd.ExecuteReader();
            while (r.Read()) toRename.Add(r.GetString(0));
        }
        catch (Exception ex)
        {
            Log($"RenameHyphenTables 탐색 실패: {ex.Message}");
            return;
        }

        foreach (var oldName in toRename)
        {
            var newName = oldName.Replace("-", "");
            if (oldName == newName) continue;
            // 새 이름과 충돌하면 스킵
            if (DbConnectionFactory.TableExists(conn, newName))
            {
                Log($"RENAME 스킵 — 대상 {newName} 이미 존재");
                continue;
            }
            try
            {
                Exec(conn, $"RENAME TABLE `{oldName}` TO `{newName}`");
                Log($"RENAME {oldName} → {newName}");
            }
            catch (Exception ex) { Log($"RENAME 실패 ({oldName}): {ex.Message}"); }
        }
    }

    // ── 테이블명 안전 변환 ────────────────────────────────────────
    //   처리시설_TN_DATA / 처리시설_TP_DATA 네이밍과 통일:
    //   · 하이픈 제거 (T-N → TN, T-P → TP, N-Hexan → NHexan)
    //   · 공백/특수문자 치환
    internal static string SafeName(string analyte)
    {
        var result = analyte
            .Replace("-", "")      // 하이픈 제거 (T-N → TN)
            .Replace(" ", "_")
            .Replace("/", "_")
            .Replace("\\", "_")
            .Replace("(", "")
            .Replace(")", "")
            .Replace(",", "_")
            .Replace("·", "_")
            .Replace("+", "_plus_")
            .Replace("'", "")
            .Replace("\"", "");
        // 중복 언더스코어 정리
        while (result.Contains("__")) result = result.Replace("__", "_");
        return result.Trim('_');
    }

    // ── 스키마별 CREATE TABLE 실행 ────────────────────────────────
    private static int EnsureTable(DbConnection conn, string tableName, string schema)
    {
        if (DbConnectionFactory.TableExists(conn, tableName)) return 0;
        var ai = DbConnectionFactory.AutoIncrement;

        // 스키마별 고유 컬럼 블록
        string specificCols = schema switch
        {
            "BOD" => @"
                    시료량      TEXT DEFAULT '',
                    D1          TEXT DEFAULT '',
                    D2          TEXT DEFAULT '',
                    희석배수    TEXT DEFAULT '',
                    식종시료량  TEXT DEFAULT '',
                    식종D1      TEXT DEFAULT '',
                    식종D2      TEXT DEFAULT '',
                    식종BOD     TEXT DEFAULT '',
                    식종함유량  TEXT DEFAULT '',",

            "SS" => @"
                    시료량   TEXT DEFAULT '',
                    전무게   TEXT DEFAULT '',
                    후무게   TEXT DEFAULT '',
                    무게차   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',",

            "UvVis" => @"
                    시료량   TEXT DEFAULT '',
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    검량선_a TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',",

            // TocTcic형 = UvVis 베이스에서 시료량 제외 + TC/IC 검량선 4개 컬럼
            "TocTcic" => @"
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    검량선_a TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',
                    TCAU     TEXT DEFAULT '',
                    TCcon    TEXT DEFAULT '',
                    ICAU     TEXT DEFAULT '',
                    ICcon    TEXT DEFAULT '',",

            // GcLc형 = UvVis형 + ISTD (GC/LC는 내부표준 물질 사용)
            "GcLc" => @"
                    시료량   TEXT DEFAULT '',
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    검량선_a TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',
                    ISTD     TEXT DEFAULT '',",

            // Metal형 = UvVis 베이스 + 검량선 6점(농도/값 페어) + 회귀 결과
            "Metal" => @"
                    시료량   TEXT DEFAULT '',
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',
                    ST1_농도 TEXT DEFAULT '', ST1_값 TEXT DEFAULT '',
                    ST2_농도 TEXT DEFAULT '', ST2_값 TEXT DEFAULT '',
                    ST3_농도 TEXT DEFAULT '', ST3_값 TEXT DEFAULT '',
                    ST4_농도 TEXT DEFAULT '', ST4_값 TEXT DEFAULT '',
                    ST5_농도 TEXT DEFAULT '', ST5_값 TEXT DEFAULT '',
                    ST6_농도 TEXT DEFAULT '', ST6_값 TEXT DEFAULT '',
                    기울기   TEXT DEFAULT '',
                    절편     TEXT DEFAULT '',
                    R값      TEXT DEFAULT '',",

            // Cr6형 = UvVis 베이스 + 검량선 5점(농도/값 페어) + 회귀 결과 — 6가크롬 전용
            "Cr6" => @"
                    시료량   TEXT DEFAULT '',
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',
                    ST1_농도 TEXT DEFAULT '', ST1_값 TEXT DEFAULT '',
                    ST2_농도 TEXT DEFAULT '', ST2_값 TEXT DEFAULT '',
                    ST3_농도 TEXT DEFAULT '', ST3_값 TEXT DEFAULT '',
                    ST4_농도 TEXT DEFAULT '', ST4_값 TEXT DEFAULT '',
                    ST5_농도 TEXT DEFAULT '', ST5_값 TEXT DEFAULT '',
                    기울기   TEXT DEFAULT '',
                    절편     TEXT DEFAULT '',
                    R값      TEXT DEFAULT '',",

            // VOC형 = UvVis 베이스 + ISTD + 검량선 7점(농도/값 페어) + 회귀 결과 — 휘발성/유기물질
            "VOC" => @"
                    시료량   TEXT DEFAULT '',
                    흡광도   TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    농도     TEXT DEFAULT '',
                    ISTD     TEXT DEFAULT '',
                    ST1_농도 TEXT DEFAULT '', ST1_값 TEXT DEFAULT '', ST1_ISTD TEXT DEFAULT '',
                    ST2_농도 TEXT DEFAULT '', ST2_값 TEXT DEFAULT '', ST2_ISTD TEXT DEFAULT '',
                    ST3_농도 TEXT DEFAULT '', ST3_값 TEXT DEFAULT '', ST3_ISTD TEXT DEFAULT '',
                    ST4_농도 TEXT DEFAULT '', ST4_값 TEXT DEFAULT '', ST4_ISTD TEXT DEFAULT '',
                    ST5_농도 TEXT DEFAULT '', ST5_값 TEXT DEFAULT '', ST5_ISTD TEXT DEFAULT '',
                    ST6_농도 TEXT DEFAULT '', ST6_값 TEXT DEFAULT '', ST6_ISTD TEXT DEFAULT '',
                    ST7_농도 TEXT DEFAULT '', ST7_값 TEXT DEFAULT '', ST7_ISTD TEXT DEFAULT '',
                    기울기   TEXT DEFAULT '',
                    절편     TEXT DEFAULT '',
                    R값      TEXT DEFAULT '',",

            // NHexan (단순)
            _ => @"
                    시료량 TEXT DEFAULT '',"
        };

        try
        {
            // MariaDB: UNIQUE 키 컬럼은 TEXT 불가 → VARCHAR 사용
            Exec(conn, $@"
                CREATE TABLE `{tableName}` (
                    id       INTEGER PRIMARY KEY {ai},
                    분석일   VARCHAR(20)  NOT NULL,
                    SN       VARCHAR(100) NOT NULL,
                    업체명   VARCHAR(191) DEFAULT '',
                    구분     VARCHAR(50)  DEFAULT '',
                    {specificCols}
                    결과     TEXT DEFAULT '',
                    비고     TEXT DEFAULT '',
                    등록일시 VARCHAR(20) DEFAULT '',
                    UNIQUE(분석일, SN)
                )");
            Log($"{tableName} 생성 ({schema}형)");
            return 1;
        }
        catch (Exception ex)
        {
            Log($"{tableName} 생성 실패: {ex.Message}");
            return 0;
        }
    }
}
