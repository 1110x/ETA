using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>처리시설_마스터 / 폐수배출업소_분석결과 테이블 생성 및 초기 데이터 투입</summary>
public static class FacilityDbMigration
{
    // ══════════════════════════════════════════════════════════════════════
    public static void EnsureTables()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureProcessingFacilityMaster(conn);
        EnsureWasteCompanyResults(conn);
        EnsureFacilityResults(conn);
        EnsureWasteRequests(conn);
        EnsureWasteResultTables(conn);
        EnsureResultSubmitLog(conn);
        EnsureYeosuCompanyTable(conn);
        EnsureWasteRequestResultTable(conn);
        EnsureAnalysisDataTables(conn);
        EnsureMigrationTable(conn);
        if (!IsMigrationDone(conn, "wasteresult_v2"))
        {
            MigrateDataToWasteRequestResult(conn);
            MarkMigrationDone(conn, "wasteresult_v2");
        }
        else
        {
            Log("마이그레이션 이미 완료됨 (wasteresult_v2) — 건너뜀");
        }
    }

    // ── 분석 원시 데이터 테이블 (*_DATA) ────────────────────────────────────
    private static void EnsureAnalysisDataTables(DbConnection conn)
    {
        var ai = DbConnectionFactory.AutoIncrement;

        // BOD_DATA
        if (!DbConnectionFactory.TableExists(conn, "BOD_DATA"))
        {
            Exec(conn, $@"
                CREATE TABLE `BOD_DATA` (
                    id          INTEGER PRIMARY KEY {ai},
                    분석일      TEXT NOT NULL,
                    SN          TEXT NOT NULL,
                    업체명      TEXT DEFAULT '',
                    구분        TEXT DEFAULT '',
                    시료량      TEXT DEFAULT '',
                    D1          TEXT DEFAULT '',
                    D2          TEXT DEFAULT '',
                    희석배수    TEXT DEFAULT '',
                    결과        TEXT DEFAULT '',
                    식종시료량  TEXT DEFAULT '',
                    식종D1      TEXT DEFAULT '',
                    식종D2      TEXT DEFAULT '',
                    식종BOD     TEXT DEFAULT '',
                    식종함유량  TEXT DEFAULT '',
                    등록일시    TEXT DEFAULT '',
                    UNIQUE(분석일, SN)
                )");
            Log("BOD_DATA 테이블 생성");
        }

        // SS_DATA — 기존 테이블 컬럼명 불일치 수정 (한자 析U+6790·日U+65E5 → 한글 석U+C11D·일U+C77C)
        if (DbConnectionFactory.ColumnExists(conn, "SS_DATA", "\uBD84\u6790\u65E5"))
        {
            try { Exec(conn, "ALTER TABLE `SS_DATA` CHANGE `\uBD84\u6790\u65E5` `\uBD84\uC11D\uC77C` TEXT NOT NULL"); Log("SS_DATA 컬럼명 한자→한글 마이그레이션 완료"); }
            catch (Exception ex) { Log($"SS_DATA 컬럼명 마이그레이션 실패: {ex.Message}"); }
        }

        // SS_DATA
        if (!DbConnectionFactory.TableExists(conn, "SS_DATA"))
        {
            Exec(conn, $@"
                CREATE TABLE `SS_DATA` (
                    id      INTEGER PRIMARY KEY {ai},
                    분석일  TEXT NOT NULL,
                    SN      TEXT NOT NULL,
                    업체명  TEXT DEFAULT '',
                    구분    TEXT DEFAULT '',
                    시료량  TEXT DEFAULT '',
                    전무게  TEXT DEFAULT '',
                    후무게  TEXT DEFAULT '',
                    무게차  TEXT DEFAULT '',
                    희석배수 TEXT DEFAULT '',
                    결과    TEXT DEFAULT '',
                    등록일시 TEXT DEFAULT '',
                    UNIQUE(분석일, SN)
                )");
            Log("SS_DATA 테이블 생성");
        }

        // NHexan_DATA (단순 결과값 형식)
        if (!DbConnectionFactory.TableExists(conn, "NHexan_DATA"))
        {
            Exec(conn, $@"
                CREATE TABLE `NHexan_DATA` (
                    id      INTEGER PRIMARY KEY {ai},
                    분석일  TEXT NOT NULL,
                    SN      TEXT NOT NULL,
                    업체명  TEXT DEFAULT '',
                    구분    TEXT DEFAULT '',
                    시료량  TEXT DEFAULT '',
                    결과    TEXT DEFAULT '',
                    등록일시 TEXT DEFAULT '',
                    UNIQUE(분석일, SN)
                )");
            Log("NHexan_DATA 테이블 생성");
        }

        // TN_DATA, TP_DATA, Phenols_DATA — UvVis 형식 (흡광도/검량선_a/농도)
        foreach (var tbl in new[] { "TN_DATA", "TP_DATA", "Phenols_DATA" })
        {
            if (!DbConnectionFactory.TableExists(conn, tbl))
            {
                Exec(conn, $@"
                    CREATE TABLE `{tbl}` (
                        id      INTEGER PRIMARY KEY {ai},
                        분석일  TEXT NOT NULL,
                        SN      TEXT NOT NULL,
                        업체명  TEXT DEFAULT '',
                        구분    TEXT DEFAULT '',
                        시료량  TEXT DEFAULT '',
                        흡광도  TEXT DEFAULT '',
                        희석배수 TEXT DEFAULT '',
                        검량선_a TEXT DEFAULT '',
                        농도    TEXT DEFAULT '',
                        등록일시 TEXT DEFAULT '',
                        UNIQUE(분석일, SN)
                    )");
                Log($"{tbl} 테이블 생성");
            }
        }

        // TOC_TCIC_DATA, TOC_NPOC_DATA — UvVis 형식
        foreach (var tbl in new[] { "TOC_TCIC_DATA", "TOC_NPOC_DATA" })
        {
            if (!DbConnectionFactory.TableExists(conn, tbl))
            {
                Exec(conn, $@"
                    CREATE TABLE `{tbl}` (
                        id      INTEGER PRIMARY KEY {ai},
                        분석일  TEXT NOT NULL,
                        SN      TEXT NOT NULL,
                        업체명  TEXT DEFAULT '',
                        구분    TEXT DEFAULT '',
                        시료량  TEXT DEFAULT '',
                        흡광도  TEXT DEFAULT '',
                        희석배수 TEXT DEFAULT '',
                        검량선_a TEXT DEFAULT '',
                        농도    TEXT DEFAULT '',
                        등록일시 TEXT DEFAULT '',
                        UNIQUE(분석일, SN)
                    )");
                Log($"{tbl} 테이블 생성");
            }
        }
    }

    // ── 처리시설_마스터 ────────────────────────────────────────────────────
    private static void EnsureProcessingFacilityMaster(DbConnection conn)
    {
        if (!DbConnectionFactory.TableExists(conn, "처리시설_마스터"))
        {
            Exec(conn, $@"
                CREATE TABLE `처리시설_마스터` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    시설명      TEXT NOT NULL,
                    시료명      TEXT NOT NULL,
                    BOD         TEXT DEFAULT '',
                    TOC         TEXT DEFAULT '',
                    SS          TEXT DEFAULT '',
                    `T-N`       TEXT DEFAULT '',
                    `T-P`       TEXT DEFAULT '',
                    총대장균군  TEXT DEFAULT '',
                    COD         TEXT DEFAULT '',
                    염소이온    TEXT DEFAULT '',
                    영양염류    TEXT DEFAULT '',
                    함수율      TEXT DEFAULT '',
                    중금속      TEXT DEFAULT '',
                    비고        TEXT DEFAULT ''
                )");
            Log("처리시설_마스터 테이블 생성");
        }

        // 테이블이 비어 있으면 초기 데이터 투입 (최초 생성 시 또는 이전 시드 실패 시)
        using var countCmd = conn.CreateCommand();
        countCmd.CommandText = "SELECT COUNT(*) FROM `처리시설_마스터`";
        if (Convert.ToInt32(countCmd.ExecuteScalar()) == 0)
            SeedFacilityMaster(conn);

        // 시설명 오타 수정 (기존 DB 데이터 일괄 교정)
        Exec(conn, "UPDATE `처리시설_마스터` SET 시설명 = '세풍' WHERE 시설명 = '세룡'");
        Exec(conn, "UPDATE `처리시설_측정결과` SET 시설명 = '세풍' WHERE 시설명 = '세룡'");
        Exec(conn, "UPDATE `처리시설_마스터` SET 시설명 = '해룡개인산단' WHERE 시설명 = '폐흥개인산단'");
        Exec(conn, "UPDATE `처리시설_측정결과` SET 시설명 = '해룡개인산단' WHERE 시설명 = '폐흥개인산단'");
        Exec(conn, "UPDATE `처리시설_마스터` SET 시설명 = REPLACE(시설명, '원내', '월내') WHERE 시설명 LIKE '%원내%'");
        Exec(conn, "UPDATE `처리시설_측정결과` SET 시설명 = REPLACE(시설명, '원내', '월내') WHERE 시설명 LIKE '%원내%'");
    }

    private static void SeedFacilityMaster(DbConnection conn)
    {
        // (시설명, 시료명, BOD, TOC, SS, T-N, T-P, 총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속, 비고)
        var rows = new List<(string 시설명, string 시료명, string BOD, string TOC, string SS,
            string TN, string TP, string 총대장균군, string COD, string 염소이온,
            string 영양염류, string 함수율, string 중금속, string 비고)>
        {
            // ── 중흘처리장 ──────────────────────────────────────────────────
            ("중흘처리장", "농축기 탈리액",      "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("중흘처리장", "탈수기 탈리액",      "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("중흘처리장", "유입수",             "O","O","O","O","O","O","O","O","","","", "주1회 목요일"),
            ("중흘처리장", "유입수",             "O","","","","","","","","","","",  ""),
            ("중흘처리장", "유입수",             "O","","","","","","","","","","",  ""),
            ("중흘처리장", "유량 조정조",        "O","O","O","O","O","","","","O","","", ""),
            ("중흘처리장", "생물반응조 A",       "","","O(MLSS)","","","","","","","","", ""),
            ("중흘처리장", "생물반응조 B",       "","","O(MLSS)","","","","","","","","", ""),
            ("중흘처리장", "생물반응조 C",       "","","O(MLSS)","","","","","","","","", ""),
            ("중흘처리장", "생물반응조 D",       "","","O(MLSS)","","","","","","","","", ""),
            ("중흘처리장", "반송",               "","","O(MLSS)","","","","","","","","", ""),
            ("중흘처리장", "슬러지저류조",       "","","O(MLSS)","","","","","","","","", "주1회 목요일"),
            ("중흘처리장", "고침침어",           "","","O(MLSS)","","","","","","","","", "주1회 목요일"),
            ("중흘처리장", "약침침어",           "","","O(MLSS)","","","","","","","","", "주1회 목요일"),
            ("중흘처리장", "기존 2차 침전지",   "O","O","O","O","O","","","","","","", ""),
            ("중흘처리장", "증설 2차 침전지",   "O","O","O","O","O","","","","","","", ""),
            ("중흘처리장", "약품응집침전지-기존","O","O","O","O","O","","","","O","","", ""),
            ("중흘처리장", "약품응집침전지-증설","O","O","O","O","O","","","","O","","", ""),
            ("중흘처리장", "MDF 여과기",         "O","O","O","O","O","","","","O","","", ""),
            ("중흘처리장", "A/C 여과기",         "O","O","O","O","O","","","","","","", ""),
            ("중흘처리장", "방류수",             "O","O","O","O","O","O","O","O","","","O", ""),
            ("중흘처리장", "탈수케이크",         "","","","","","","","","","O","O", ""),

            // ── 월내처리장 ──────────────────────────────────────────────────
            ("월내처리장", "농축기 탈리액",      "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("월내처리장", "탈수기 탈리액",      "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("월내처리장", "유입수",             "O","O","O","O","O","O","O","O","O","","O", "주1회 목요일"),
            ("월내처리장", "유입수",             "O","","","","","","","","","","", ""),
            ("월내처리장", "유입수",             "O","","","","","","","","","","", ""),
            ("월내처리장", "유량 조정조",        "O","O","O","O","O","","","","O","","", ""),
            ("월내처리장", "생물반응조",         "","","O(MLSS)","","","","","","","","", ""),
            ("월내처리장", "2차 침전지",         "O","O","O","O","O","","","","","","", ""),
            ("월내처리장", "가압부상조 유출수",  "O","O","O","O","O","","","","","","", ""),
            ("월내처리장", "방류수",             "O","O","O","O","O","O","O","O","O","","O", ""),
            ("월내처리장", "기존반송",           "","","O(MLSS)","","","","","","","","", ""),
            ("월내처리장", "증설반송",           "","","O(MLSS)","","","","","","","","", ""),
            ("월내처리장", "신설저류조",         "","","O(MLSS)","","","","","","","","", ""),
            ("월내처리장", "가압부저류조",       "","","O(MLSS)","","","","","","","","", ""),
            ("월내처리장", "탈수케이크",         "","","","","","","","","","O","O", ""),

            // ── 4단계 ────────────────────────────────────────────────────────
            ("4단계", "농축기 탈리액",           "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("4단계", "탈수기 탈리액",           "O","O","O","O","O","","","","","","", "주1회 목요일"),
            ("4단계", "유입수",                  "O","O","O","O","O","O","","","","","", "주2회 화,목"),
            ("4단계", "유입수",                  "O","","","","","","","","","","", ""),
            ("4단계", "유입수",                  "O","","","","","","","","","","", ""),
            ("4단계", "유량 조정조",             "O","O","O","O","O","","","","","","", ""),
            ("4단계", "생물반응조",              "","","O(MLSS)","","","","","","","","", ""),
            ("4단계", "고속응집침전지",          "O","O","O","O","O","","","","","","", "주2회 화,목"),
            ("4단계", "복합처리설비",            "O","O","O","O","O","","","","","","", "주2회 화,목"),
            ("4단계", "방류수",                  "O","O","O","O","O","O","","","","","", ""),
            ("4단계", "탈수케이크",              "","","","","","","","","","O","O", ""),

            // ── 율촌처리장 ──────────────────────────────────────────────────
            ("율촌처리장", "유입수",             "O","O","O","O","O","O","","","","","O", ""),
            ("율촌처리장", "유입수",             "O","","","","","","","","","","", ""),
            ("율촌처리장", "유입수",             "O","","","","","","","","O","","", ""),
            ("율촌처리장", "분배조",             "O","O","O","O","O","","","","","","", ""),
            ("율촌처리장", "기존 고속응집침전지-유출","O","O","O","O","O","","","","","","", ""),
            ("율촌처리장", "증설 고속응집침전지-유출","O","O","O","O","O","","","","","","", ""),
            ("율촌처리장", "방류수",             "O","O","O","O","O","O","","","O","","O", ""),
            ("율촌처리장", "탈수케이크",         "","","","","","","","","","","O", ""),

            // ── 세풍 ────────────────────────────────────────────────────────
            ("세풍", "유입수",                   "O","O","O","O","O","O","","","","","", ""),
            ("세풍", "유입수",                   "O","","","","","","","","","","", ""),
            ("세풍", "유입수",                   "O","","","","","","","","","","", ""),
            ("세풍", "분배조",                   "O","O","O","O","O","","","","","","", ""),
            ("세풍", "방류수",                   "O","O","O","O","O","O","","","","","", ""),

            // ── 폐흘개인산단 ─────────────────────────────────────────────────
            ("폐흘개인산단", "유입수",           "O","O","O","O","O","O","","","","","", "주1회"),
            ("폐흘개인산단", "유입수",           "O","","","","","","","","","","", "주1회"),
            ("폐흘개인산단", "유입수",           "O","","","","","","","","","","", "주1회"),
            ("폐흘개인산단", "방류수",           "O","O","O","O","O","","","","","","", "주1회"),
        };

        foreach (var r in rows)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO `처리시설_마스터`
                    (시설명, 시료명, BOD, TOC, SS, `T-N`, `T-P`,
                     총대장균군, COD, 염소이온, 영양염류, 함수율, 중금속, 비고)
                VALUES
                    (@시설명, @시료명, @BOD, @TOC, @SS, @TN, @TP,
                     @총대장균군, @COD, @염소이온, @영양염류, @함수율, @중금속, @비고)";
            cmd.Parameters.AddWithValue("@시설명",    r.시설명);
            cmd.Parameters.AddWithValue("@시료명",    r.시료명);
            cmd.Parameters.AddWithValue("@BOD",       r.BOD);
            cmd.Parameters.AddWithValue("@TOC",       r.TOC);
            cmd.Parameters.AddWithValue("@SS",        r.SS);
            cmd.Parameters.AddWithValue("@TN",        r.TN);
            cmd.Parameters.AddWithValue("@TP",        r.TP);
            cmd.Parameters.AddWithValue("@총대장균군", r.총대장균군);
            cmd.Parameters.AddWithValue("@COD",       r.COD);
            cmd.Parameters.AddWithValue("@염소이온",  r.염소이온);
            cmd.Parameters.AddWithValue("@영양염류",  r.영양염류);
            cmd.Parameters.AddWithValue("@함수율",    r.함수율);
            cmd.Parameters.AddWithValue("@중금속",    r.중금속);
            cmd.Parameters.AddWithValue("@비고",      r.비고);
            cmd.ExecuteNonQuery();
        }

        Log($"처리시설_마스터 시드 완료: {rows.Count}행");
    }

    // ── 폐수배출업소_분석결과 → 분석결과 ─────────────────────────────────
    private static void EnsureWasteCompanyResults(DbConnection conn)
    {
        // 구버전 한자 테이블명이 있으면 먼저 이름 변경
        if (DbConnectionFactory.TableExists(conn, "폐수배출업소_분석결과"))
            Exec(conn, "ALTER TABLE `폐수배출업소_분석결과` RENAME TO `폐수배출업소_분석결과`");

        if (DbConnectionFactory.TableExists(conn, "폐수배출업소_분석결과")) return;

        Exec(conn, $@"
            CREATE TABLE `폐수배출업소_분석결과` (
                id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                관리번호    TEXT NOT NULL,
                업체명      TEXT NOT NULL,
                채취일자    TEXT NOT NULL,
                BOD         TEXT DEFAULT '',
                TOC         TEXT DEFAULT '',
                SS          TEXT DEFAULT '',
                `T-N`       TEXT DEFAULT '',
                `T-P`       TEXT DEFAULT '',
                Phenols     TEXT DEFAULT '',
                `N-Hexane`  TEXT DEFAULT '',
                비고        TEXT DEFAULT '',
                입력자      TEXT DEFAULT '',
                입력일시    TEXT DEFAULT ''
            )");

        Log("폐수배출업소_분석결과 테이블 생성");
    }

    // ── 처리시설_측정결과 ──────────────────────────────────────────────────
    private static void EnsureFacilityResults(DbConnection conn)
    {
        if (DbConnectionFactory.TableExists(conn, "처리시설_측정결과")) return;

        Exec(conn, $@"
            CREATE TABLE `처리시설_측정결과` (
                id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                마스터_id   INTEGER NOT NULL,
                시설명      TEXT NOT NULL,
                시료명      TEXT NOT NULL,
                채취일자    TEXT NOT NULL,
                BOD         TEXT DEFAULT '',
                TOC         TEXT DEFAULT '',
                SS          TEXT DEFAULT '',
                `T-N`       TEXT DEFAULT '',
                `T-P`       TEXT DEFAULT '',
                총대장균군  TEXT DEFAULT '',
                COD         TEXT DEFAULT '',
                염소이온    TEXT DEFAULT '',
                영양염류    TEXT DEFAULT '',
                함수율      TEXT DEFAULT '',
                중금속      TEXT DEFAULT '',
                비고        TEXT DEFAULT '',
                입력자      TEXT DEFAULT '',
                입력일시    TEXT DEFAULT ''
            )");

        Log("처리시설_측정결과 테이블 생성");
    }

    // ── 폐수_의뢰 / 폐수_의뢰_항목 / 처리시설_작업 ───────────────────────
    private static void EnsureWasteRequests(DbConnection conn)
    {
        if (!DbConnectionFactory.TableExists(conn, "폐수_의뢰"))
        {
            Exec(conn, $@"
                CREATE TABLE `폐수_의뢰` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    의뢰번호    VARCHAR(191) NOT NULL UNIQUE,
                    구분        TEXT NOT NULL,
                    채취일자    TEXT NOT NULL,
                    업체명      TEXT NOT NULL,
                    관리번호    TEXT DEFAULT '',
                    상태        TEXT DEFAULT '대기',
                    등록자      TEXT DEFAULT '',
                    등록일시    TEXT DEFAULT ''
                )");
            Log("폐수_의뢰 테이블 생성");
        }

        if (!DbConnectionFactory.TableExists(conn, "폐수_의뢰_항목"))
        {
            Exec(conn, $@"
                CREATE TABLE `폐수_의뢰_항목` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    의뢰_id     INTEGER NOT NULL,
                    항목        TEXT NOT NULL,
                    상태        TEXT DEFAULT '미담',
                    배정자      TEXT DEFAULT '',
                    배정일시    TEXT DEFAULT '',
                    완료일시    TEXT DEFAULT ''
                )");
            Log("폐수_의뢰_항목 테이블 생성");
        }

        if (!DbConnectionFactory.TableExists(conn, "처리시설_작업"))
        {
            Exec(conn, $@"
                CREATE TABLE `처리시설_작업` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    마스터_id   INTEGER NOT NULL,
                    채취일자    VARCHAR(20) NOT NULL,
                    시설명      TEXT NOT NULL,
                    시료명      TEXT NOT NULL,
                    항목목록    TEXT DEFAULT '',
                    상태        TEXT DEFAULT '미담',
                    배정자      TEXT DEFAULT '',
                    배정일시    TEXT DEFAULT '',
                    완료일시    TEXT DEFAULT '',
                    UNIQUE(마스터_id, 채취일자)
                )");
            Log("처리시설_작업 테이블 생성");
        }
    }

    // ── 폐수_결과 / 폐수_결과_항목 ───────────────────────────────────────
    private static void EnsureWasteResultTables(DbConnection conn)
    {
        if (!DbConnectionFactory.TableExists(conn, "폐수_결과"))
        {
            Exec(conn, $@"
                CREATE TABLE `폐수_결과` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    의뢰_id     INTEGER DEFAULT NULL,
                    관리번호    VARCHAR(191) NOT NULL,
                    업체명      TEXT NOT NULL,
                    채취일자    VARCHAR(50) NOT NULL,
                    입력자      TEXT DEFAULT '',
                    입력일시    TEXT DEFAULT '',
                    UNIQUE(관리번호, 채취일자)
                )");
            Log("폐수_결과 테이블 생성");
        }

        if (!DbConnectionFactory.TableExists(conn, "폐수_결과_항목"))
        {
            Exec(conn, $@"
                CREATE TABLE `폐수_결과_항목` (
                    id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    결과_id     INTEGER NOT NULL,
                    항목        TEXT NOT NULL,
                    결과값      TEXT DEFAULT '',
                    단위        TEXT DEFAULT '',
                    비고        TEXT DEFAULT ''
                )");
            Log("폐수_결과_항목 테이블 생성");
        }
    }

    // ── 결과_제출이력 ──────────────────────────────────────────────────────
    private static void EnsureResultSubmitLog(DbConnection conn)
    {
        if (DbConnectionFactory.TableExists(conn, "결과_제출이력")) return;

        Exec(conn, $@"
            CREATE TABLE `결과_제출이력` (
                id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                제출유형    TEXT NOT NULL,
                대상유형    TEXT NOT NULL,
                대상명      TEXT NOT NULL,
                기간_시작   TEXT NOT NULL,
                기간_종료   TEXT NOT NULL,
                제출건수    INTEGER DEFAULT 0,
                제출자      TEXT DEFAULT '',
                제출일시    TEXT DEFAULT '',
                파일경로    TEXT DEFAULT '',
                상태        TEXT DEFAULT '대기',
                비고        TEXT DEFAULT ''
            )");
        Log("결과_제출이력 테이블 생성");
    }

    // ── 여수_폐수배출업소 ─────────────────────────────────────────────────
    private static void EnsureYeosuCompanyTable(DbConnection conn)
    {
        if (!DbConnectionFactory.TableExists(conn, "여수_폐수배출업소"))
        {
            Exec(conn, $@"
                CREATE TABLE `여수_폐수배출업소` (
                    id      INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    업체명  TEXT NOT NULL,
                    약칭    TEXT DEFAULT ''
                )");
            Log("여수_폐수배출업소 테이블 생성");
        }

        using var countCmd = conn.CreateCommand();
        countCmd.CommandText = "SELECT COUNT(*) FROM `여수_폐수배출업소`";
        if (Convert.ToInt32(countCmd.ExecuteScalar()) > 0) return;

        var names = new[]
        {
            "㈜동성코퍼레이션 여수공장",
            "㈜세아엠앤에스",
            "㈜엘지화학",
            "㈜엘지화학[CA/EDC공장]",
            "㈜엘지화학[브이시엠공장]",
            "㈜엘지화학[아크릴레이트2공장]",
            "㈜엘지화학[엔씨씨옥소알콜-엔씨씨]",
            "㈜엘지화학[엔씨씨옥소알콜-옥소알콜]",
            "㈜유니테크 여수지점",
            "㈜일렘테크놀러지",
            "㈜청경",
            "㈜케이아이엔티",
            "GS 칼텍스 중흥주유소",
            "KPX라이프사이언스㈜",
            "구다우케미칼",
            "금호미쓰이화학㈜",
            "금호석유화학㈜[1공장]",
            "금호석유화학㈜[2공장]",
            "금호석유화학㈜[AU공장]",
            "금호석유화학㈜[고무약품공장]",
            "금호석유화학㈜[제2열병합발전소]",
            "금호티앤엘㈜",
            "금호폴리켐㈜여수1공장",
            "금호폴리켐㈜여수2공장",
            "금호피앤비화학㈜",
            "남해화학㈜ 여수공장",
            "대한통운㈜",
            "더블유알그레이스코리아(유)",
            "데이원에너지㈜",
            "롯데케미칼 주식회사 1공장",
            "롯데케미칼 주식회사 1공장(확장단지)",
            "롯데케미칼 주식회사 2공장",
            "롯데케미칼㈜[첨단소재여수공장]",
            "맛시락",
            "백광산업㈜여수공장",
            "비를라카본코리아㈜",
            "사후관리",
            "삼남석유화학㈜ 여수공장",
            "삼양화학실업㈜",
            "스미토모세이카폴리머스코리아㈜",
            "신대신카서비스,세차장",
            "신승오앤에프",
            "에어리퀴드코리아㈜",
            "엘지엠엠에이㈜",
            "여수시청(폐기물소각장)",
            "여수시청[위생매립장 월내처리장]",
            "여수환경산업",
            "여천NCC㈜1공장-환경안전팀",
            "여천NCC㈜2공장-환경안전팀",
            "여천NCC㈜3공장-환경안전팀",
            "오리온엔지니어드카본즈㈜",
            "재원산업㈜",
            "제이셀㈜",
            "㈜동남환경",
            "㈜와이엔텍",
            "㈜와이엔텍(5매립장)",
            "㈜와이엔텍(6매립장)",
            "㈜케미렉스",
            "㈜한국환경사업단",
            "㈜한화 여수사업장",
            "지에스칼텍스㈜여수공장",
            "지에스칼텍스㈜여수공장[NO.2 HOU]",
            "지에스칼텍스㈜여수공장[NO.3 HOU]",
            "지에스칼텍스㈜[바이오부탄올실증센터]",
            "지에스칼텍스㈜MFC",
            "케이씨환경서비스㈜",
            "케이씨환경서비스㈜여수사업부",
            "코오롱인더스트리㈜여수공장",
            "한국바스프㈜",
            "한국실리콘㈜",
            "한화솔루션㈜여수[1공장]",
            "한화솔루션㈜여수[2공장]",
            "한화솔루션㈜여수[3공장]",
            "한화솔루션㈜티디아이",
            "한화솔루션㈜폴리실리콘",
            "한화에너지㈜여수공장",
            "한화컴파운드㈜",
            "해인기업㈜여수공장",
            "해인기업㈜ 여천",
            "해인기업㈜(화치공장)",
            "현대에너지㈜",
            "호남환경에너지",
            "환경시설관리㈜ 슬러지 자원화",
            "휴켐스㈜ 여수공장",
            "태경케미컬㈜",
            "그린생명과학주식회사",
            "금호석유화학㈜정밀화학AU공장",
            "금호석유화학㈜[정밀고무약품공장]",
            "넥스워터㈜",
            "오라이온코리아㈜",
            "보임열병합발전㈜",
            "한국동서발전㈜ 신호남건설추진본부",
            "호남화력발전(한국남서발전)",
            "씨이케이㈜",
            "엘엑스엠엠에이㈜",
            "씨이케이㈜여수사업부",
            "신대신카서비스세차장",
            "㈜동성케미컬",
            "리뉴어스㈜ 슬러지자원화",
            "㈜롯데지에스화학",
            "티케이지휴켐스㈜",
            "피케이씨",
            "㈜티케이지일렘",
            "월드이엔티㈜",
        };

        foreach (var name in names)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO `여수_폐수배출업소` (업체명, 약칭) VALUES (@n, '')";
            cmd.Parameters.AddWithValue("@n", name);
            cmd.ExecuteNonQuery();
        }
        Log($"여수_폐수배출업소 시드 완료: {names.Length}행");
    }

    // ── 폐수의뢰및결과 (테이블 생성 + 분석항목 컬럼 추가) ───────────────
    private static void EnsureWasteRequestResultTable(DbConnection conn)
    {
        // 구 테이블명 → 신 테이블명 RENAME
        if (DbConnectionFactory.TableExists(conn, "폐수채수의뢰")
            && !DbConnectionFactory.TableExists(conn, "폐수의뢰및결과"))
        {
            Exec(conn, "ALTER TABLE `폐수채수의뢰` RENAME TO `폐수의뢰및결과`");
            Log("폐수채수의뢰 → 폐수의뢰및결과 테이블명 변경");
        }

        if (!DbConnectionFactory.TableExists(conn, "폐수의뢰및결과"))
        {
            Exec(conn, $@"
                CREATE TABLE `폐수의뢰및결과` (
                    Id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    채수일      TEXT NOT NULL,
                    구분        TEXT NOT NULL DEFAULT '여수',
                    순서        INTEGER NOT NULL DEFAULT 0,
                    SN          TEXT DEFAULT '',
                    업체명      TEXT NOT NULL DEFAULT '',
                    관리번호    TEXT DEFAULT '',
                    BOD              TEXT DEFAULT '',
                    `TOC`     TEXT DEFAULT '',
                    SS               TEXT DEFAULT '',
                    `T-N`            TEXT DEFAULT '',
                    `T-P`            TEXT DEFAULT '',
                    `N-Hexan`        TEXT DEFAULT '',
                    Phenols          TEXT DEFAULT '',
                    비고        TEXT DEFAULT '',
                    확인자      TEXT DEFAULT ''
                )");
            Log("폐수의뢰및결과 테이블 생성");
            return;
        }

        // 구 컬럼명 TOC(TC-IC) → TOC 이관
        if (ColumnExists(conn, "폐수의뢰및결과", "TOC(TC-IC)") && !ColumnExists(conn, "폐수의뢰및결과", "TOC"))
        {
            Exec(conn, "ALTER TABLE `폐수의뢰및결과` ADD COLUMN `TOC` TEXT DEFAULT ''");
            Exec(conn, "UPDATE `폐수의뢰및결과` SET `TOC` = `TOC(TC-IC)` WHERE `TOC(TC-IC)` <> ''");
            Log("폐수의뢰및결과 TOC(TC-IC) → TOC 컬럼 이관");
        }

        // 기존 테이블에 컬럼이 없으면 추가
        var cols = new[] { "BOD", "TOC", "SS", "T-N", "T-P", "N-Hexan", "Phenols" };
        foreach (var col in cols)
        {
            if (!ColumnExists(conn, "폐수의뢰및결과", col))
            {
                Exec(conn, $"ALTER TABLE `폐수의뢰및결과` ADD COLUMN `{col}` TEXT DEFAULT ''");
                Log($"폐수의뢰및결과 컬럼 추가: {col}");
            }
        }
    }

    // ── BOD_DATA에서 ROW 생성 + *_DATA 결과값 마이그레이션 ─────────────────
    /// <summary>
    /// 폐수의뢰및결과 데이터 재구축 (v2)
    /// — *_DATA 테이블의 '분석일' = 분석한 날 (DB 컬럼명 변경 완료)
    /// — 진짜 채수일(시료발생일)은 SN에서 추출 (MM-DD), 연도는 분석일 기준
    /// — SN이 시료 고유 식별자이므로 SN 기반 매칭
    /// </summary>
    private static void MigrateDataToWasteRequestResult(DbConnection conn)
    {
        const string T = "폐수의뢰및결과";

        try
        {
            using var chk = conn.CreateCommand();
            chk.CommandText = $"SELECT COUNT(*) FROM `{T}`";
            int total = Convert.ToInt32(chk.ExecuteScalar());

            if (total == 0)
            {
                // 비어있으면 전면 구축
                RebuildFromDataTables(conn, T);
                return;
            }

            // 빈 행 비율 확인
            using var chk2 = conn.CreateCommand();
            chk2.CommandText = $@"SELECT COUNT(*) FROM `{T}`
                WHERE (BOD IS NULL OR BOD = '') AND (`TOC` IS NULL OR `TOC` = '')
                  AND (SS IS NULL OR SS = '') AND (`T-N` IS NULL OR `T-N` = '')
                  AND (`T-P` IS NULL OR `T-P` = '') AND (`N-Hexan` IS NULL OR `N-Hexan` = '')
                  AND (Phenols IS NULL OR Phenols = '')";
            int empty = Convert.ToInt32(chk2.ExecuteScalar());

            if ((double)empty / total > 0.3)
            {
                // 빈 행 30% 이상 → 전면 재구축
                Log($"빈 행 {empty}/{total} ({100.0*empty/total:F1}%) → 전면 재구축");
                RebuildFromDataTables(conn, T);
                return;
            }

            // ── 보충 모드: 누락 행 추가 + 빈 행 수리
            // 1) BOD_DATA에 있지만 폐수의뢰및결과에 없는 SN 추가
            int added = AddMissingSamples(conn, T);
            if (added > 0) Log($"누락 시료 추가: {added}건");

            // 2) 빈 행이 있으면 SN 기반 수리
            if (empty > 0)
            {
                Log($"빈 행 {empty}건 → SN 기반 수리");
                RepairEmptyRows(conn);
            }
            else
            {
                Log($"마이그레이션 양호 (빈 행 0/{total})");
            }
        }
        catch (Exception ex) { Log($"마이그레이션 오류: {ex.Message}"); }
    }

    /// <summary>BOD_DATA에 있지만 폐수의뢰및결과에 없는 시료 추가 (SN+업체명 복합키)</summary>
    private static int AddMissingSamples(DbConnection conn, string T)
    {
        // 기존 (SN|업체명) 조합 로드
        var existing = new HashSet<string>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"SELECT SN, 업체명 FROM `{T}` WHERE SN IS NOT NULL AND SN <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string sn = r.IsDBNull(0) ? "" : r.GetString(0);
                string company = r.IsDBNull(1) ? "" : r.GetString(1);
                existing.Add($"{sn}|{company}");
            }
        }

        // BOD_DATA에서 없는 SN+업체명 추출
        var toAdd = new List<(string 채수일, string sn, string 업체명, string 구분)>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT DISTINCT 분석일, SN, 업체명 FROM `BOD_DATA` WHERE SN IS NOT NULL AND SN <> '' ORDER BY 분석일, SN";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string sn = r.GetString(1);
                string company = r.IsDBNull(2) ? "" : r.GetString(2);
                string key = $"{sn}|{company}";
                if (existing.Contains(key)) continue;
                string analysisDateStr = r.GetValue(0)?.ToString() ?? "";
                if (!DateTime.TryParse(analysisDateStr, out var ad)) continue;
                string samplingDate = ExtractSamplingDate(sn, ad);
                if (string.IsNullOrEmpty(samplingDate)) continue;
                string 구분 = sn.StartsWith("[세풍]") ? "세풍" : sn.StartsWith("[율촌]") ? "율촌" : "여수";
                toAdd.Add((samplingDate, sn, company, 구분));
                existing.Add(key);
            }
        }

        if (toAdd.Count == 0) return 0;

        // 순서 부여 및 INSERT
        var seqCounters = new Dictionary<string, int>();
        // 기존 최대 순서 로드
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"SELECT 채수일, 구분, MAX(순서) FROM `{T}` GROUP BY 채수일, 구분";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string d = r.IsDBNull(0) ? "" : r.GetString(0);
                string g = r.IsDBNull(1) ? "" : r.GetString(1);
                seqCounters[$"{d}|{g}"] = r.GetInt32(2);
            }
        }

        using var txn = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.Transaction = txn;
        ins.CommandText = $@"INSERT INTO `{T}` (채수일, 구분, 순서, SN, 업체명, 관리번호, BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols, 비고, 확인자)
            VALUES (@d, @g, @s, @sn, @name, '', '', '', '', '', '', '', '', '', '')";
        var pD = ins.CreateParameter(); pD.ParameterName = "@d"; ins.Parameters.Add(pD);
        var pG = ins.CreateParameter(); pG.ParameterName = "@g"; ins.Parameters.Add(pG);
        var pS = ins.CreateParameter(); pS.ParameterName = "@s"; ins.Parameters.Add(pS);
        var pSn = ins.CreateParameter(); pSn.ParameterName = "@sn"; ins.Parameters.Add(pSn);
        var pN = ins.CreateParameter(); pN.ParameterName = "@name"; ins.Parameters.Add(pN);
        ins.Prepare();

        int inserted = 0;
        foreach (var (date, sn, company, group) in toAdd)
        {
            string key = $"{date}|{group}";
            if (!seqCounters.ContainsKey(key)) seqCounters[key] = 0;
            seqCounters[key]++;
            pD.Value = date; pG.Value = group; pS.Value = seqCounters[key]; pSn.Value = sn; pN.Value = company;
            ins.ExecuteNonQuery();
            inserted++;
        }
        txn.Commit();

        // 추가된 행에 대해 결과값 매칭
        if (inserted > 0)
        {
            var mappings = new (string table, string srcCol, string destCol)[]
            {
                ("BOD_DATA", "결과", "BOD"), ("SS_DATA", "결과", "SS"),
                ("TN_DATA", "농도", "T-N"), ("TP_DATA", "농도", "T-P"),
                ("NHexan_DATA", "결과", "N-Hexan"), ("Phenols_DATA", "농도", "Phenols"),
            };
            foreach (var (table, srcCol, destCol) in mappings)
                MigrateBySnMatch(conn, T, table, srcCol, destCol);
            MigrateBySnMatch(conn, T, "TOC_TCIC_DATA", "검량선_a", "TOC");
            MigrateBySnMatchEmpty(conn, T, "TOC_NPOC_DATA", "검량선_a", "TOC");
        }

        return inserted;
    }

    /// <summary>
    /// SN에서 채수일(시료발생일) 추출: SN "02-24-07" + 분석일 연도 → "2026-02-24"
    /// — *_DATA.분석일에서 연도만 참고, MM-DD는 SN에서 추출
    /// — 12월 채수 → 1월 분석 케이스: 분석연도 - 1 = 채수연도
    /// </summary>
    private static string ExtractSamplingDate(string sn, DateTime analysisDate)
    {
        string clean = StripSnPrefix(sn);
        var parts = clean.Split('-');
        if (parts.Length < 2) return "";
        if (!int.TryParse(parts[0], out int mm) || !int.TryParse(parts[1], out int dd)) return "";
        if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return "";

        int year = analysisDate.Year;
        // 12월 채수 → 1월 분석: SN=12-XX인데 분석일이 1~2월이면 전년도
        if (mm >= 11 && analysisDate.Month <= 2) year--;
        // 1월 채수 → 12월 분석: SN=01-XX인데 분석일이 12월이면 다음연도 (드문 케이스)
        if (mm <= 2 && analysisDate.Month >= 11) year++;

        try { return new DateTime(year, mm, dd).ToString("yyyy-MM-dd"); }
        catch { return ""; }
    }

    /// <summary>
    /// 폐수의뢰및결과 전면 재구축
    /// 1) 기존 데이터 삭제
    /// 2) BOD_DATA에서 고유 시료 행 생성 (채수일은 SN에서 추출)
    /// 3) 각 DATA 테이블에서 SN 기반으로 결과값 매칭
    /// </summary>
    private static void RebuildFromDataTables(DbConnection conn, string T)
    {
        Log("=== 폐수의뢰및결과 전면 재구축 시작 ===");

        // ── Step 1: 기존 데이터 삭제
        Exec(conn, $"DELETE FROM `{T}`");
        Log("기존 데이터 전체 삭제");

        // ── Step 2: BOD_DATA에서 고유 시료 목록 추출 + INSERT
        //   BOD_DATA.분석일 = 분석한 날짜, SN = 시료번호 (MM-DD-NN), 업체명 = 업체명
        //   진짜 채수일 = SN의 MM-DD + 분석일 연도
        var samples = new List<(string 채수일, string sn, string 업체명, string 구분, int seq)>();
        var seen = new HashSet<string>();  // 채수일|SN|업체명 중복 방지
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT DISTINCT 분석일, SN, 업체명 FROM `BOD_DATA` WHERE SN IS NOT NULL AND SN <> '' ORDER BY 분석일, SN";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string analysisDateStr = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                string sn = r.GetString(1);
                string company = r.IsDBNull(2) ? "" : r.GetString(2);
                if (string.IsNullOrEmpty(sn)) continue;

                if (!DateTime.TryParse(analysisDateStr, out var analysisDate)) continue;

                string samplingDate = ExtractSamplingDate(sn, analysisDate);
                if (string.IsNullOrEmpty(samplingDate)) continue;

                // 채수일+SN+업체명 중복 방지 (다른 분석일이 같은 채수일로 매핑될 때)
                string dedupKey = $"{samplingDate}|{sn}|{company}";
                if (!seen.Add(dedupKey)) continue;

                string 구분 = "여수";
                if (sn.StartsWith("[세풍]")) 구분 = "세풍";
                else if (sn.StartsWith("[율촌]")) 구분 = "율촌";

                samples.Add((samplingDate, sn, company, 구분, 0));
            }
        }
        Log($"BOD_DATA 고유 시료: {samples.Count}건 (중복 제거 후)");

        // 채수일+구분별 순서 부여
        var seqCounters = new Dictionary<string, int>();
        var orderedSamples = new List<(string 채수일, string sn, string 업체명, string 구분, int seq)>();
        foreach (var s in samples)
        {
            string key = $"{s.채수일}|{s.구분}";
            if (!seqCounters.ContainsKey(key)) seqCounters[key] = 0;
            seqCounters[key]++;
            orderedSamples.Add((s.채수일, s.sn, s.업체명, s.구분, seqCounters[key]));
        }

        // 배치 INSERT
        using (var txn = conn.BeginTransaction())
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = txn;
            ins.CommandText = $@"
                INSERT INTO `{T}` (채수일, 구분, 순서, SN, 업체명, 관리번호, BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols, 비고, 확인자)
                VALUES (@d, @g, @s, @sn, @name, '', '', '', '', '', '', '', '', '', '')";
            var pD = ins.CreateParameter(); pD.ParameterName = "@d"; ins.Parameters.Add(pD);
            var pG = ins.CreateParameter(); pG.ParameterName = "@g"; ins.Parameters.Add(pG);
            var pS = ins.CreateParameter(); pS.ParameterName = "@s"; ins.Parameters.Add(pS);
            var pSn = ins.CreateParameter(); pSn.ParameterName = "@sn"; ins.Parameters.Add(pSn);
            var pName = ins.CreateParameter(); pName.ParameterName = "@name"; ins.Parameters.Add(pName);
            ins.Prepare();
            foreach (var (date, sn, company, group, seq) in orderedSamples)
            {
                pD.Value = date; pG.Value = group; pS.Value = seq; pSn.Value = sn; pName.Value = company;
                ins.ExecuteNonQuery();
            }
            txn.Commit();
        }
        Log($"시료 행 INSERT 완료: {orderedSamples.Count}건");

        // ── Step 3: 각 DATA 테이블에서 SN 기반으로 결과값 매칭
        var mappings = new (string table, string srcCol, string destCol)[]
        {
            ("BOD_DATA",      "결과",    "BOD"),
            ("SS_DATA",       "결과",    "SS"),
            ("TN_DATA",       "농도",    "T-N"),
            ("TP_DATA",       "농도",    "T-P"),
            ("NHexan_DATA",   "결과",    "N-Hexan"),
            ("Phenols_DATA",  "농도",    "Phenols"),
        };

        int totalUpdated = 0;
        foreach (var (table, srcCol, destCol) in mappings)
            totalUpdated += MigrateBySnMatch(conn, T, table, srcCol, destCol);

        // TOC: 1법(TC-IC) 우선, 2법(NPOC) 보충
        totalUpdated += MigrateBySnMatch(conn, T, "TOC_TCIC_DATA", "검량선_a", "TOC");
        totalUpdated += MigrateBySnMatchEmpty(conn, T, "TOC_NPOC_DATA", "검량선_a", "TOC");

        Log($"=== 재구축 완료: {orderedSamples.Count}행, {totalUpdated}건 결과값 갱신 ===");
    }

    /// <summary>
    /// 비고+업체명+30일 범위 매칭 (v5 — BOHEMCIGAR 매크로 완전 재현)
    /// — 폐수의뢰및결과.SN(접두사제거) = 의뢰SN (채수일 인코딩: MM-DD-NN)
    /// — *_DATA.비고(접두사제거) = 의뢰SN → 1순위 매칭키
    /// — *_DATA.SN(접두사제거) = 분석배치SN → 2순위 매칭키 (BOD처럼 당일 분석한 경우)
    /// — 채수일(SN에서 추출)로부터 +30일 이내 DATA.분석일만 검색
    /// </summary>
    private static int MigrateBySnMatch(DbConnection conn, string T, string srcTable, string srcCol, string destCol)
    {
        try
        {
            // 1) 폐수의뢰및결과 행 로드
            var targets = new List<(int id, string sn, string company, string samplingDate)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT Id, SN, 업체명, 채수일 FROM `{T}`";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.IsDBNull(1) ? "" : r.GetString(1);
                    string company = r.IsDBNull(2) ? "" : r.GetString(2);
                    string samplingDate = r.IsDBNull(3) ? "" : r.GetString(3);
                    if (!string.IsNullOrEmpty(sn))
                        targets.Add((r.GetInt32(0), sn, company, samplingDate));
                }
            }

            // 2) DATA 테이블 전체 로드: (비고접두사제거|업체명, 분석일) → 결과값
            //    + (SN|업체명, 분석일) → 결과값 (fallback)
            var dataRows = new List<(string cleanRemark, string cleanSn, string company, string val, DateTime analysisDate)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT SN, 업체명, `{srcCol}`, 분석일, 비고 FROM `{srcTable}` WHERE `{srcCol}` IS NOT NULL AND `{srcCol}` <> '' AND SN IS NOT NULL AND SN <> ''";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.GetString(0);
                    string company = r.IsDBNull(1) ? "" : r.GetString(1);
                    string val = r.GetValue(2)?.ToString() ?? "";
                    string rawDate = r.IsDBNull(3) ? "" : r.GetValue(3)?.ToString() ?? "";
                    string remark = r.IsDBNull(4) ? "" : r.GetString(4);
                    if (string.IsNullOrEmpty(val)) continue;
                    DateTime.TryParse(rawDate, out var analysisDate);
                    string cleanRemark = !string.IsNullOrEmpty(remark) ? StripSnPrefix(remark) : "";
                    string cleanSn = StripSnPrefix(sn);
                    dataRows.Add((cleanRemark, cleanSn, company, val, analysisDate));
                }
            }

            // 인덱스 구축: (cleanRemark|company) → rows, (cleanSn|company) → rows
            var byRemark = new Dictionary<string, List<int>>();
            var bySn = new Dictionary<string, List<int>>();
            for (int i = 0; i < dataRows.Count; i++)
            {
                var d = dataRows[i];
                if (!string.IsNullOrEmpty(d.cleanRemark))
                {
                    string rk = $"{d.cleanRemark}|{d.company}";
                    if (!byRemark.ContainsKey(rk)) byRemark[rk] = new();
                    byRemark[rk].Add(i);
                }
                string sk = $"{d.cleanSn}|{d.company}";
                if (!bySn.ContainsKey(sk)) bySn[sk] = new();
                bySn[sk].Add(i);
            }

            // 3) 매칭: 비고 우선 → SN fallback, 30일 범위 제한
            var updates = new List<(int id, string val)>();
            foreach (var (id, sn, company, samplingDate) in targets)
            {
                string cleanSn = StripSnPrefix(sn);
                if (!DateTime.TryParse(samplingDate, out var sampDate)) continue;
                DateTime limitDate = sampDate.AddDays(30);

                // 1순위: 비고(접두사제거)|업체명 + 30일 이내
                string? val = FindInRange(dataRows, byRemark, $"{cleanSn}|{company}", sampDate, limitDate);
                // 2순위: SN(접두사제거)|업체명 + 30일 이내
                if (val == null)
                    val = FindInRange(dataRows, bySn, $"{cleanSn}|{company}", sampDate, limitDate);
                if (val != null)
                    updates.Add((id, val));
            }

            // 4) 배치 UPDATE
            int rows = 0;
            if (updates.Count > 0)
            {
                using var txn = conn.BeginTransaction();
                using var upd = conn.CreateCommand();
                upd.Transaction = txn;
                upd.CommandText = $"UPDATE `{T}` SET `{destCol}` = @v WHERE Id = @id";
                var pV = upd.CreateParameter(); pV.ParameterName = "@v"; upd.Parameters.Add(pV);
                var pId = upd.CreateParameter(); pId.ParameterName = "@id"; upd.Parameters.Add(pId);
                upd.Prepare();
                foreach (var (id, val) in updates)
                {
                    pV.Value = val; pId.Value = id;
                    rows += upd.ExecuteNonQuery();
                }
                txn.Commit();
            }
            Log($"{srcTable}.{srcCol} → {destCol}: {rows}건");
            return rows;
        }
        catch (Exception ex) { Log($"{srcTable} 오류: {ex.Message}"); return 0; }
    }

    /// <summary>인덱스에서 키 찾고, 분석일이 범위 내인 첫 번째 결과 반환</summary>
    private static string? FindInRange(
        List<(string cleanRemark, string cleanSn, string company, string val, DateTime analysisDate)> dataRows,
        Dictionary<string, List<int>> index, string key,
        DateTime fromDate, DateTime toDate)
    {
        if (!index.TryGetValue(key, out var indices)) return null;
        // 범위 내에서 가장 가까운 분석일 결과
        string? bestVal = null;
        double bestDiff = double.MaxValue;
        foreach (int i in indices)
        {
            var d = dataRows[i];
            if (d.analysisDate >= fromDate && d.analysisDate <= toDate)
            {
                double diff = (d.analysisDate - fromDate).TotalDays;
                if (diff < bestDiff) { bestDiff = diff; bestVal = d.val; }
            }
        }
        return bestVal;
    }

    /// <summary>TOC 보충용: 기존 값이 비어있는 행만 업데이트 (비고+업체명+30일 범위)</summary>
    private static int MigrateBySnMatchEmpty(DbConnection conn, string T, string srcTable, string srcCol, string destCol)
    {
        try
        {
            var targets = new List<(int id, string sn, string company, string samplingDate)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT Id, SN, 업체명, 채수일 FROM `{T}` WHERE (`{destCol}` IS NULL OR `{destCol}` = '')";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.IsDBNull(1) ? "" : r.GetString(1);
                    string company = r.IsDBNull(2) ? "" : r.GetString(2);
                    string samplingDate = r.IsDBNull(3) ? "" : r.GetString(3);
                    if (!string.IsNullOrEmpty(sn))
                        targets.Add((r.GetInt32(0), sn, company, samplingDate));
                }
            }

            var dataRows = new List<(string cleanRemark, string cleanSn, string company, string val, DateTime analysisDate)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT SN, 업체명, `{srcCol}`, 분석일, 비고 FROM `{srcTable}` WHERE `{srcCol}` IS NOT NULL AND `{srcCol}` <> '' AND SN IS NOT NULL AND SN <> ''";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.GetString(0);
                    string company = r.IsDBNull(1) ? "" : r.GetString(1);
                    string val = r.GetValue(2)?.ToString() ?? "";
                    string rawDate = r.IsDBNull(3) ? "" : r.GetValue(3)?.ToString() ?? "";
                    string remark = r.IsDBNull(4) ? "" : r.GetString(4);
                    if (string.IsNullOrEmpty(val)) continue;
                    DateTime.TryParse(rawDate, out var analysisDate);
                    dataRows.Add((!string.IsNullOrEmpty(remark) ? StripSnPrefix(remark) : "", StripSnPrefix(sn), company, val, analysisDate));
                }
            }

            var byRemark = new Dictionary<string, List<int>>();
            var bySn = new Dictionary<string, List<int>>();
            for (int i = 0; i < dataRows.Count; i++)
            {
                var d = dataRows[i];
                if (!string.IsNullOrEmpty(d.cleanRemark))
                {
                    string rk = $"{d.cleanRemark}|{d.company}";
                    if (!byRemark.ContainsKey(rk)) byRemark[rk] = new();
                    byRemark[rk].Add(i);
                }
                string sk = $"{d.cleanSn}|{d.company}";
                if (!bySn.ContainsKey(sk)) bySn[sk] = new();
                bySn[sk].Add(i);
            }

            var updates = new List<(int id, string val)>();
            foreach (var (id, sn, company, samplingDate) in targets)
            {
                string cleanSn = StripSnPrefix(sn);
                if (!DateTime.TryParse(samplingDate, out var sampDate)) continue;
                DateTime limitDate = sampDate.AddDays(30);
                string? val = FindInRange(dataRows, byRemark, $"{cleanSn}|{company}", sampDate, limitDate)
                           ?? FindInRange(dataRows, bySn, $"{cleanSn}|{company}", sampDate, limitDate);
                if (val != null)
                    updates.Add((id, val));
            }

            int rows = 0;
            if (updates.Count > 0)
            {
                using var txn = conn.BeginTransaction();
                using var upd = conn.CreateCommand();
                upd.Transaction = txn;
                upd.CommandText = $"UPDATE `{T}` SET `{destCol}` = @v WHERE Id = @id";
                var pV = upd.CreateParameter(); pV.ParameterName = "@v"; upd.Parameters.Add(pV);
                var pId = upd.CreateParameter(); pId.ParameterName = "@id"; upd.Parameters.Add(pId);
                upd.Prepare();
                foreach (var (id, val) in updates)
                { pV.Value = val; pId.Value = id; rows += upd.ExecuteNonQuery(); }
                txn.Commit();
            }
            Log($"{srcTable}.{srcCol} → {destCol} (보충): {rows}건");
            return rows;
        }
        catch (Exception ex) { Log($"{srcTable} 보충 오류: {ex.Message}"); return 0; }
    }

    /// <summary>복합키(SN|업체명)+연도 기반 매칭</summary>
    private static string? MatchByKeyYear(Dictionary<string, List<(string val, int year)>> byKey, string key, int targetYear)
    {
        if (!byKey.TryGetValue(key, out var candidates)) return null;
        if (candidates.Count == 1) return candidates[0].val;

        // 같은 연도 우선, ±1년 허용
        string? bestVal = null;
        int bestDiff = int.MaxValue;
        foreach (var (val, year) in candidates)
        {
            int diff = year == 0 ? 0 : Math.Abs(year - targetYear);
            if (diff < bestDiff) { bestDiff = diff; bestVal = val; }
        }
        return bestVal;
    }

    /// <summary>SN만으로 매칭 (fallback): 업체명 무관하게 같은 SN 중 연도 가장 가까운 것</summary>
    private static string? MatchBySnOnlyYear(Dictionary<string, List<(string val, int year)>> byKey, string sn, int targetYear)
    {
        // byKey의 키가 "SN|업체명" 형태이므로 SN이 일치하는 항목을 모두 수집
        string? bestVal = null;
        int bestDiff = int.MaxValue;
        foreach (var (key, candidates) in byKey)
        {
            // key = "SN|업체명" → SN 부분만 비교
            int sep = key.IndexOf('|');
            string keySn = sep >= 0 ? key[..sep] : key;
            if (keySn != sn) continue;

            foreach (var (val, year) in candidates)
            {
                int diff = year == 0 ? 0 : Math.Abs(year - targetYear);
                if (diff < bestDiff) { bestDiff = diff; bestVal = val; }
            }
        }
        return bestDiff <= 1 ? bestVal : null;  // ±1년 이내만 허용
    }

    /// <summary>
    /// BOD_DATA에서 고유 (채수일, SN, 업체명) 추출하여 폐수의뢰및결과에 없는 ROW 추가
    /// </summary>
    private static int InsertRowsFromBodData(DbConnection conn, string T)
    {
        try
        {
            // 1) BOD_DATA에서 고유 (분석일, SN, 업체명) 추출
            var bodRows = new List<(string 채수일, string sn, string 업체명)>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT DISTINCT 분석일, SN, 업체명 FROM `BOD_DATA` WHERE SN IS NOT NULL AND SN <> '' ORDER BY 분석일, SN";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string date = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                    string sn = r.IsDBNull(1) ? "" : r.GetString(1);
                    string company = r.IsDBNull(2) ? "" : r.GetString(2);
                    if (string.IsNullOrEmpty(date) || string.IsNullOrEmpty(sn)) continue;
                    // 날짜 정규화
                    if (DateTime.TryParse(date, out var dt))
                        date = dt.ToString("yyyy-MM-dd");
                    bodRows.Add((date, sn, company));
                }
            }
            Log($"BOD_DATA 고유 행: {bodRows.Count}건");

            // 2) 기존 폐수의뢰및결과의 (SN, 채수일연도) 조합 로드
            var existing = new HashSet<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT SN, 채수일 FROM `{T}`";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.IsDBNull(0) ? "" : r.GetString(0);
                    string date = r.IsDBNull(1) ? "" : r.GetString(1);
                    if (!string.IsNullOrEmpty(sn) && !string.IsNullOrEmpty(date))
                        existing.Add($"{sn}|{date}");
                }
            }
            Log($"기존 폐수의뢰및결과 행: {existing.Count}건");

            // 3) 없는 행만 INSERT
            int inserted = 0;
            // 날짜+구분별 순서 카운터
            var seqCounters = new Dictionary<string, int>();

            // 기존 최대 순서 로드
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT 채수일, 구분, MAX(순서) FROM `{T}` GROUP BY 채수일, 구분";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string d = r.IsDBNull(0) ? "" : r.GetString(0);
                    string g = r.IsDBNull(1) ? "" : r.GetString(1);
                    int maxSeq = r.GetInt32(2);
                    seqCounters[$"{d}|{g}"] = maxSeq;
                }
            }

            // INSERT할 행 준비
            var toInsert = new List<(string date, string 구분, int seq, string sn, string company)>();
            foreach (var (채수일, sn, 업체명) in bodRows)
            {
                string 구분 = "여수";
                if (sn.StartsWith("[세풍]")) 구분 = "세풍";
                else if (sn.StartsWith("[율촌]")) 구분 = "율촌";

                if (existing.Contains($"{sn}|{채수일}")) continue;

                string seqKey = $"{채수일}|{구분}";
                if (!seqCounters.ContainsKey(seqKey)) seqCounters[seqKey] = 0;
                seqCounters[seqKey]++;
                int seq = seqCounters[seqKey];

                // 원본 SN 보존 — BuildSN 대신 BOD_DATA 원래 SN 사용 (DATA 테이블 매칭용)
                toInsert.Add((채수일, 구분, seq, sn, 업체명));
                existing.Add($"{sn}|{채수일}");
            }

            Log($"INSERT 대상: {toInsert.Count}건, 트랜잭션 시작...");

            // 트랜잭션으로 배치 INSERT
            using var txn = conn.BeginTransaction();
            using var ins = conn.CreateCommand();
            ins.Transaction = txn;
            ins.CommandText = $@"
                INSERT INTO `{T}` (채수일, 구분, 순서, SN, 업체명, 관리번호, BOD, `TOC`, SS, `T-N`, `T-P`, `N-Hexan`, Phenols, 비고, 확인자)
                VALUES (@d, @g, @s, @sn, @name, '', '', '', '', '', '', '', '', '', '')";
            var pD = ins.CreateParameter(); pD.ParameterName = "@d"; ins.Parameters.Add(pD);
            var pG = ins.CreateParameter(); pG.ParameterName = "@g"; ins.Parameters.Add(pG);
            var pS = ins.CreateParameter(); pS.ParameterName = "@s"; ins.Parameters.Add(pS);
            var pSn = ins.CreateParameter(); pSn.ParameterName = "@sn"; ins.Parameters.Add(pSn);
            var pName = ins.CreateParameter(); pName.ParameterName = "@name"; ins.Parameters.Add(pName);
            ins.Prepare();

            foreach (var (date, 구분, seq, sn, company) in toInsert)
            {
                pD.Value = date; pG.Value = 구분; pS.Value = seq; pSn.Value = sn; pName.Value = company;
                ins.ExecuteNonQuery();
                inserted++;
            }
            txn.Commit();
            Log($"트랜잭션 커밋 완료: {inserted}건 INSERT");

            return inserted;
        }
        catch (Exception ex) { Log($"InsertRowsFromBodData 오류: {ex.Message}"); return 0; }
    }

    /// <summary>
    /// SN(접두사 제거) + 업체명(유사도 95%) 으로 C# 루프 매칭하여 DATA 결과값을 폐수의뢰및결과에 반영
    /// </summary>
    private static int MigrateBySnJoin(DbConnection conn, string T, string srcTable, string srcCol, string destCol, string? onlyEmptyCol)
    {
        try
        {
            // 1) 폐수의뢰및결과 행 로드 (전체: 여수/율촌/세풍)
            var targets = new List<(int id, string sn, string company, DateTime date)>();
            using (var cmd = conn.CreateCommand())
            {
                string where = onlyEmptyCol != null
                    ? $" WHERE (`{onlyEmptyCol}` IS NULL OR `{onlyEmptyCol}` = '')"
                    : "";
                cmd.CommandText = $"SELECT Id, SN, 업체명, 채수일 FROM `{T}`{where}";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.IsDBNull(1) ? "" : r.GetString(1);
                    string dateStr = r.IsDBNull(3) ? "" : r.GetString(3);
                    if (!DateTime.TryParse(dateStr, out var dt)) continue;
                    targets.Add((r.GetInt32(0), sn, r.IsDBNull(2) ? "" : r.GetString(2), dt));
                }
            }

            // 2) DATA 테이블에서 SN → (결과값, 업체명, 분석일) 딕셔너리
            var sources = new Dictionary<string, List<(string value, string company, DateTime date)>>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT SN, `{srcCol}`, 업체명, 분석일 FROM `{srcTable}` WHERE `{srcCol}` IS NOT NULL";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string sn = r.IsDBNull(0) ? "" : r.GetString(0);
                    string val = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString() ?? "";
                    string co = r.IsDBNull(2) ? "" : r.GetString(2);
                    var rawDate = r.IsDBNull(3) ? "" : r.GetValue(3)?.ToString() ?? "";
                    if (string.IsNullOrEmpty(sn) || string.IsNullOrEmpty(val)) continue;
                    if (!DateTime.TryParse(rawDate, out var dt)) continue;
                    if (!sources.ContainsKey(sn)) sources[sn] = new();
                    sources[sn].Add((val, co, dt));
                }
            }

            Log($"{srcTable}: targets={targets.Count}, sources={sources.Count}");

            // 매칭 결과 수집
            var updates = new List<(int id, string val)>();
            foreach (var (id, sn, company, targetDate) in targets)
            {
                string cleanSn = StripSnPrefix(sn);
                if (!sources.TryGetValue(cleanSn, out var candidates)) continue;

                // SN 매칭: 같은 연도 우선, 없으면 가장 가까운 날짜
                string? bestVal = null;
                double bestDayDiff = double.MaxValue;
                bool bestSameYear = false;
                foreach (var (val, srcCompany, srcDate) in candidates)
                {
                    bool sameYear = srcDate.Year == targetDate.Year;
                    double dayDiff = Math.Abs((srcDate - targetDate).TotalDays);
                    if (sameYear && !bestSameYear)
                    {
                        bestSameYear = true;
                        bestDayDiff = dayDiff;
                        bestVal = val;
                    }
                    else if (sameYear == bestSameYear && dayDiff < bestDayDiff)
                    {
                        bestDayDiff = dayDiff;
                        bestVal = val;
                    }
                }
                if (bestVal != null)
                    updates.Add((id, bestVal));
            }

            // 트랜잭션으로 배치 UPDATE
            int rows = 0;
            if (updates.Count > 0)
            {
                using var txn = conn.BeginTransaction();
                using var upd = conn.CreateCommand();
                upd.Transaction = txn;
                upd.CommandText = $"UPDATE `{T}` SET `{destCol}` = @v WHERE Id = @id";
                var pV = upd.CreateParameter(); pV.ParameterName = "@v"; upd.Parameters.Add(pV);
                var pId = upd.CreateParameter(); pId.ParameterName = "@id"; upd.Parameters.Add(pId);
                upd.Prepare();

                foreach (var (id, val) in updates)
                {
                    pV.Value = val; pId.Value = id;
                    rows += upd.ExecuteNonQuery();
                }
                txn.Commit();
            }

            string suffix = onlyEmptyCol != null ? " (보충)" : "";
            Log($"마이그레이션 {srcTable}.{srcCol} -> {destCol}{suffix}: {rows}건");
            return rows;
        }
        catch (Exception ex) { Log($"마이그레이션 {srcTable} 오류: {ex.Message}"); return 0; }
    }

    /// <summary>
    /// 빈 결과 행 수리: SN 기반 매칭으로 DATA 테이블에서 보충
    /// — DATA.분석일 = 분석한 날, 폐수의뢰및결과.채수일 = 채수일(시료발생일)이므로
    ///   날짜가 아닌 SN(시료번호)으로 매칭해야 함
    /// </summary>
    public static int RepairEmptyRows(DbConnection conn)
    {
        const string T = "폐수의뢰및결과";
        int totalRepaired = 0;

        // 1) 빈 행(모든 분석항목이 비어있는 행) 로드
        var emptyRows = new List<(int id, string 업체명, string 채수일, string sn)>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                SELECT Id, 업체명, 채수일, SN FROM `{T}`
                WHERE (BOD IS NULL OR BOD = '')
                  AND (`TOC` IS NULL OR `TOC` = '')
                  AND (SS IS NULL OR SS = '')
                  AND (`T-N` IS NULL OR `T-N` = '')
                  AND (`T-P` IS NULL OR `T-P` = '')
                  AND (`N-Hexan` IS NULL OR `N-Hexan` = '')
                  AND (Phenols IS NULL OR Phenols = '')";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                emptyRows.Add((
                    r.GetInt32(0),
                    r.IsDBNull(1) ? "" : r.GetString(1),
                    r.IsDBNull(2) ? "" : r.GetString(2),
                    r.IsDBNull(3) ? "" : r.GetString(3)));
            }
        }
        Log($"빈 행 수리 대상: {emptyRows.Count}건");
        if (emptyRows.Count == 0) return 0;

        // 2) 각 DATA 테이블에서 SN 기반으로 결과값 매칭
        var mappings = new (string table, string srcCol, string destCol)[]
        {
            ("BOD_DATA",      "결과",    "BOD"),
            ("SS_DATA",       "결과",    "SS"),
            ("TN_DATA",       "농도",    "T-N"),
            ("TP_DATA",       "농도",    "T-P"),
            ("NHexan_DATA",   "결과",    "N-Hexan"),
            ("Phenols_DATA",  "농도",    "Phenols"),
        };

        foreach (var (table, srcCol, destCol) in mappings)
        {
            try
            {
                // DATA 테이블 로드 + 비고/SN 인덱스 구축
                var dataRows = new List<(string cleanRemark, string cleanSn, string company, string val, DateTime analysisDate)>();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $"SELECT SN, 업체명, `{srcCol}`, 분석일, 비고 FROM `{table}` WHERE `{srcCol}` IS NOT NULL AND `{srcCol}` <> '' AND SN IS NOT NULL AND SN <> ''";
                    using var r = cmd.ExecuteReader();
                    while (r.Read())
                    {
                        string sn = r.GetString(0);
                        string company = r.IsDBNull(1) ? "" : r.GetString(1);
                        string val = r.GetValue(2)?.ToString() ?? "";
                        string rawDate = r.IsDBNull(3) ? "" : r.GetValue(3)?.ToString() ?? "";
                        string remark = r.IsDBNull(4) ? "" : r.GetString(4);
                        if (string.IsNullOrEmpty(val)) continue;
                        DateTime.TryParse(rawDate, out var ad);
                        dataRows.Add((!string.IsNullOrEmpty(remark) ? StripSnPrefix(remark) : "", StripSnPrefix(sn), company, val, ad));
                    }
                }
                var byRemark = new Dictionary<string, List<int>>();
                var bySn2 = new Dictionary<string, List<int>>();
                for (int i = 0; i < dataRows.Count; i++)
                {
                    var d = dataRows[i];
                    if (!string.IsNullOrEmpty(d.cleanRemark))
                    { string rk = $"{d.cleanRemark}|{d.company}"; if (!byRemark.ContainsKey(rk)) byRemark[rk] = new(); byRemark[rk].Add(i); }
                    string sk = $"{d.cleanSn}|{d.company}";
                    if (!bySn2.ContainsKey(sk)) bySn2[sk] = new(); bySn2[sk].Add(i);
                }

                // 빈 행에 대해 비고+30일 매칭
                var updates = new List<(int id, string val)>();
                foreach (var (id, 업체명, 채수일, sn) in emptyRows)
                {
                    if (string.IsNullOrEmpty(sn)) continue;
                    if (!DateTime.TryParse(채수일, out var sampDate)) continue;
                    string cleanSn = StripSnPrefix(sn);
                    DateTime limitDate = sampDate.AddDays(30);
                    string? matched = FindInRange(dataRows, byRemark, $"{cleanSn}|{업체명}", sampDate, limitDate)
                                   ?? FindInRange(dataRows, bySn2, $"{cleanSn}|{업체명}", sampDate, limitDate);
                    if (matched != null)
                        updates.Add((id, matched));
                }

                // UPDATE 적용
                if (updates.Count > 0)
                {
                    using var txn = conn.BeginTransaction();
                    using var upd = conn.CreateCommand();
                    upd.Transaction = txn;
                    upd.CommandText = $"UPDATE `{T}` SET `{destCol}` = @v WHERE Id = @id";
                    var pV = upd.CreateParameter(); pV.ParameterName = "@v"; upd.Parameters.Add(pV);
                    var pId = upd.CreateParameter(); pId.ParameterName = "@id"; upd.Parameters.Add(pId);
                    upd.Prepare();
                    foreach (var (id, val) in updates)
                    {
                        pV.Value = val; pId.Value = id;
                        upd.ExecuteNonQuery();
                    }
                    txn.Commit();
                    totalRepaired += updates.Count;
                    Log($"빈 행 수리 {table}.{srcCol} → {destCol}: {updates.Count}건");
                }
            }
            catch (Exception ex) { Log($"빈 행 수리 {table} 오류: {ex.Message}"); }
        }

        // TOC 별도 처리: 1법 우선, 2법 보충
        totalRepaired += RepairEmptyToc(conn, emptyRows);

        Log($"빈 행 수리 완료: 총 {totalRepaired}건");
        return totalRepaired;
    }

    private static int RepairEmptyToc(DbConnection conn, List<(int id, string 업체명, string 채수일, string sn)> emptyRows)
    {
        const string T = "폐수의뢰및결과";
        int repaired = 0;

        var tocEmptyIds = new HashSet<int>();
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"SELECT Id FROM `{T}` WHERE (`TOC` IS NULL OR `TOC` = '')";
            using var r = cmd.ExecuteReader();
            while (r.Read()) tocEmptyIds.Add(r.GetInt32(0));
        }
        var tocTargets = emptyRows.Where(e => tocEmptyIds.Contains(e.id)).ToList();
        if (tocTargets.Count == 0) return 0;

        // 1법(TC-IC) 우선, 2법(NPOC) 보충
        foreach (var (table, srcCol) in new[] { ("TOC_TCIC_DATA", "검량선_a"), ("TOC_NPOC_DATA", "검량선_a") })
        {
            try
            {
                var dataRows = new List<(string cleanRemark, string cleanSn, string company, string val, DateTime analysisDate)>();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $"SELECT SN, 업체명, `{srcCol}`, 분석일, 비고 FROM `{table}` WHERE `{srcCol}` IS NOT NULL AND `{srcCol}` <> '' AND SN IS NOT NULL AND SN <> ''";
                    using var r = cmd.ExecuteReader();
                    while (r.Read())
                    {
                        string sn = r.GetString(0);
                        string company = r.IsDBNull(1) ? "" : r.GetString(1);
                        string val = r.GetValue(2)?.ToString() ?? "";
                        string rawDate = r.IsDBNull(3) ? "" : r.GetValue(3)?.ToString() ?? "";
                        string remark = r.IsDBNull(4) ? "" : r.GetString(4);
                        if (string.IsNullOrEmpty(val)) continue;
                        DateTime.TryParse(rawDate, out var ad);
                        dataRows.Add((!string.IsNullOrEmpty(remark) ? StripSnPrefix(remark) : "", StripSnPrefix(sn), company, val, ad));
                    }
                }
                var byRemark = new Dictionary<string, List<int>>();
                var bySn2 = new Dictionary<string, List<int>>();
                for (int i = 0; i < dataRows.Count; i++)
                {
                    var d = dataRows[i];
                    if (!string.IsNullOrEmpty(d.cleanRemark))
                    { string rk = $"{d.cleanRemark}|{d.company}"; if (!byRemark.ContainsKey(rk)) byRemark[rk] = new(); byRemark[rk].Add(i); }
                    string sk = $"{d.cleanSn}|{d.company}";
                    if (!bySn2.ContainsKey(sk)) bySn2[sk] = new(); bySn2[sk].Add(i);
                }

                var updates = new List<(int id, string val)>();
                foreach (var (id, 업체명, 채수일, sn) in tocTargets)
                {
                    if (!tocEmptyIds.Contains(id)) continue;
                    if (string.IsNullOrEmpty(sn)) continue;
                    if (!DateTime.TryParse(채수일, out var sampDate)) continue;
                    string cleanSn = StripSnPrefix(sn);
                    DateTime limitDate = sampDate.AddDays(30);
                    string? matched = FindInRange(dataRows, byRemark, $"{cleanSn}|{업체명}", sampDate, limitDate)
                                   ?? FindInRange(dataRows, bySn2, $"{cleanSn}|{업체명}", sampDate, limitDate);
                    if (matched != null)
                    { updates.Add((id, matched)); tocEmptyIds.Remove(id); }
                }

                if (updates.Count > 0)
                {
                    using var txn = conn.BeginTransaction();
                    using var upd = conn.CreateCommand();
                    upd.Transaction = txn;
                    upd.CommandText = $"UPDATE `{T}` SET `TOC` = @v WHERE Id = @id";
                    var pV = upd.CreateParameter(); pV.ParameterName = "@v"; upd.Parameters.Add(pV);
                    var pId = upd.CreateParameter(); pId.ParameterName = "@id"; upd.Parameters.Add(pId);
                    upd.Prepare();
                    foreach (var (id, val) in updates)
                    { pV.Value = val; pId.Value = id; upd.ExecuteNonQuery(); }
                    txn.Commit();
                    repaired += updates.Count;
                    Log($"빈 행 TOC 수리 ({table}): {updates.Count}건");
                }
            }
            catch (Exception ex) { Log($"빈 행 TOC 수리 {table} 오류: {ex.Message}"); }
        }
        return repaired;
    }


    /// <summary>SN 접두사 [세풍] [율촌] 제거</summary>
    private static string StripSnPrefix(string sn)
        => sn.Replace("[세풍]", "").Replace("[율촌]", "");

    /// <summary>두 문자열의 유사도 (0~1). Levenshtein 기반.</summary>
    private static double Similarity(string a, string b)
    {
        if (a == b) return 1.0;
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
        int len = Math.Max(a.Length, b.Length);
        return 1.0 - (double)LevenshteinDistance(a, b) / len;
    }

    private static int LevenshteinDistance(string s, string t)
    {
        int n = s.Length, m = t.Length;
        var d = new int[n + 1, m + 1];
        for (int i = 0; i <= n; i++) d[i, 0] = i;
        for (int j = 0; j <= m; j++) d[0, j] = j;
        for (int i = 1; i <= n; i++)
            for (int j = 1; j <= m; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        return d[n, m];
    }

    private static bool ColumnExists(DbConnection conn, string table, string column)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `{column}` FROM `{table}` LIMIT 0";
            cmd.ExecuteNonQuery();
            return true;
        }
        catch { return false; }
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static void Exec(DbConnection conn, string sql)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    // ── 마이그레이션 버전 마커 ────────────────────────────────────────
    private static void EnsureMigrationTable(DbConnection conn)
    {
        Exec(conn, "CREATE TABLE IF NOT EXISTS `_eta_migration` (`key` VARCHAR(255) PRIMARY KEY, done_at TEXT)");
    }

    private static bool IsMigrationDone(DbConnection conn, string key)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM `_eta_migration` WHERE key=@k";
            cmd.Parameters.AddWithValue("@k", key);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
        catch { return false; }
    }

    private static void MarkMigrationDone(DbConnection conn, string key)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT OR REPLACE INTO `_eta_migration` (key, done_at) VALUES (@k, @t)";
        cmd.Parameters.AddWithValue("@k", key);
        cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();
        Log($"마이그레이션 완료 표시: {key}");
    }

    private static readonly string LogPath = System.IO.Path.GetFullPath(
        System.IO.Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "FacilityMigration.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [FacilityMigration] {msg}";
        Debug.WriteLine(line);
        try { System.IO.File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }
}
