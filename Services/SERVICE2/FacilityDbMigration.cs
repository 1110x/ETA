using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
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
    }

    // ── 처리시설_마스터 ────────────────────────────────────────────────────
    private static void EnsureProcessingFacilityMaster(DbConnection conn)
    {
        if (DbConnectionFactory.TableExists(conn, "처리시설_마스터")) return;

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
        SeedFacilityMaster(conn);
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

            // ── 세룡 ────────────────────────────────────────────────────────
            ("세룡", "유입수",                   "O","O","O","O","O","O","","","","","", ""),
            ("세룡", "유입수",                   "O","","","","","","","","","","", ""),
            ("세룡", "유입수",                   "O","","","","","","","","","","", ""),
            ("세룡", "분배조",                   "O","O","O","O","O","","","","","","", ""),
            ("세룡", "방류수",                   "O","O","O","O","O","O","","","","","", ""),

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
                    의뢰번호    {(DbConnectionFactory.IsMariaDb ? "VARCHAR(191)" : "TEXT")} NOT NULL UNIQUE,
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
                    채취일자    {(DbConnectionFactory.IsMariaDb ? "VARCHAR(20)" : "TEXT")} NOT NULL,
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
                    관리번호    {(DbConnectionFactory.IsMariaDb ? "VARCHAR(191)" : "TEXT")} NOT NULL,
                    업체명      TEXT NOT NULL,
                    채취일자    {(DbConnectionFactory.IsMariaDb ? "VARCHAR(50)" : "TEXT")} NOT NULL,
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

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static void Exec(DbConnection conn, string sql)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    private static void Log(string msg)
        => Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] [FacilityMigration] {msg}");
}
