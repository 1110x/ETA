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
        EnsureYeosuCompanyTable(conn);
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
