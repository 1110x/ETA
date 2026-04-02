using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using ETA.Models;

namespace ETA.Services.Common;

public static class ScheduleService
{
    // ── 테이블 생성 ────────────────────────────────────────────────────────────
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `일정` (
                Id       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                날짜     {(DbConnectionFactory.IsMariaDb ? "VARCHAR(10)" : "TEXT")} NOT NULL,
                직원명   TEXT DEFAULT '',
                직원id   TEXT DEFAULT '',
                분류     TEXT NOT NULL DEFAULT '출장',
                사이트   TEXT DEFAULT '',
                업체약칭 TEXT DEFAULT '',
                제목     TEXT DEFAULT '',
                내용     TEXT DEFAULT '',
                시작시간 TEXT DEFAULT '',
                종료시간 TEXT DEFAULT '',
                첨부파일 TEXT DEFAULT '',
                등록일시 TEXT NOT NULL DEFAULT '',
                등록자   TEXT DEFAULT ''
            )";
        cmd.ExecuteNonQuery();

        // 기존 테이블에 업체약칭 컬럼 없으면 추가
        if (!DbConnectionFactory.ColumnExists(conn, "일정", "업체약칭"))
        {
            using var alt = conn.CreateCommand();
            alt.CommandText = "ALTER TABLE `일정` ADD COLUMN `업체약칭` TEXT DEFAULT ''";
            try { alt.ExecuteNonQuery(); } catch { }
        }
    }

    private static string S(DbDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);

    // ── 월별 조회 ─────────────────────────────────────────────────────────────
    public static List<ScheduleEntry> GetByMonth(int year, int month)
    {
        var list = new List<ScheduleEntry>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT Id, 날짜, 직원명, 직원id, 분류, 사이트, 업체약칭, 제목, 내용,
                   시작시간, 종료시간, 첨부파일, 등록일시, 등록자
            FROM `일정`
            WHERE 날짜 LIKE @prefix
            ORDER BY 날짜, Id";
        cmd.Parameters.AddWithValue("@prefix", $"{year:D4}-{month:D2}%");

        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        Debug.WriteLine($"[Schedule] {year}-{month} 로드 {list.Count}건");
        return list;
    }

    // ── 날짜별 조회 ───────────────────────────────────────────────────────────
    public static List<ScheduleEntry> GetByDate(string date)
    {
        var list = new List<ScheduleEntry>();
        if (string.IsNullOrEmpty(date)) return list;
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT Id, 날짜, 직원명, 직원id, 분류, 사이트, 업체약칭, 제목, 내용,
                   시작시간, 종료시간, 첨부파일, 등록일시, 등록자
            FROM `일정`
            WHERE 날짜 = @date
            ORDER BY Id";
        cmd.Parameters.AddWithValue("@date", date);

        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    // ── 삽입 ─────────────────────────────────────────────────────────────────
    public static void Insert(ScheduleEntry e)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `일정`
                (날짜, 직원명, 직원id, 분류, 사이트, 업체약칭, 제목, 내용,
                 시작시간, 종료시간, 첨부파일, 등록일시, 등록자)
            VALUES
                (@날짜, @직원명, @직원id, @분류, @사이트, @업체약칭, @제목, @내용,
                 @시작시간, @종료시간, @첨부파일, @등록일시, @등록자)";
        cmd.Parameters.AddWithValue("@날짜",     e.날짜);
        cmd.Parameters.AddWithValue("@직원명",   e.직원명);
        cmd.Parameters.AddWithValue("@직원id",   e.직원id);
        cmd.Parameters.AddWithValue("@분류",     e.분류);
        cmd.Parameters.AddWithValue("@사이트",   e.사이트);
        cmd.Parameters.AddWithValue("@업체약칭", e.업체약칭);
        cmd.Parameters.AddWithValue("@제목",     e.제목);
        cmd.Parameters.AddWithValue("@내용",     e.내용);
        cmd.Parameters.AddWithValue("@시작시간", e.시작시간);
        cmd.Parameters.AddWithValue("@종료시간", e.종료시간);
        cmd.Parameters.AddWithValue("@첨부파일", e.첨부파일);
        cmd.Parameters.AddWithValue("@등록일시", e.등록일시);
        cmd.Parameters.AddWithValue("@등록자",   e.등록자);
        cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Schedule] 삽입: {e.날짜} {e.분류} {e.업체약칭} {e.직원명}");
    }

    // ── 수정 ─────────────────────────────────────────────────────────────────
    public static void Update(ScheduleEntry e)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `일정` SET
                날짜=@날짜, 직원명=@직원명, 직원id=@직원id,
                분류=@분류, 사이트=@사이트, 업체약칭=@업체약칭, 제목=@제목, 내용=@내용,
                시작시간=@시작시간, 종료시간=@종료시간, 첨부파일=@첨부파일
            WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id",     e.Id);
        cmd.Parameters.AddWithValue("@날짜",     e.날짜);
        cmd.Parameters.AddWithValue("@직원명",   e.직원명);
        cmd.Parameters.AddWithValue("@직원id",   e.직원id);
        cmd.Parameters.AddWithValue("@분류",     e.분류);
        cmd.Parameters.AddWithValue("@사이트",   e.사이트);
        cmd.Parameters.AddWithValue("@업체약칭", e.업체약칭);
        cmd.Parameters.AddWithValue("@제목",     e.제목);
        cmd.Parameters.AddWithValue("@내용",     e.내용);
        cmd.Parameters.AddWithValue("@시작시간", e.시작시간);
        cmd.Parameters.AddWithValue("@종료시간", e.종료시간);
        cmd.Parameters.AddWithValue("@첨부파일", e.첨부파일);
        cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Schedule] 수정: Id={e.Id} {e.날짜} {e.분류} {e.직원명}");
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────────
    public static void Delete(int id)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `일정` WHERE Id = @id";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── 매핑 ─────────────────────────────────────────────────────────────────
    private static ScheduleEntry Map(DbDataReader r) => new()
    {
        Id       = r.IsDBNull(0) ? 0 : r.GetInt32(0),
        날짜     = S(r, 1),
        직원명   = S(r, 2),
        직원id   = S(r, 3),
        분류     = S(r, 4),
        사이트   = S(r, 5),
        업체약칭 = S(r, 6),
        제목     = S(r, 7),
        내용     = S(r, 8),
        시작시간 = S(r, 9),
        종료시간 = S(r, 10),
        첨부파일 = S(r, 11),
        등록일시 = S(r, 12),
        등록자   = S(r, 13),
    };
}
