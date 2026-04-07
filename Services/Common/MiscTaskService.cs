using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using ETA.Models;

namespace ETA.Services.Common;

public static class MiscTaskService
{
    private static void EnsureTable()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            CREATE TABLE IF NOT EXISTS `기타업무` (
                Id        INTEGER PRIMARY KEY AUTO_INCREMENT,
                업무명    TEXT    NOT NULL DEFAULT '',
                내용      TEXT    DEFAULT '',
                배정자    TEXT    DEFAULT '',
                담당자id  TEXT    DEFAULT '',
                담당자명  TEXT    DEFAULT '',
                상태      TEXT    DEFAULT '대기',
                마감일    TEXT    DEFAULT '',
                등록일시  TEXT    DEFAULT '',
                완료일시  TEXT    DEFAULT ''
            )";
        try { cmd.ExecuteNonQuery(); } catch { }
    }

    public static List<MiscTask> GetAll()
    {
        EnsureTable();
        var list = new List<MiscTask>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM `기타업무` ORDER BY Id DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Read(r));
        return list;
    }

    public static List<MiscTask> GetByAgent(string 담당자id)
    {
        EnsureTable();
        var list = new List<MiscTask>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM `기타업무` WHERE 담당자id=@id ORDER BY Id DESC";
        cmd.Parameters.AddWithValue("@id", 담당자id);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Read(r));
        return list;
    }

    public static void Insert(MiscTask t)
    {
        EnsureTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `기타업무` (업무명, 내용, 배정자, 담당자id, 담당자명, 상태, 마감일, 등록일시)
            VALUES (@업무명, @내용, @배정자, @담당자id, @담당자명, @상태, @마감일, @등록일시)";
        SetParams(cmd, t);
        cmd.ExecuteNonQuery();
    }

    public static void Update(MiscTask t)
    {
        EnsureTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `기타업무` SET
                업무명=@업무명, 내용=@내용, 배정자=@배정자,
                담당자id=@담당자id, 담당자명=@담당자명,
                상태=@상태, 마감일=@마감일, 완료일시=@완료일시
            WHERE Id=@Id";
        SetParams(cmd, t);
        cmd.Parameters.AddWithValue("@완료일시", t.완료일시 ?? "");
        cmd.Parameters.AddWithValue("@Id", t.Id);
        cmd.ExecuteNonQuery();
    }

    public static void Delete(int id)
    {
        EnsureTable();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `기타업무` WHERE Id=@Id";
        cmd.Parameters.AddWithValue("@Id", id);
        cmd.ExecuteNonQuery();
    }

    private static void SetParams(DbCommand cmd, MiscTask t)
    {
        cmd.Parameters.AddWithValue("@업무명",   t.업무명 ?? "");
        cmd.Parameters.AddWithValue("@내용",     t.내용 ?? "");
        cmd.Parameters.AddWithValue("@배정자",   t.배정자 ?? "");
        cmd.Parameters.AddWithValue("@담당자id", t.담당자id ?? "");
        cmd.Parameters.AddWithValue("@담당자명", t.담당자명 ?? "");
        cmd.Parameters.AddWithValue("@상태",     t.상태 ?? "대기");
        cmd.Parameters.AddWithValue("@마감일",   t.마감일 ?? "");
        cmd.Parameters.AddWithValue("@등록일시", t.등록일시 ?? "");
    }

    private static string S(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) : ""; }
        catch { return ""; }
    }

    private static MiscTask Read(DbDataReader r) => new()
    {
        Id       = r.GetInt32(r.GetOrdinal("Id")),
        업무명   = S(r, "업무명"),
        내용     = S(r, "내용"),
        배정자   = S(r, "배정자"),
        담당자id = S(r, "담당자id"),
        담당자명 = S(r, "담당자명"),
        상태     = S(r, "상태"),
        마감일   = S(r, "마감일"),
        등록일시 = S(r, "등록일시"),
        완료일시 = S(r, "완료일시"),
    };
}
