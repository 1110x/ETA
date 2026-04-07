using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class AttachmentService
{
    private const string TableName = "시험성적서_첨부";

    // ── 테이블 보장 ──────────────────────────────────────────────────────────
    public static void EnsureTable()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);
        }
        catch (Exception ex) { Debug.WriteLine($"[Attachment] EnsureTable 오류: {ex.Message}"); }
    }

    private static void EnsureTableOnConn(DbConnection conn)
    {
        if (DbConnectionFactory.TableExists(conn, TableName)) return;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS `{TableName}` (
                    `_id`       INT          NOT NULL AUTO_INCREMENT,
                    `sample_id` INT          NOT NULL,
                    `원본파일명` VARCHAR(512) NOT NULL,
                    `저장경로`  VARCHAR(1024) NOT NULL,
                    `등록일시`  DATETIME     DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (`_id`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
        cmd.ExecuteNonQuery();
        Debug.WriteLine($"[Attachment] 테이블 생성: {TableName}");
    }

    // ── 조회 ─────────────────────────────────────────────────────────────────
    public static List<SampleAttachment> GetAttachments(int sampleId)
    {
        var list = new List<SampleAttachment>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, TableName)) return list;
            using var cmd = conn.CreateCommand();
            var rid = DbConnectionFactory.RowId;
            cmd.CommandText =
                $"SELECT {rid}, `sample_id`, `원본파일명`, `저장경로`, `등록일시` " +
                $"FROM `{TableName}` WHERE `sample_id` = @sid ORDER BY {rid} ASC";
            cmd.Parameters.AddWithValue("@sid", sampleId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new SampleAttachment
                {
                    Id       = Convert.ToInt32(r.GetValue(0)),
                    SampleId = Convert.ToInt32(r.GetValue(1)),
                    원본파일명 = r.IsDBNull(2) ? "" : r.GetString(2),
                    저장경로  = r.IsDBNull(3) ? "" : r.GetString(3),
                    등록일시  = r.IsDBNull(4) ? "" : r.GetValue(4)?.ToString() ?? "",
                });
        }
        catch (Exception ex) { Debug.WriteLine($"[Attachment] GetAttachments 오류: {ex.Message}"); }
        return list;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────────
    public static int AddAttachment(int sampleId, string originalFileName, string savedPath)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTableOnConn(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"INSERT INTO `{TableName}` (`sample_id`, `원본파일명`, `저장경로`) " +
                $"VALUES (@sid, @fn, @sp)";
            cmd.Parameters.AddWithValue("@sid", sampleId);
            cmd.Parameters.AddWithValue("@fn", originalFileName);
            cmd.Parameters.AddWithValue("@sp", savedPath);
            cmd.ExecuteNonQuery();
            using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            return Convert.ToInt32(cmd2.ExecuteScalar());
        }
        catch (Exception ex) { Debug.WriteLine($"[Attachment] AddAttachment 오류: {ex.Message}"); return -1; }
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────────
    public static bool DeleteAttachment(int id, string savedPath)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"DELETE FROM `{TableName}` WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
            if (File.Exists(savedPath))
                File.Delete(savedPath);
            return true;
        }
        catch (Exception ex) { Debug.WriteLine($"[Attachment] DeleteAttachment 오류: {ex.Message}"); return false; }
    }
}
