using System;
using System.Collections.Generic;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 폐수배출업소 분석결과 2-table 구조:
///   폐수_결과      — 헤더 (업체 x 채취일자)
///   폐수_결과_항목 — 항목별 결과값
///   결과_제출이력  — 측정인/ERP/Zero4 제출 로그
/// </summary>
public static class WasteAnalysisResultService
{
    // ── 헤더 조회/생성 ────────────────────────────────────────────────────
    public static WasteResultEntry? GetEntry(string 관리번호, string 채취일자)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT id, 의뢰_id, 업체명, 입력자, 입력일시 " +
            "FROM `폐수_결과` WHERE 관리번호=@mn AND 채취일자=@dt LIMIT 1";
        cmd.Parameters.AddWithValue("@mn", 관리번호);
        cmd.Parameters.AddWithValue("@dt", 채취일자);
        using var r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new WasteResultEntry
        {
            Id       = Convert.ToInt32(r["id"]),
            의뢰Id   = r["의뢰_id"] is DBNull ? null : Convert.ToInt32(r["의뢰_id"]),
            관리번호 = 관리번호,
            업체명   = r["업체명"]?.ToString() ?? "",
            채취일자 = 채취일자,
            입력자   = r["입력자"]?.ToString() ?? "",
            입력일시 = r["입력일시"]?.ToString() ?? "",
        };
    }

    /// <summary>헤더가 없으면 INSERT 후 반환</summary>
    public static WasteResultEntry EnsureEntry(string 관리번호, string 업체명, string 채취일자, int? 의뢰Id = null)
    {
        var existing = GetEntry(관리번호, 채취일자);
        if (existing != null) return existing;

        string now  = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string user = CurrentUserManager.Instance.CurrentUserId ?? "";

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "INSERT INTO `폐수_결과` (의뢰_id, 관리번호, 업체명, 채취일자, 입력자, 입력일시) " +
            "VALUES (@rid, @mn, @co, @dt, @user, @now)";
        cmd.Parameters.AddWithValue("@rid",  (object?)의뢰Id ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@mn",   관리번호);
        cmd.Parameters.AddWithValue("@co",   업체명);
        cmd.Parameters.AddWithValue("@dt",   채취일자);
        cmd.Parameters.AddWithValue("@user", user);
        cmd.Parameters.AddWithValue("@now",  now);
        cmd.ExecuteNonQuery();

        return GetEntry(관리번호, 채취일자)!;
    }

    // ── 항목 조회 ─────────────────────────────────────────────────────────
    public static List<WasteResultItem> GetItems(int 결과Id)
    {
        var list = new List<WasteResultItem>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT id, 항목, 결과값, 단위, 비고 " +
            "FROM `폐수_결과_항목` WHERE 결과_id=@rid ORDER BY id";
        cmd.Parameters.AddWithValue("@rid", 결과Id);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new WasteResultItem
            {
                Id     = Convert.ToInt32(r["id"]),
                결과Id = 결과Id,
                항목   = r["항목"]?.ToString() ?? "",
                결과값 = r["결과값"]?.ToString() ?? "",
                단위   = r["단위"]?.ToString() ?? "",
                비고   = r["비고"]?.ToString() ?? "",
            });
        return list;
    }

    /// <summary>헤더 + 항목 한꺼번에 조회</summary>
    public static WasteResultEntry? GetEntryWithItems(string 관리번호, string 채취일자)
    {
        var entry = GetEntry(관리번호, 채취일자);
        if (entry == null) return null;
        entry.항목들 = GetItems(entry.Id);
        return entry;
    }

    // ── 날짜별 헤더 목록 ──────────────────────────────────────────────────
    public static List<WasteResultEntry> GetEntriesByDate(string 채취일자)
    {
        var list = new List<WasteResultEntry>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT id, 의뢰_id, 관리번호, 업체명, 입력자, 입력일시 " +
            "FROM `폐수_결과` WHERE 채취일자=@dt ORDER BY 관리번호";
        cmd.Parameters.AddWithValue("@dt", 채취일자);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new WasteResultEntry
            {
                Id       = Convert.ToInt32(r["id"]),
                의뢰Id   = r["의뢰_id"] is DBNull ? null : Convert.ToInt32(r["의뢰_id"]),
                관리번호 = r["관리번호"]?.ToString() ?? "",
                업체명   = r["업체명"]?.ToString() ?? "",
                채취일자 = 채취일자,
                입력자   = r["입력자"]?.ToString() ?? "",
                입력일시 = r["입력일시"]?.ToString() ?? "",
            });
        return list;
    }

    // ── 항목 저장 (INSERT or UPDATE) ──────────────────────────────────────
    public static void SaveItems(int 결과Id, IEnumerable<WasteResultItem> items)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        foreach (var item in items)
        {
            using var cmd = conn.CreateCommand();
            if (item.Id == 0)
            {
                cmd.CommandText =
                    "INSERT INTO `폐수_결과_항목` (결과_id, 항목, 결과값, 단위, 비고) " +
                    "VALUES (@rid, @it, @val, @unit, @memo)";
                cmd.Parameters.AddWithValue("@rid",  결과Id);
                cmd.Parameters.AddWithValue("@it",   item.항목);
                cmd.Parameters.AddWithValue("@val",  item.결과값);
                cmd.Parameters.AddWithValue("@unit", item.단위);
                cmd.Parameters.AddWithValue("@memo", item.비고);
            }
            else
            {
                cmd.CommandText =
                    "UPDATE `폐수_결과_항목` SET 결과값=@val, 단위=@unit, 비고=@memo WHERE id=@id";
                cmd.Parameters.AddWithValue("@id",   item.Id);
                cmd.Parameters.AddWithValue("@val",  item.결과값);
                cmd.Parameters.AddWithValue("@unit", item.단위);
                cmd.Parameters.AddWithValue("@memo", item.비고);
            }
            cmd.ExecuteNonQuery();
        }
    }

    // ── 항목 전체 삭제 ────────────────────────────────────────────────────
    public static void DeleteItems(int 결과Id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `폐수_결과_항목` WHERE 결과_id=@rid";
        cmd.Parameters.AddWithValue("@rid", 결과Id);
        cmd.ExecuteNonQuery();
    }

    // ── 헤더 + 항목 전체 삭제 ────────────────────────────────────────────
    public static void DeleteEntry(int 결과Id)
    {
        DeleteItems(결과Id);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `폐수_결과` WHERE id=@id";
        cmd.Parameters.AddWithValue("@id", 결과Id);
        cmd.ExecuteNonQuery();
    }

    // ── 제출 이력 저장 ────────────────────────────────────────────────────
    public static void LogSubmit(ResultSubmitLog log)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `결과_제출이력`
                (제출유형, 대상유형, 대상명, 기간_시작, 기간_종료,
                 제출건수, 제출자, 제출일시, 파일경로, 상태, 비고)
            VALUES
                (@type, @target, @name, @from, @to,
                 @cnt, @user, @dt, @file, @status, @memo)";
        cmd.Parameters.AddWithValue("@type",   log.제출유형);
        cmd.Parameters.AddWithValue("@target", log.대상유형);
        cmd.Parameters.AddWithValue("@name",   log.대상명);
        cmd.Parameters.AddWithValue("@from",   log.기간시작);
        cmd.Parameters.AddWithValue("@to",     log.기간종료);
        cmd.Parameters.AddWithValue("@cnt",    log.제출건수);
        cmd.Parameters.AddWithValue("@user",   log.제출자);
        cmd.Parameters.AddWithValue("@dt",     log.제출일시);
        cmd.Parameters.AddWithValue("@file",   log.파일경로);
        cmd.Parameters.AddWithValue("@status", log.상태);
        cmd.Parameters.AddWithValue("@memo",   log.비고);
        cmd.ExecuteNonQuery();
    }

    public static List<ResultSubmitLog> GetSubmitLogs(string 제출유형 = "")
    {
        var list = new List<ResultSubmitLog>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        bool filter = !string.IsNullOrEmpty(제출유형);
        cmd.CommandText =
            "SELECT id, 제출유형, 대상유형, 대상명, 기간_시작, 기간_종료, " +
            "제출건수, 제출자, 제출일시, 파일경로, 상태, 비고 " +
            "FROM `결과_제출이력` " +
            (filter ? "WHERE 제출유형=@type " : "") +
            "ORDER BY id DESC";
        if (filter) cmd.Parameters.AddWithValue("@type", 제출유형);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new ResultSubmitLog
            {
                Id       = Convert.ToInt32(r["id"]),
                제출유형 = r["제출유형"]?.ToString() ?? "",
                대상유형 = r["대상유형"]?.ToString() ?? "",
                대상명   = r["대상명"]?.ToString() ?? "",
                기간시작 = r["기간_시작"]?.ToString() ?? "",
                기간종료 = r["기간_종료"]?.ToString() ?? "",
                제출건수 = r["제출건수"] is DBNull ? 0 : Convert.ToInt32(r["제출건수"]),
                제출자   = r["제출자"]?.ToString() ?? "",
                제출일시 = r["제출일시"]?.ToString() ?? "",
                파일경로 = r["파일경로"]?.ToString() ?? "",
                상태     = r["상태"]?.ToString() ?? "",
                비고     = r["비고"]?.ToString() ?? "",
            });
        return list;
    }
}
