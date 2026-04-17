using System;
using System.Collections.Generic;
using System.Data.Common;
using ETA.Models;

namespace ETA.Services.Common;

public static class ReagentService
{
    private static void EnsureTable(DbConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            CREATE TABLE IF NOT EXISTS `시약` (
                Id           INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                ITEM_NO      TEXT DEFAULT '',
                품목명       TEXT NOT NULL DEFAULT '',
                영문명       TEXT DEFAULT '',
                CAS번호      TEXT DEFAULT '',
                화학식       TEXT DEFAULT '',
                규격         TEXT DEFAULT '',
                단위         TEXT DEFAULT '',
                제조사       TEXT DEFAULT '',
                위험등급     TEXT DEFAULT '일반',
                GHS          TEXT DEFAULT '',
                보관조건     TEXT DEFAULT '',
                재고량       INTEGER DEFAULT 0,
                당월사용량    INTEGER DEFAULT 0,
                전월사용량    INTEGER DEFAULT 0,
                적정사용량    INTEGER DEFAULT 0,
                최대적정보유량 INTEGER DEFAULT 0,
                만료일       TEXT DEFAULT '',
                비고         TEXT DEFAULT '',
                등록일       TEXT DEFAULT '',
                상태         TEXT DEFAULT '정상'
            )";
        cmd.ExecuteNonQuery();

        // 기존 테이블 마이그레이션 — 없는 컬럼만 추가
        foreach (var col in new[]
        {
            "ITEM_NO TEXT DEFAULT ''",
            "영문명 TEXT DEFAULT ''",
            "화학식 TEXT DEFAULT ''",
            "단위 TEXT DEFAULT ''",
            "GHS TEXT DEFAULT ''",
            "재고량 INTEGER DEFAULT 0",
            "당월사용량 INTEGER DEFAULT 0",
            "전월사용량 INTEGER DEFAULT 0",
            "적정사용량 INTEGER DEFAULT 0",
            "최대적정보유량 INTEGER DEFAULT 0",
        })
        {
            try
            {
                using var alter = conn.CreateCommand();
                alter.CommandText = $"ALTER TABLE `시약` ADD COLUMN {col}";
                alter.ExecuteNonQuery();
            }
            catch { /* 이미 존재 */ }
        }
    }

    public static List<Reagent> GetAll() => GetAllInternal(includeDeleted: false);
    public static List<Reagent> GetAllDeleted() => GetAllInternal(deletedOnly: true);

    private static List<Reagent> GetAllInternal(bool includeDeleted = false, bool deletedOnly = false)
    {
        var list = new List<Reagent>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        var where = deletedOnly ? "WHERE `삭제여부`=1"
                  : includeDeleted ? ""
                  : "WHERE (`삭제여부` IS NULL OR `삭제여부`=0)";
        cmd.CommandText = $@"
            SELECT Id,ITEM_NO,품목명,영문명,CAS번호,화학식,규격,단위,제조사,위험등급,
                   GHS,보관조건,재고량,당월사용량,전월사용량,적정사용량,최대적정보유량,
                   만료일,비고,등록일,상태
            FROM `시약` {where} ORDER BY 품목명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Map(r));
        return list;
    }

    public static int Count()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM `시약`";
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    public static bool Insert(Reagent item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `시약`
                (ITEM_NO,품목명,영문명,CAS번호,화학식,규격,단위,제조사,위험등급,
                 GHS,보관조건,재고량,당월사용량,전월사용량,적정사용량,최대적정보유량,
                 만료일,비고,등록일,상태)
            VALUES
                (@ITEM_NO,@품목명,@영문명,@CAS번호,@화학식,@규격,@단위,@제조사,@위험등급,
                 @GHS,@보관조건,@재고량,@당월사용량,@전월사용량,@적정사용량,@최대적정보유량,
                 @만료일,@비고,@등록일,@상태)";
        SetParams(cmd, item);
        int rows = cmd.ExecuteNonQuery();
        if (rows > 0)
        {
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
            item.Id = Convert.ToInt32(idCmd.ExecuteScalar());
        }
        return rows > 0;
    }

    public static bool Update(Reagent item)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `시약` SET
                ITEM_NO=@ITEM_NO, 품목명=@품목명, 영문명=@영문명, CAS번호=@CAS번호,
                화학식=@화학식, 규격=@규격, 단위=@단위, 제조사=@제조사, 위험등급=@위험등급,
                GHS=@GHS, 보관조건=@보관조건, 재고량=@재고량, 당월사용량=@당월사용량,
                전월사용량=@전월사용량, 적정사용량=@적정사용량, 최대적정보유량=@최대적정보유량,
                만료일=@만료일, 비고=@비고, 등록일=@등록일, 상태=@상태
            WHERE Id=@id";
        SetParams(cmd, item);
        cmd.Parameters.AddWithValue("@id", item.Id);
        return cmd.ExecuteNonQuery() > 0;
    }

    /// <summary>소프트 삭제 (삭제여부=1로 UPDATE)</summary>
    public static bool Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `시약` SET `삭제여부`=1 WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    /// <summary>휴지통에서 복원 (삭제여부=0)</summary>
    public static bool Restore(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `시약` SET `삭제여부`=0 WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    private static void SetParams(DbCommand cmd, Reagent item)
    {
        cmd.Parameters.AddWithValue("@ITEM_NO",      item.ITEM_NO        ?? "");
        cmd.Parameters.AddWithValue("@품목명",       item.품목명         ?? "");
        cmd.Parameters.AddWithValue("@영문명",       item.영문명         ?? "");
        cmd.Parameters.AddWithValue("@CAS번호",      item.CAS번호        ?? "");
        cmd.Parameters.AddWithValue("@화학식",       item.화학식         ?? "");
        cmd.Parameters.AddWithValue("@규격",         item.규격           ?? "");
        cmd.Parameters.AddWithValue("@단위",         item.단위           ?? "");
        cmd.Parameters.AddWithValue("@제조사",       item.제조사         ?? "");
        cmd.Parameters.AddWithValue("@위험등급",     item.위험등급       ?? "일반");
        cmd.Parameters.AddWithValue("@GHS",          item.GHS            ?? "");
        cmd.Parameters.AddWithValue("@보관조건",     item.보관조건       ?? "");
        cmd.Parameters.AddWithValue("@재고량",       item.재고량);
        cmd.Parameters.AddWithValue("@당월사용량",   item.당월사용량);
        cmd.Parameters.AddWithValue("@전월사용량",   item.전월사용량);
        cmd.Parameters.AddWithValue("@적정사용량",   item.적정사용량);
        cmd.Parameters.AddWithValue("@최대적정보유량", item.최대적정보유량);
        cmd.Parameters.AddWithValue("@만료일",       item.만료일         ?? "");
        cmd.Parameters.AddWithValue("@비고",         item.비고           ?? "");
        cmd.Parameters.AddWithValue("@등록일",       item.등록일         ?? "");
        cmd.Parameters.AddWithValue("@상태",         item.상태           ?? "정상");
    }

    private static Reagent Map(DbDataReader r) => new()
    {
        Id           = r.IsDBNull(0)  ? 0  : r.GetInt32(0),
        ITEM_NO      = S(r, 1),
        품목명       = S(r, 2),
        영문명       = S(r, 3),
        CAS번호      = S(r, 4),
        화학식       = S(r, 5),
        규격         = S(r, 6),
        단위         = S(r, 7),
        제조사       = S(r, 8),
        위험등급     = S(r, 9),
        GHS          = S(r, 10),
        보관조건     = S(r, 11),
        재고량       = N(r, 12),
        당월사용량    = N(r, 13),
        전월사용량    = N(r, 14),
        적정사용량    = N(r, 15),
        최대적정보유량 = N(r, 16),
        만료일       = S(r, 17),
        비고         = S(r, 18),
        등록일       = S(r, 19),
        상태         = S(r, 20),
    };

    private static string S(DbDataReader r, int i) =>
        r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
    private static int N(DbDataReader r, int i) =>
        r.IsDBNull(i) ? 0 : Convert.ToInt32(r.GetValue(i));
}
