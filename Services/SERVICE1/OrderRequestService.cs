using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class OrderRequestService
{

    // ── 시료명칭 테이블 컬럼헤더(업체명) 전체 조회 ───────────────────────
    public static List<string> GetSampleNameColumns()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            list.AddRange(DbConnectionFactory.GetColumnNames(conn, "시료명칭"));
        }
        catch (Exception ex) { }
        return list;
    }

    // ── 업체명으로 시료명칭 컬럼 매칭 (정확→유사 순) ───────────────────────
    public static string? FindColumnByCompany(string companyName)
    {
        var cols = GetSampleNameColumns();
        if (cols.Count == 0) return null;

        // 1순위: 정확히 일치
        var exact = cols.FirstOrDefault(c =>
            string.Equals(c.Trim(), companyName.Trim(), StringComparison.OrdinalIgnoreCase));
        if (exact != null) return exact;

        // 2순위: 정규화 후 유사매칭
        static string Norm(string s) => s
            .Replace("㈜", "").Replace("(주)", "").Replace("(㈜)", "")
            .Replace(" ", "").Replace("　", "").ToLower().Trim();

        var normTarget = Norm(companyName);
        var best = cols
            .Select(c => (col: c, score: LCS(Norm(c), normTarget)))
            .Where(x => x.score > 2)
            .OrderByDescending(x => x.score)
            .FirstOrDefault();

        return best.col;
    }

    // ── 시료명칭 컬럼에서 NULL 아닌 값 조회 ──────────────────────────────
    public static List<string> GetSampleNames(string columnName)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT `{columnName}` FROM `시료명칭` WHERE `{columnName}` IS NOT NULL AND `{columnName}` <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
        }
        catch (Exception ex) { }
        return list;
    }

    // ── 시료명칭 테이블에 업체 컬럼 생성 ─────────────────────────────────
    public static bool CreateCompanyColumn(string companyName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            // 이미 있는지 확인
            var existing = DbConnectionFactory.GetColumnNames(conn, "시료명칭");
            if (existing.Any(c => string.Equals(c.Trim(), companyName.Trim(), StringComparison.OrdinalIgnoreCase)))
                return true;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"ALTER TABLE `시료명칭` ADD COLUMN `{companyName}` TEXT DEFAULT NULL";
            cmd.ExecuteNonQuery();
            return true;
        }
        catch (Exception ex) { return false; }
    }

    // ── 시료명칭 추가 ────────────────────────────────────────────────────
    public static bool AddSampleName(string columnName, string sampleName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = $@"SELECT {DbConnectionFactory.RowId} FROM `시료명칭` WHERE `{columnName}` IS NULL OR `{columnName}` = '' LIMIT 1";
            var rowidObj = chk.ExecuteScalar();

            if (rowidObj != null)
            {
                using var upd = conn.CreateCommand();
                upd.CommandText = $@"UPDATE `시료명칭` SET `{columnName}` = @val WHERE {DbConnectionFactory.RowId} = @id";
                upd.Parameters.AddWithValue("@val", sampleName);
                upd.Parameters.AddWithValue("@id", Convert.ToInt32(rowidObj));
                upd.ExecuteNonQuery();
            }
            else
            {
                using var ins = conn.CreateCommand();
                ins.CommandText = $@"INSERT INTO `시료명칭` (`{columnName}`) VALUES (@val)";
                ins.Parameters.AddWithValue("@val", sampleName);
                ins.ExecuteNonQuery();
            }
            return true;
        }
        catch (Exception ex) { return false; }
    }

    // ── 수질분석센터_결과 테이블 분석항목 컬럼 목록 ─────────────────────────
    public static List<string> GetAnalysisColumns()
    {
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "_id","id","rowid",
            "접수번호","의뢰일","업체명","대표자","담당자","연락처","이메일","비고",
            "채취일자","채취시간","의뢰사업장","약칭","시료명","견적번호",
            "입회자","시료채취자-1","시료채취자-2","방류허용기준 적용유무",
            "정도보증유무","분석완료일자","견적구분"
        };
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            foreach (var col in DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과"))
            {
                if (!fixedCols.Contains(col.Trim())) list.Add(col.Trim());
            }
        }
        catch (Exception ex) { }
        return list;
    }

    // ── 중복 확인 (견적번호 + 시료명) ────────────────────────────────────
    public static bool CheckDuplicate(string 견적번호, string 시료명)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT COUNT(*) FROM `수질분석센터_결과` WHERE `견적번호`=@no AND `시료명`=@sample";
            cmd.Parameters.AddWithValue("@no",     견적번호);
            cmd.Parameters.AddWithValue("@sample", 시료명);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
        catch { return false; }
    }

    // ── 기존 의뢰 삭제 (덮어쓰기용) ─────────────────────────────────────
    public static void DeleteByKey(string 견적번호, string 시료명)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"DELETE FROM `수질분석센터_결과` WHERE `견적번호`=@no AND `시료명`=@sample";
            cmd.Parameters.AddWithValue("@no",     견적번호);
            cmd.Parameters.AddWithValue("@sample", 시료명);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { }
    }

    // ── 수질분석센터_결과 테이블에 의뢰서 INSERT ────────────────────────────
    public static bool InsertOrderRequest(
        string sampleName,
        QuotationIssue issue,
        HashSet<string> checkedItems)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            var tableCols = DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과");

            var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id","id","rowid",
                "접수번호","의뢰일","업체명","대표자","담당자","연락처","이메일","비고",
                "채취일자","채취시간","의뢰사업장","약칭","시료명","견적번호",
                "입회자","시료채취자-1","시료채취자-2","방류허용기준 적용유무",
                "정도보증유무","분석완료일자","견적구분"
            };

            var analysisCols = tableCols.Where(c => !fixedCols.Contains(c)).ToList();

            var colList   = new List<string> { "`의뢰사업장`","`약칭`","`시료명`","`견적번호`","`견적구분`","`채취일자`" };
            var paramList = new List<string> { "@company","@abbr","@sample","@no","@type","@date" };

            foreach (var col in analysisCols)
            {
                colList.Add($"`{col}`");
                paramList.Add($"@a_{ToParam(col)}");
            }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"INSERT INTO `수질분석센터_결과` ({string.Join(",", colList)}) VALUES ({string.Join(",", paramList)})";

            cmd.Parameters.AddWithValue("@company", issue.업체명   ?? "");
            cmd.Parameters.AddWithValue("@abbr",    issue.약칭     ?? "");
            cmd.Parameters.AddWithValue("@sample",  sampleName);
            cmd.Parameters.AddWithValue("@no",      issue.견적번호 ?? "");
            cmd.Parameters.AddWithValue("@type",    issue.견적구분 ?? "");
            cmd.Parameters.AddWithValue("@date",    DateTime.Today.ToString("yyyy-MM-dd"));

            foreach (var col in analysisCols)
            {
                object val = checkedItems.Contains(col) ? (object)"O" : DBNull.Value;
                cmd.Parameters.AddWithValue($"@a_{ToParam(col)}", val);
            }

            int rows = cmd.ExecuteNonQuery();
            return rows > 0;
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static string ToParam(string name)
        => Regex.Replace(name, @"[^a-zA-Z0-9가-힣]", "_");

    private static int LCS(string a, string b)
    {
        int max = 0;
        for (int i = 0; i < a.Length; i++)
            for (int j = i + 1; j <= a.Length; j++)
                if (b.Contains(a[i..j])) max = Math.Max(max, j - i);
        return max;
    }
}
