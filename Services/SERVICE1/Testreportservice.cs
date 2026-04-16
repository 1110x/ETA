using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class TestReportService
{
    private const string TableName = "수질분석센터_결과";


    private static readonly HashSet<string> FixedColumns =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "_id", "Id", "채취일자", "채취시간", "의뢰사업장", "약칭", "시료명",
            "견적번호", "입회자", "시료채취자-1", "시료채취자-2",
            "방류허용기준 적용유무", "정도보증유무", "분석완료일자", "분석완료일자", "견적구분"
        };

    public static List<string> GetAnalyteColumns()
    {
        var cols   = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        foreach (var col in DbConnectionFactory.GetColumnNames(conn, TableName))
        {
            if (!FixedColumns.Contains(col)) cols.Add(col);
        }
        return cols;
    }

    public static List<string> GetCompanyList()
    {
        var list   = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!DbConnectionFactory.TableExists(conn, TableName)) { return list; }
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT DISTINCT `약칭` FROM `{TableName}` WHERE `약칭` IS NOT NULL AND `약칭` <> '' ORDER BY `약칭` ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public static List<SampleRequest> GetSamplesByCompany(string 약칭)
    {
        var list   = new List<SampleRequest>();
        var analyteCols = GetAnalyteColumns();
        var colSelect   = analyteCols.Count > 0 ? "," + string.Join(",", analyteCols.Select(c => $"`{c}`")) : "";
        using var conn  = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        var rid = DbConnectionFactory.RowId;
        cmd.CommandText = $@"SELECT {rid} AS Id,
                `채취일자`,`채취시간`,`의뢰사업장`,`약칭`,`시료명`,
                `견적번호`,`입회자`,`시료채취자-1`,`시료채취자-2`,
                `방류허용기준 적용유무`,`정도보증유무`,`분석완료일자`,`견적구분`
                {colSelect}
            FROM `{TableName}` WHERE `약칭` = @약칭 ORDER BY `채취일자` DESC, {rid} DESC";
        cmd.Parameters.AddWithValue("@약칭", 약칭);
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var s = new SampleRequest
            {
                Id           = r.IsDBNull(0) ? 0 : Convert.ToInt32(r.GetValue(0)),
                채취일자     = S(r, 1),
                채취시간     = S(r, 2),
                의뢰사업장   = S(r, 3),
                약칭         = S(r, 4),
                시료명       = S(r, 5),
                견적번호     = S(r, 6),
                입회자       = S(r, 7),
                시료채취자1  = S(r, 8),
                시료채취자2  = S(r, 9),
                방류허용기준 = S(r, 10),
                정도보증     = S(r, 11),
                분석종료일 = S(r, 12),
                견적구분     = S(r, 13),
            };
            for (int i = 0; i < analyteCols.Count; i++)
            {
                int idx = 14 + i;
                if (!r.IsDBNull(idx))
                {
                    var val = r.GetValue(idx)?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(val)) s.분석결과[analyteCols[i]] = val;
                }
            }
            list.Add(s);
        }
        return list;
    }

    public static bool UpdateResult(int rowId, string columnName, string newValue)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"UPDATE `{TableName}` SET `{columnName.Trim()}` = @val WHERE {DbConnectionFactory.RowId} = @id";
        cmd.Parameters.AddWithValue("@val", string.IsNullOrEmpty(newValue) ? DBNull.Value : (object)newValue);
        cmd.Parameters.AddWithValue("@id", rowId);
        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
    }

    /// <summary>여러 분석항목 결과값을 한 번에 업데이트. 반환값 = 성공한 항목 수</summary>
    public static int BulkUpdateResults(int rowId, Dictionary<string, string> analyteValues)
    {
        if (analyteValues.Count == 0) return 0;
        int count = 0;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var tx = conn.BeginTransaction();
        try
        {
            foreach (var kv in analyteValues)
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = $"UPDATE `{TableName}` SET `{kv.Key.Trim()}` = @val WHERE {DbConnectionFactory.RowId} = @id";
                cmd.Parameters.AddWithValue("@val", string.IsNullOrEmpty(kv.Value) ? DBNull.Value : (object)kv.Value);
                cmd.Parameters.AddWithValue("@id", rowId);
                if (cmd.ExecuteNonQuery() > 0) count++;
            }
            tx.Commit();
        }
        catch { tx.Rollback(); throw; }
        return count;
    }

    public static bool DeleteSample(int rowId)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"DELETE FROM `{TableName}` WHERE {DbConnectionFactory.RowId} = @id";
        cmd.Parameters.AddWithValue("@id", rowId);
        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
    }

    /// <summary>특정 rowId의 단일 분석항목 결과값 조회. 없으면 null</summary>
    public static string? GetAnalyteValue(int rowId, string columnName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `{columnName.Trim()}` FROM `{TableName}` WHERE {DbConnectionFactory.RowId} = @id LIMIT 1";
            cmd.Parameters.AddWithValue("@id", rowId);
            var result = cmd.ExecuteScalar();
            return result == null || result == DBNull.Value ? null : result.ToString();
        }
        catch (Exception ex) { return null; }
    }

    /// <summary>견적번호+시료명 또는 약칭+시료명으로 rowid 조회. 없으면 null</summary>
    public static int? FindRowId(string 견적번호, string 약칭, string 시료명)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (!string.IsNullOrEmpty(견적번호))
            {
                cmd.CommandText = $"SELECT {DbConnectionFactory.RowId} FROM `{TableName}` WHERE `견적번호` = @q AND `시료명` = @s LIMIT 1";
                cmd.Parameters.AddWithValue("@q", 견적번호);
                cmd.Parameters.AddWithValue("@s", 시료명);
            }
            else
            {
                cmd.CommandText = $"SELECT {DbConnectionFactory.RowId} FROM `{TableName}` WHERE `약칭` = @y AND `시료명` = @s LIMIT 1";
                cmd.Parameters.AddWithValue("@y", 약칭);
                cmd.Parameters.AddWithValue("@s", 시료명);
            }
            var result = cmd.ExecuteScalar();
            return result == null || result == DBNull.Value ? null : Convert.ToInt32(result);
        }
        catch (Exception ex) { return null; }
    }

    public static Dictionary<string, AnalysisItem> GetAnalyteMeta()
    {
        var dict   = new Dictionary<string, AnalysisItem>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        if (!DbConnectionFactory.TableExists(conn, "분석정보")) return dict;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Analyte, unit, Method, instrument, ES, Category FROM `분석정보`";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var analyte = S(r, 0);
            if (!string.IsNullOrEmpty(analyte))
                dict[analyte] = new AnalysisItem { Analyte = analyte, unit = S(r,1), Method = S(r,2), instrument = S(r,3), ES = S(r,4), Category = S(r,5) };
        }
        return dict;
    }

    // 방류기준표: 구분(col0)=항목명(행), 헤더=방류허용기준 컬럼명(열)
    public static string GetStandardValue(string 항목명, string 방류허용기준컬럼)
    {
        if (string.IsNullOrEmpty(항목명) || string.IsNullOrEmpty(방류허용기준컬럼)) return "";
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "방류기준표")) { return ""; }

            var safeCol = 방류허용기준컬럼.Trim();
            var cols = DbConnectionFactory.GetColumnNames(conn, "방류기준표");

            // 정확히 일치 → 없으면 하이픈/공백 무시 후 유사 매칭
            string? matchCol = cols.FirstOrDefault(col => col.Trim() == safeCol);
            if (matchCol == null)
            {
                var norm = (string s) => s.Replace("-", "").Replace(" ", "").ToLower().Trim();
                matchCol = cols.FirstOrDefault(col => norm(col) == norm(safeCol));
            }
            if (matchCol == null)
            {
                return "";
            }
            safeCol = matchCol.Trim();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `{safeCol}` FROM `방류기준표` WHERE TRIM(`구분`) = @항목 LIMIT 1";
            cmd.Parameters.AddWithValue("@항목", 항목명.Trim());
            var result = cmd.ExecuteScalar();
            var val    = result == null || result == DBNull.Value ? "" : result.ToString()?.Trim() ?? "";
            return val == "해당없음" ? "" : val;
        }
        catch (Exception ex) { return ""; }
    }

    private static string S(DbDataReader r, int i)
    {
        try { return r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? ""; }
        catch { return ""; }
    }
}
