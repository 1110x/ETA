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
            "방류허용기준 적용유무", "정도보증유무", "분석완료일자", "분석완료일자", "견적구분",
            // 시료유형 분기 (접수/채수)
            "시료유형", "접수일자", "접수담당자", "업체담당자",
        };

    /// <summary>시료유형 분기에 사용되는 4개 컬럼 — 자가치유 대상</summary>
    public static readonly string[] SampleTypeCols =
        new[] { "시료유형", "접수일자", "접수담당자", "업체담당자" };

    public static void EnsureSampleTypeColumns(System.Data.Common.DbConnection conn)
    {
        try
        {
            if (!DbConnectionFactory.TableExists(conn, TableName)) return;
            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, TableName),
                StringComparer.OrdinalIgnoreCase);
            foreach (var col in SampleTypeCols)
            {
                if (existing.Contains(col)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `{TableName}` ADD COLUMN `{col}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); existing.Add(col); }
                catch { /* 컬럼 충돌 등은 무시 */ }
            }
        }
        catch { }
    }

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
        using var conn  = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureSampleTypeColumns(conn);
        var analyteCols = GetAnalyteColumns();
        var colSelect   = analyteCols.Count > 0 ? "," + string.Join(",", analyteCols.Select(c => $"`{c}`")) : "";
        using var cmd = conn.CreateCommand();
        var rid = DbConnectionFactory.RowId;
        cmd.CommandText = $@"SELECT {rid} AS Id,
                `채취일자`,`채취시간`,`의뢰사업장`,`약칭`,`시료명`,
                `견적번호`,`입회자`,`시료채취자-1`,`시료채취자-2`,
                `방류허용기준 적용유무`,`정도보증유무`,`분석완료일자`,`견적구분`,
                `시료유형`,`접수일자`,`접수담당자`,`업체담당자`
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
                시료유형     = S(r, 14),
                접수일자     = S(r, 15),
                접수담당자   = S(r, 16),
                업체담당자   = S(r, 17),
            };
            int analyteOffset = 18;
            for (int i = 0; i < analyteCols.Count; i++)
            {
                int idx = analyteOffset + i;
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

    /// <summary>
    /// 특정 방류허용기준 코드(행)의 전체 항목값을 한번에 로드
    /// 새 구조: 행=방류기준코드, 열=분석항목
    /// </summary>
    public static Dictionary<string, string> GetStandardValueMap(string 방류허용기준코드)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(방류허용기준코드)) return map;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "방류기준표")) return map;

            // 구분 컬럼 매칭: 정확 일치 → 하이픈/공백 무시 매칭
            using var findCmd = conn.CreateCommand();
            findCmd.CommandText = "SELECT 구분 FROM `방류기준표`";
            var allCodes = new List<string>();
            using (var r0 = findCmd.ExecuteReader())
                while (r0.Read()) allCodes.Add(r0.GetValue(0)?.ToString()?.Trim() ?? "");

            var target = 방류허용기준코드.Trim();
            string? matchCode = allCodes.FirstOrDefault(c => c == target);
            if (matchCode == null)
            {
                var norm = (string s) => s.Replace("-", "").Replace(" ", "").ToLower().Trim();
                matchCode = allCodes.FirstOrDefault(c => norm(c) == norm(target));
            }
            if (matchCode == null) return map;

            // 해당 행의 모든 컬럼(분석항목) 조회
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `방류기준표` WHERE TRIM(`구분`) = @code LIMIT 1";
            cmd.Parameters.AddWithValue("@code", matchCode);
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                for (int i = 0; i < r.FieldCount; i++)
                {
                    var col = r.GetName(i);
                    if (col == "_id" || col == "구분") continue;
                    var val = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (val == "해당없음") val = "";
                    map[col] = val;
                }
            }
        }
        catch { }
        return map;
    }

    // 방류기준표: 행=방류기준코드, 열=분석항목
    public static string GetStandardValue(string 항목명, string 방류허용기준코드)
    {
        if (string.IsNullOrEmpty(항목명) || string.IsNullOrEmpty(방류허용기준코드)) return "";
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "방류기준표")) { return ""; }

            var safeCol = 항목명.Trim();
            var cols = DbConnectionFactory.GetColumnNames(conn, "방류기준표");

            // 분석항목 컬럼 매칭
            string? matchCol = cols.FirstOrDefault(col => col.Trim() == safeCol);
            if (matchCol == null)
            {
                var norm = (string s) => s.Replace("-", "").Replace(" ", "").ToLower().Trim();
                matchCol = cols.FirstOrDefault(col => norm(col) == norm(safeCol));
            }
            if (matchCol == null) return "";
            safeCol = matchCol.Trim();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `{safeCol}` FROM `방류기준표` WHERE TRIM(`구분`) = @code LIMIT 1";
            cmd.Parameters.AddWithValue("@code", 방류허용기준코드.Trim());
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
