using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Data.Sqlite;
using ETA.Models;
using System.Diagnostics;

namespace ETA.Services;

public static class TestReportService
{
    private const string TableName = "분석의뢰및결과";

    private static string GetDatabasePath() => DbPathHelper.DbPath;

    private static readonly HashSet<string> FixedColumns =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "Id", "채취일자", "채취시간", "의뢰사업장", "약칭", "시료명",
            "견적번호", "입회자", "시료채취자-1", "시료채취자-2",
            "방류허용기준 적용유무", "정도보증유무", "분석완료일자", "견적구분"
        };

    public static List<string> GetAnalyteColumns()
    {
        var cols   = new List<string>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return cols;
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info(\"{TableName}\")";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var col = r.GetString(1);
            if (!FixedColumns.Contains(col)) cols.Add(col);
        }
        Debug.WriteLine($"[TestReport] 분석항목 {cols.Count}개");
        return cols;
    }

    public static List<string> GetCompanyList()
    {
        var list   = new List<string>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) { Debug.WriteLine($"[TestReport] DB없음"); return list; }
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var chk = conn.CreateCommand();
        chk.CommandText = $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{TableName}'";
        if (Convert.ToInt32(chk.ExecuteScalar()) == 0) { Debug.WriteLine($"[TestReport] 테이블없음"); return list; }
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT DISTINCT \"약칭\" FROM \"{TableName}\" WHERE \"약칭\" IS NOT NULL AND \"약칭\" <> '' ORDER BY \"약칭\" ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        Debug.WriteLine($"[TestReport] 업체 {list.Count}개");
        return list;
    }

    public static List<SampleRequest> GetSamplesByCompany(string 약칭)
    {
        var list   = new List<SampleRequest>();
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return list;
        var analyteCols = GetAnalyteColumns();
        var colSelect   = analyteCols.Count > 0 ? "," + string.Join(",", analyteCols.Select(c => $"\"{c}\"")) : "";
        using var conn  = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"SELECT rowid AS Id,
                ""채취일자"",""채취시간"",""의뢰사업장"",""약칭"",""시료명"",
                ""견적번호"",""입회자"",""시료채취자-1"",""시료채취자-2"",
                ""방류허용기준 적용유무"",""정도보증유무"",""분석완료일자"",""견적구분""
                {colSelect}
            FROM ""{TableName}"" WHERE ""약칭"" = @약칭 ORDER BY ""채취일자"" DESC, rowid DESC";
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
        Debug.WriteLine($"[TestReport] {약칭} → {list.Count}건");
        return list;
    }

    public static bool UpdateResult(int rowId, string columnName, string newValue)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"UPDATE \"{TableName}\" SET \"{columnName.Trim()}\" = @val WHERE rowid = @id";
        cmd.Parameters.AddWithValue("@val", string.IsNullOrEmpty(newValue) ? DBNull.Value : (object)newValue);
        cmd.Parameters.AddWithValue("@id", rowId);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[UPDATE] rowid={rowId} {columnName}={newValue} → {rows}행");
        return rows > 0;
    }

    /// <summary>여러 분析항목 결과값을 한 번에 업데이트. 반환값 = 성공한 항목 수</summary>
    public static int BulkUpdateResults(int rowId, Dictionary<string, string> analyteValues)
    {
        if (analyteValues.Count == 0) return 0;
        var dbPath = GetDatabasePath();
        int count = 0;
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var tx = conn.BeginTransaction();
        try
        {
            foreach (var kv in analyteValues)
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = $"UPDATE \"{TableName}\" SET \"{kv.Key.Trim()}\" = @val WHERE rowid = @id";
                cmd.Parameters.AddWithValue("@val", string.IsNullOrEmpty(kv.Value) ? DBNull.Value : (object)kv.Value);
                cmd.Parameters.AddWithValue("@id", rowId);
                if (cmd.ExecuteNonQuery() > 0) count++;
            }
            tx.Commit();
        }
        catch { tx.Rollback(); throw; }
        Debug.WriteLine($"[BulkUpdate] rowid={rowId} → {count}/{analyteValues.Count}항목 저장");
        return count;
    }

    public static bool DeleteSample(int rowId)
    {
        var dbPath = GetDatabasePath();
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"DELETE FROM \"{TableName}\" WHERE rowid = @id";
        cmd.Parameters.AddWithValue("@id", rowId);
        int rows = cmd.ExecuteNonQuery();
        Debug.WriteLine($"[DELETE] rowid={rowId} → {rows}행");
        return rows > 0;
    }

    /// <summary>특정 rowId의 단일 분析항목 결과값 조회. 없으면 null</summary>
    public static string? GetAnalyteValue(int rowId, string columnName)
    {
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return null;
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT \"{columnName.Trim()}\" FROM \"{TableName}\" WHERE rowid = @id LIMIT 1";
            cmd.Parameters.AddWithValue("@id", rowId);
            var result = cmd.ExecuteScalar();
            return result == null || result == DBNull.Value ? null : result.ToString();
        }
        catch (Exception ex) { Debug.WriteLine($"[GetAnalyteValue] 오류: {ex.Message}"); return null; }
    }

    /// <summary>견적번호+시료명 또는 약칭+시료명으로 rowid 조회. 없으면 null</summary>
    public static int? FindRowId(string 견적번호, string 약칭, string 시료명)
    {
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return null;
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (!string.IsNullOrEmpty(견적번호))
            {
                cmd.CommandText = $"SELECT rowid FROM \"{TableName}\" WHERE \"견적번호\" = @q AND \"시료명\" = @s LIMIT 1";
                cmd.Parameters.AddWithValue("@q", 견적번호);
                cmd.Parameters.AddWithValue("@s", 시료명);
            }
            else
            {
                cmd.CommandText = $"SELECT rowid FROM \"{TableName}\" WHERE \"약칭\" = @y AND \"시료명\" = @s LIMIT 1";
                cmd.Parameters.AddWithValue("@y", 약칭);
                cmd.Parameters.AddWithValue("@s", 시료명);
            }
            var result = cmd.ExecuteScalar();
            return result == null || result == DBNull.Value ? null : Convert.ToInt32(result);
        }
        catch (Exception ex) { Debug.WriteLine($"[FindRowId] 오류: {ex.Message}"); return null; }
    }

    public static Dictionary<string, AnalysisItem> GetAnalyteMeta()
    {
        var dict   = new Dictionary<string, AnalysisItem>(StringComparer.OrdinalIgnoreCase);
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return dict;
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var chk = conn.CreateCommand();
        chk.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='분석정보'";
        if (Convert.ToInt32(chk.ExecuteScalar()) == 0) return dict;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Analyte, unit, Method, instrument, ES, Category FROM \"분석정보\"";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var analyte = S(r, 0);
            if (!string.IsNullOrEmpty(analyte))
                dict[analyte] = new AnalysisItem { Analyte = analyte, unit = S(r,1), Method = S(r,2), instrument = S(r,3), ES = S(r,4), Category = S(r,5) };
        }
        Debug.WriteLine($"[TestReport] 메타 {dict.Count}개");
        return dict;
    }

    // 방류기준표: 구분(col0)=항목명(행), 헤더=방류허용기준 컬럼명(열)
    public static string GetStandardValue(string 항목명, string 방류허용기준컬럼)
    {
        if (string.IsNullOrEmpty(항목명) || string.IsNullOrEmpty(방류허용기준컬럼)) return "";
        var dbPath = GetDatabasePath();
        if (!File.Exists(dbPath)) return "";
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            using var chk = conn.CreateCommand();
            chk.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='방류기준표'";
            if (Convert.ToInt32(chk.ExecuteScalar()) == 0) { Debug.WriteLine("[방류기준] 테이블없음"); return ""; }

            var safeCol = 방류허용기준컬럼.Trim();
            using var colCmd = conn.CreateCommand();
            colCmd.CommandText = "PRAGMA table_info(\"방류기준표\")";
            var cols = new List<string>();
            using (var cr = colCmd.ExecuteReader()) while (cr.Read()) cols.Add(cr.GetString(1));

            // 정확히 일치 → 없으면 하이픈/공백 무시 후 유사 매칭
            string? matchCol = cols.FirstOrDefault(col => col.Trim() == safeCol);
            if (matchCol == null)
            {
                var norm = (string s) => s.Replace("-", "").Replace(" ", "").ToLower().Trim();
                matchCol = cols.FirstOrDefault(col => norm(col) == norm(safeCol));
                if (matchCol != null)
                    Debug.WriteLine($"[방류기준] 유사매칭: '{safeCol}' → '{matchCol}'");
            }
            if (matchCol == null)
            {
                Debug.WriteLine($"[방류기준] '{safeCol}' 컬럼없음. 전체: {string.Join(", ", cols.Take(5))}");
                return "";
            }
            safeCol = matchCol.Trim();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT \"{safeCol}\" FROM \"방류기준표\" WHERE TRIM(\"구분\") = @항목 LIMIT 1";
            cmd.Parameters.AddWithValue("@항목", 항목명.Trim());
            var result = cmd.ExecuteScalar();
            var val    = result == null || result == DBNull.Value ? "" : result.ToString()?.Trim() ?? "";
            Debug.WriteLine($"[방류기준] {항목명} / {safeCol} → '{val}'");
            return val == "해당없음" ? "" : val;
        }
        catch (Exception ex) { Debug.WriteLine($"[방류기준] 오류: {ex.Message}"); return ""; }
    }

    private static string S(SqliteDataReader r, int i)
    {
        try { return r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? ""; }
        catch { return ""; }
    }
}
