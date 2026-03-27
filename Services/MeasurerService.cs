using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ETA.Services;

public static class MeasurerService
{
    private static string GetDatabasePath() => DbPathHelper.DbPath;

    // ── 테이블 생성 + 컬럼 마이그레이션 ─────────────────────────────────────
    public static void EnsureMeasurerTable()
    {
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();

        // 기본 테이블 생성
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS 측정인_채취지점 (
                    Id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    계약번호   TEXT,
                    약칭       TEXT,
                    계약기간   TEXT,
                    업체명     TEXT NOT NULL,
                    채취지점명 TEXT NOT NULL,
                    UNIQUE(업체명, 채취지점명)
                );";
            cmd.ExecuteNonQuery();
        }

        // 기존 DB 마이그레이션 — 컬럼이 없으면 추가
        foreach (var col in new[] { ("계약번호", "TEXT"), ("약칭", "TEXT"), ("계약기간", "TEXT"),
                                    ("계약금액", "TEXT"), ("견적구분", "TEXT") })
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE 측정인_채취지점 ADD COLUMN {col.Item1} {col.Item2}";
                cmd.ExecuteNonQuery();
            }
            catch { /* 이미 존재하면 무시 */ }
        }
    }

    // ── 전체 데이터 초기화 (재스크래핑 전 호출) ──────────────────────────────
    public static void ClearAll()
    {
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM 측정인_채취지점";
        cmd.ExecuteNonQuery();
    }

    // ── 채취지점 데이터 저장 ─────────────────────────────────────────────────
    public static void SaveSamplingPoints(
        string companyName,
        List<string> points,
        string 계약번호 = "",
        string 계약기간 = "")
    {
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();

        using var txn = conn.BeginTransaction();
        try
        {
            foreach (var point in points.Where(p => !string.IsNullOrWhiteSpace(p)))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    INSERT INTO 측정인_채취지점 (계약번호, 계약기간, 업체명, 채취지점명)
                    VALUES (@contractNo, @period, @company, @point)
                    ON CONFLICT(업체명, 채취지점명) DO UPDATE SET
                        계약번호 = excluded.계약번호,
                        계약기간 = excluded.계약기간;";
                cmd.Parameters.AddWithValue("@contractNo", 계약번호.Trim());
                cmd.Parameters.AddWithValue("@period",     계약기간.Trim());
                cmd.Parameters.AddWithValue("@company",    companyName.Trim());
                cmd.Parameters.AddWithValue("@point",      point.Trim());
                cmd.ExecuteNonQuery();
            }
            txn.Commit();
        }
        catch (Exception ex)
        {
            txn.Rollback();
            Debug.WriteLine($"[MeasurerService] 저장 오류: {ex.Message}");
            throw;
        }
    }

    // ── 업체별 관리 정보 조회 ────────────────────────────────────────────────
    public static (string 약칭, string 계약금액, string 견적구분) GetCompanyInfo(string companyName)
    {
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT IFNULL(약칭,''), IFNULL(계약금액,''), IFNULL(견적구분,'')
            FROM 측정인_채취지점 WHERE 업체명 = @company LIMIT 1";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        return r.Read()
            ? (r.GetString(0), r.GetString(1), r.GetString(2))
            : ("", "", "");
    }

    // ── 업체별 관리 정보 저장 ────────────────────────────────────────────────
    public static void UpdateCompanyInfo(string companyName, string 약칭, string 계약금액, string 견적구분)
    {
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE 측정인_채취지점
            SET 약칭 = @alias, 계약금액 = @amount, 견적구분 = @quotType
            WHERE 업체명 = @company";
        cmd.Parameters.AddWithValue("@alias",    약칭.Trim());
        cmd.Parameters.AddWithValue("@amount",   계약금액.Trim());
        cmd.Parameters.AddWithValue("@quotType", 견적구분.Trim());
        cmd.Parameters.AddWithValue("@company",  companyName.Trim());
        cmd.ExecuteNonQuery();
    }

    // ── 저장된 업체명 목록 조회 ──────────────────────────────────────────────
    public static List<string> GetCompanies()
    {
        var list = new List<string>();
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT 업체명 FROM 측정인_채취지점 ORDER BY 업체명 ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 특정 업체의 채취지점명 목록 조회 ─────────────────────────────────────
    public static List<string> GetSamplingPoints(string companyName)
    {
        var list = new List<string>();
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 채취지점명 FROM 측정인_채취지점 WHERE 업체명 = @company ORDER BY 채취지점명 ASC";
        cmd.Parameters.AddWithValue("@company", companyName);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    // ── 전체 데이터 조회 ──────────────────────────────────────────────────────
    public static List<(string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명)> GetAllData()
    {
        var list = new List<(string, string, string, string, string)>();
        using var conn = new SqliteConnection($"Data Source={GetDatabasePath()}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT 계약번호, IFNULL(약칭,''), IFNULL(계약기간,''), 업체명, 채취지점명
            FROM 측정인_채취지점
            ORDER BY 업체명, 채취지점명";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4)));
        return list;
    }
}
