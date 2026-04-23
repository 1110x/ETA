using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE2;

namespace ETA.Services.SERVICE1;

public static class AnalysisService
{

    public static List<AnalysisItem> GetAllItems()
    {
        var items = new List<AnalysisItem>();

        using var connection = DbConnectionFactory.CreateConnection();
        connection.Open();

        // 약칭 컬럼 없으면 자동 추가 (마이그레이션)
        if (!DbConnectionFactory.ColumnExists(connection, "분석정보", "약칭"))
        {
            using var alt = connection.CreateCommand();
            alt.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `약칭` TEXT DEFAULT ''";
            try { alt.ExecuteNonQuery(); } catch { }
        }
        // AliasX 컬럼 없으면 자동 추가 (파서 키워드 자동 매핑용, 쉼표 구분)
        if (!DbConnectionFactory.ColumnExists(connection, "분석정보", "AliasX"))
        {
            using var alt = connection.CreateCommand();
            alt.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `AliasX` TEXT DEFAULT ''";
            try { alt.ExecuteNonQuery(); } catch { }
        }
        // 분석조건 컬럼 없으면 자동 추가 (시험기록부 상단 블록용, JSON Key/Value 배열)
        if (!DbConnectionFactory.ColumnExists(connection, "분석정보", "분석조건"))
        {
            using var alt = connection.CreateCommand();
            alt.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `분석조건` TEXT DEFAULT ''";
            try { alt.ExecuteNonQuery(); } catch { }
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
            SELECT Category,
            Analyte,
            COALESCE(`약칭`, '') AS `약칭`,
            COALESCE(`AliasX`, '') AS `AliasX`,
            Parts,
            DecimalPlaces,
            unit,
            ES,
            Method,
            instrument
            FROM `분석정보`
            ORDER BY ES ASC";

        using var reader = command.ExecuteReader();

        while (reader.Read())
        {
            var item = new AnalysisItem
            {
                Category      = reader.GetStringOrEmpty("Category"),
                Analyte       = reader.GetStringOrEmpty("Analyte"),
                약칭          = reader.GetStringOrEmpty("약칭"),
                Parts         = reader.GetStringOrEmpty("Parts"),
                DecimalPlaces = reader.IsDBNull(reader.GetOrdinal("DecimalPlaces")) ? 0 : reader.GetInt32(reader.GetOrdinal("DecimalPlaces")),
                unit          = reader.GetStringOrEmpty("unit"),
                ES            = reader.GetStringOrEmpty("ES"),
                Method        = reader.GetStringOrEmpty("Method"),
                instrument    = reader.GetStringOrEmpty("instrument"),
                AliasX        = reader.GetStringOrEmpty("AliasX")
            };
            items.Add(item);
            Console.WriteLine($"[GetAllItems] 아이템: {item.Analyte} ({item.Category})");
        }

        Console.WriteLine($"[GetAllItems] 총 아이템 수: {items.Count}");

        // 분석정보 항목 → 분장표준처리 컬럼 자동 동기화
        SyncColumnsToAssignmentTable(items);

        return items;
    }

    /// <summary>분석정보.Analyte 목록을 분장표준처리 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToAssignmentTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "분장표준처리")) return;

            // 분석정보 Analyte 목록 (인자로 받거나 DB에서 조회)
            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            // 분장표준처리 기존 컬럼
            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "분장표준처리"),
                StringComparer.OrdinalIgnoreCase);

            // 없는 항목 컬럼 추가
            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `분장표준처리` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToAssignmentTable] 오류: {ex.Message}"); }
    }

    /// <summary>분석정보.Analyte 목록을 수질분석센터_결과 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToRequestTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return;

            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `수질분석센터_결과` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToRequestTable] 오류: {ex.Message}"); }
    }

    /// <summary>분석정보.Analyte 목록을 계약 DB 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToContractTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "계약 DB")) return;

            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "계약 DB"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `계약 DB` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToContractTable] 오류: {ex.Message}"); }
    }

    /// <summary>분석정보.Analyte 목록을 방류기준표 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToDischargeTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "방류기준표")) return;

            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "방류기준표"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `방류기준표` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToDischargeTable] 오류: {ex.Message}"); }
    }

    /// <summary>분석정보.Analyte 목록을 견적발행내역 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToQuotationTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "견적발행내역")) return;

            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "견적발행내역"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `견적발행내역` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToQuotationTable] 오류: {ex.Message}"); }
    }

    /// <summary>분석정보.Analyte 목록을 거래명세서발행내역 컬럼에 동기화 (없는 컬럼 자동 추가)</summary>
    public static void SyncColumnsToStatementTable(List<AnalysisItem>? items = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "거래명세서발행내역")) return;

            var analytes = items?.Select(a => a.Analyte).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()
                        ?? new List<string>();
            if (analytes.Count == 0)
            {
                using var q = conn.CreateCommand();
                q.CommandText = "SELECT Analyte FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
                using var r = q.ExecuteReader();
                while (r.Read()) analytes.Add(r.GetString(0));
            }

            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "거래명세서발행내역"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var analyte in analytes)
            {
                if (existing.Contains(analyte)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `거래명세서발행내역` ADD COLUMN `{analyte.Replace("`", "")}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"[SyncColumnsToStatementTable] 오류: {ex.Message}"); }
    }

    /// <summary>약칭(영문 ID)으로 분석정보.Analyte(한글명) 조회. 없으면 null 반환</summary>
    /// <summary>Analyte 한글명으로 Category 조회. 없으면 null 반환</summary>
    public static string? GetAnalyteCategory(string analyte)
    {
        if (string.IsNullOrWhiteSpace(analyte)) return null;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `Category` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var result = cmd.ExecuteScalar();
            return result is string s && !string.IsNullOrWhiteSpace(s) ? s : null;
        }
        catch { return null; }
    }

    public static string? GetAnalyteByShortName(string shortName)
    {
        if (string.IsNullOrWhiteSpace(shortName)) return null;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `Analyte` FROM `분석정보` WHERE `약칭` = @abbrev LIMIT 1";
            cmd.Parameters.AddWithValue("@abbrev", shortName);
            var result = cmd.ExecuteScalar();
            return result is string s && !string.IsNullOrWhiteSpace(s) ? s : null;
        }
        catch { return null; }
    }

    // 약칭 또는 항목명 → 실제 시험기록부 테이블명 (분석정보 Analyte 기반)
    // 검색 우선순위: Analyte 직접 일치 → 약칭 일치 → AliasX 포함
    private static readonly Dictionary<string, string> _recordTableCache = new(StringComparer.OrdinalIgnoreCase);

    public static string? GetRecordTableName(string itemOrAbbr)
    {
        if (string.IsNullOrWhiteSpace(itemOrAbbr)) return null;
        if (_recordTableCache.TryGetValue(itemOrAbbr, out var cached)) return cached;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT `Analyte` FROM `분석정보`
                WHERE `Analyte` = @v OR `약칭` = @v
                   OR FIND_IN_SET(@v, REPLACE(COALESCE(`AliasX`,''), ' ', '')) > 0
                LIMIT 1";
            cmd.Parameters.AddWithValue("@v", itemOrAbbr);
            var analyte = cmd.ExecuteScalar() as string;
            if (string.IsNullOrEmpty(analyte)) return null;
            var tableName = WaterCenterDbMigration.SafeName(analyte) + "_시험기록부";
            _recordTableCache[itemOrAbbr] = tableName;
            return tableName;
        }
        catch { return null; }
    }

    // AliasX에 새 키워드 append — 이미 있으면 skip, 추가 시 캐시 무효화
    public static void AppendAliasX(string analyte, string keyword)
    {
        if (string.IsNullOrWhiteSpace(analyte) || string.IsNullOrWhiteSpace(keyword)) return;
        keyword = keyword.Trim();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COALESCE(`AliasX`, '') FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var current = cmd.ExecuteScalar() as string ?? "";
            var aliases = current.Split(',', System.StringSplitOptions.RemoveEmptyEntries)
                                 .Select(x => x.Trim()).ToList();
            if (aliases.Any(x => string.Equals(x, keyword, System.StringComparison.OrdinalIgnoreCase))) return;
            aliases.Add(keyword);
            using var upd = conn.CreateCommand();
            upd.CommandText = "UPDATE `분석정보` SET `AliasX` = @v WHERE `Analyte` = @a";
            upd.Parameters.AddWithValue("@v", string.Join(",", aliases));
            upd.Parameters.AddWithValue("@a", analyte);
            upd.ExecuteNonQuery();
            _recordTableCache.Remove(keyword);  // 캐시 무효화
        }
        catch { }
    }

    // 파서 키워드로 Analyte 조회 (약칭 + AliasX 검색)
    public static string? FindAnalyteByKeyword(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword)) return null;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT `Analyte` FROM `분석정보`
                WHERE `Analyte` = @v OR `약칭` = @v
                   OR FIND_IN_SET(@v, REPLACE(COALESCE(`AliasX`,''), ' ', '')) > 0
                LIMIT 1";
            cmd.Parameters.AddWithValue("@v", keyword);
            return cmd.ExecuteScalar() as string;
        }
        catch { return null; }
    }

    // 확장 메서드
    private static string GetStringOrEmpty(this DbDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        return !reader.IsDBNull(ordinal) ? reader.GetString(ordinal) : string.Empty;
    }
}