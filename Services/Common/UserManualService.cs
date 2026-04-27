using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.Common;

/// <summary>사용자메뉴얼 — 메뉴별 MD 문서 경로 매핑 + 문서 내용 조회.</summary>
public static class UserManualService
{
    public record ManualEntry(string Key, string Label, string Category, string MdPath, string Updated);

    private const string TableName = "사용자메뉴얼";
    private static string ManualsDir
    {
        get
        {
            var d = Path.Combine(AppPaths.RootPath, "Docs", "Manuals");
            try { Directory.CreateDirectory(d); } catch { }
            return d;
        }
    }

    public static void EnsureTable()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"CREATE TABLE IF NOT EXISTS `{TableName}` (
                `_id`     INT NOT NULL AUTO_INCREMENT,
                `메뉴키`  VARCHAR(100) NOT NULL,
                `메뉴명`  VARCHAR(200) NOT NULL,
                `카테고리` VARCHAR(100) NOT NULL,
                `MD경로`  VARCHAR(500) NOT NULL DEFAULT '',
                `최종수정` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                PRIMARY KEY (`_id`),
                UNIQUE KEY uk_menu (`메뉴키`)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.EnsureTable] {ex.Message}"); }
    }

    /// <summary>AccessService.AllMenus 기준으로 누락된 항목을 일괄 삽입.</summary>
    public static void SeedFromAccessMenus()
    {
        try
        {
            EnsureTable();
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            foreach (var (key, label, category) in AccessService.AllMenus)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"INSERT IGNORE INTO `{TableName}` (`메뉴키`, `메뉴명`, `카테고리`, `MD경로`)
                                    VALUES (@k, @l, @c, @p)";
                cmd.Parameters.AddWithValue("@k", key);
                cmd.Parameters.AddWithValue("@l", label);
                cmd.Parameters.AddWithValue("@c", category);
                cmd.Parameters.AddWithValue("@p", $"Docs/Manuals/{key}.md");
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.Seed] {ex.Message}"); }
    }

    public static List<ManualEntry> GetAll()
    {
        var list = new List<ManualEntry>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `메뉴키`, `메뉴명`, `카테고리`, `MD경로`, IFNULL(`최종수정`,'') FROM `{TableName}` ORDER BY `카테고리`, `_id`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ManualEntry(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4)));
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.GetAll] {ex.Message}"); }
        return list;
    }

    public static ManualEntry? GetByKey(string menuKey)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `메뉴키`, `메뉴명`, `카테고리`, `MD경로`, IFNULL(`최종수정`,'') FROM `{TableName}` WHERE `메뉴키`=@k LIMIT 1";
            cmd.Parameters.AddWithValue("@k", menuKey);
            using var r = cmd.ExecuteReader();
            if (r.Read())
                return new ManualEntry(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4));
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.GetByKey] {ex.Message}"); }
        return null;
    }

    public static void SetMdPath(string menuKey, string mdPath)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `{TableName}` SET `MD경로`=@p WHERE `메뉴키`=@k";
            cmd.Parameters.AddWithValue("@p", mdPath);
            cmd.Parameters.AddWithValue("@k", menuKey);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.SetMdPath] {ex.Message}"); }
    }

    /// <summary>MD 경로(상대/절대) 를 받아 실제 파일 내용을 읽음. 없으면 placeholder 반환.</summary>
    public static string ReadMd(string mdPath, string fallbackTitle)
    {
        try
        {
            var resolved = Path.IsPathRooted(mdPath)
                ? mdPath
                : Path.Combine(AppPaths.RootPath, mdPath);
            if (File.Exists(resolved))
                return File.ReadAllText(resolved);
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.ReadMd] {ex.Message}"); }
        return $"# {fallbackTitle}\n\n_아직 메뉴얼이 작성되지 않았습니다._\n\n경로: `{mdPath}`";
    }

    /// <summary>누락된 MD 파일에 대해 빈 스텁 생성 (있으면 건너뜀).</summary>
    public static void EnsureStubFiles()
    {
        try
        {
            foreach (var entry in GetAll())
            {
                if (string.IsNullOrWhiteSpace(entry.MdPath)) continue;
                var resolved = Path.IsPathRooted(entry.MdPath)
                    ? entry.MdPath
                    : Path.Combine(AppPaths.RootPath, entry.MdPath);
                if (File.Exists(resolved)) continue;
                Directory.CreateDirectory(Path.GetDirectoryName(resolved)!);
                var stub = $"# {entry.Label}\n\n_카테고리: {entry.Category}_\n\n## 개요\n\n(작성 예정)\n\n## 사용 흐름\n\n1. \n2. \n3. \n\n## 참고\n\n- \n";
                File.WriteAllText(resolved, stub);
            }
        }
        catch (Exception ex) { Console.WriteLine($"[UserManualService.EnsureStubFiles] {ex.Message}"); }
    }
}
