using System;
using System.Data;
using System.Diagnostics;

namespace ETA.Services.Common;

/// <summary>
/// Phase 1 DB Migration: 테이블명 통일 + 레거시 테이블 DROP
/// </summary>
public static class DbPhase1Migration
{
    public static void ExecutePhase1()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            Console.WriteLine("\n[Phase 1] 테이블명 통일 (RENAME)");
            Console.WriteLine("".PadRight(60, '='));

            // 1. RENAME 테이블
            var renames = new[]
            {
                ("분석의뢰및결과", "수질분석센터_결과"),
                ("폐수의뢰및결과", "비용부담금_결과"),
                ("처리시설_측정결과", "처리시설_결과"),
            };

            foreach (var (oldName, newName) in renames)
            {
                try
                {
                    if (DbConnectionFactory.TableExists(conn, oldName))
                    {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = $"RENAME TABLE `{oldName}` TO `{newName}`";
                        cmd.ExecuteNonQuery();
                        Console.WriteLine($"✓ {oldName} → {newName}");
                    }
                    else
                    {
                        Console.WriteLine($"⊘ {oldName} (존재하지 않음)");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"✗ {oldName}: {e.Message}");
                }
            }

            Console.WriteLine("\n[Phase 1] 레거시 *_DATA 테이블 DROP");
            Console.WriteLine("".PadRight(60, '='));

            // 2. DROP 레거시 테이블
            var legacyTables = new[]
            {
                "BOD_DATA",
                "SS_DATA",
                "NHexan_DATA",
                "TN_DATA",
                "TP_DATA",
                "Phenols_DATA",
                "TOC_TCIC_DATA",
                "TOC_NPOC_DATA",
            };

            foreach (var tableName in legacyTables)
            {
                try
                {
                    if (DbConnectionFactory.TableExists(conn, tableName))
                    {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = $"DROP TABLE IF EXISTS `{tableName}`";
                        cmd.ExecuteNonQuery();
                        Console.WriteLine($"✓ DROP {tableName}");
                    }
                    else
                    {
                        Console.WriteLine($"⊘ {tableName} (존재하지 않음)");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"✗ {tableName}: {e.Message}");
                }
            }

            Console.WriteLine("\n✓ Phase 1 마이그레이션 완료");
        }
        catch (Exception e)
        {
            Console.WriteLine($"\n✗ 마이그레이션 실패: {e.Message}");
            Debug.WriteLine($"DbPhase1Migration Error: {e}");
        }
    }
}
