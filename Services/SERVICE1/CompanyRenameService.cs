using System;
using System.Data.Common;
using System.Diagnostics;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 5개 테이블의 업체명을 oldName → newName 으로 일괄 변경합니다.
/// </summary>
public static class CompanyRenameService
{
    /// <summary>
    /// 5개 테이블의 업체명을 oldName → newName 으로 일괄 변경합니다.
    /// </summary>
    /// <returns>(updatedRows: 변경된 행 수, error: 오류 메시지 — 성공 시 빈 문자열)</returns>
    public static (int updatedRows, string error) RenameCompany(string oldName, string newName)
    {
        if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
            return (0, "변경 전/후 업체명이 비어 있습니다.");

        if (oldName == newName)
            return (0, "변경 전과 후 업체명이 동일합니다.");

        int total = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // ── 1. 견적발행내역.업체명 ──────────────────────────────────────
            total += ExecUpdate(conn,
                "UPDATE `견적발행내역` SET `업체명` = @newName WHERE `업체명` = @oldName",
                oldName, newName);

            // ── 2. 거래명세서발행내역.업체명 ────────────────────────────────
            total += ExecUpdate(conn,
                "UPDATE `거래명세서발행내역` SET `업체명` = @newName WHERE `업체명` = @oldName",
                oldName, newName);

            // ── 3. 계약 DB.C_CompanyName ────────────────────────────────────
            total += ExecUpdate(conn,
                "UPDATE `계약 DB` SET `C_CompanyName` = @newName WHERE `C_CompanyName` = @oldName",
                oldName, newName);

            // ── 4. 분석의뢰및결과.의뢰사업장 ────────────────────────────────
            total += ExecUpdate(conn,
                "UPDATE `분석의뢰및결과` SET `의뢰사업장` = @newName WHERE `의뢰사업장` = @oldName",
                oldName, newName);

            // ── 5. 시료명칭 컬럼명 변경 ─────────────────────────────────────
            // 컬럼이 존재하는 경우에만 실행
            if (DbConnectionFactory.ColumnExists(conn, "시료명칭", oldName))
            {
                using var alterCmd = conn.CreateCommand();
                if (DbConnectionFactory.IsMariaDb)
                {
                    // MariaDB: CHANGE COLUMN 구문 사용 (컬럼 타입 유지를 위해 TEXT 지정)
                    alterCmd.CommandText =
                        $"ALTER TABLE `시료명칭` CHANGE COLUMN `{EscapeBacktick(oldName)}` `{EscapeBacktick(newName)}` TEXT";
                }
                else
                {
                    // SQLite 3.25+: RENAME COLUMN 지원
                    alterCmd.CommandText =
                        $"ALTER TABLE `시료명칭` RENAME COLUMN `{EscapeBacktick(oldName)}` TO `{EscapeBacktick(newName)}`";
                }
                alterCmd.ExecuteNonQuery();
                Debug.WriteLine($"[CompanyRename] 시료명칭 컬럼 변경: {oldName} → {newName}");
            }
            else
            {
                Debug.WriteLine($"[CompanyRename] 시료명칭에 컬럼 '{oldName}' 없음 — 건너뜀");
            }

            Debug.WriteLine($"[CompanyRename] 완료: {total}행 변경, {oldName} → {newName}");
            return (total, "");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompanyRename] 오류: {ex.Message}");
            return (total, ex.Message);
        }
    }

    // ── 내부 헬퍼 ──────────────────────────────────────────────────────────────

    private static int ExecUpdate(DbConnection conn, string sql, string oldName, string newName)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@oldName", oldName);
            cmd.Parameters.AddWithValue("@newName", newName);
            int rows = cmd.ExecuteNonQuery();
            Debug.WriteLine($"[CompanyRename] {sql.Split('`')[1]}: {rows}행");
            return rows;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[CompanyRename] UPDATE 실패: {ex.Message}");
            throw;
        }
    }

    /// <summary>SQL 식별자 안의 백틱 이스케이프 (백틱 → 백틱 두 개)</summary>
    private static string EscapeBacktick(string s) => s.Replace("`", "``");
}
