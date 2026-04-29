using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ETA.Services.Common;

/// <summary>
/// 현장 측정값(온도/pH/EC/염분/잔류염소/DO) + GPS + 채취 사진 저장을 위한 DB 자가치유.
///
/// 시료 채취 시 ES 04130.1f 규정상 즉시 측정해야 하는 항목들(수온, pH 등)이
/// 기존 DB 에 컬럼이 없어 데스크톱·모바일·API 가 공통으로 참조할 컬럼을 마련.
///
/// 진입점: 데스크톱 앱 시작 시 + ETA.Api 시작 시 한 번씩 `EnsureSchema()` 호출.
/// 이미 존재하는 컬럼/테이블은 건드리지 않음 (idempotent).
/// </summary>
public static class FieldMeasurementMigration
{
    /// <summary>현장 측정값 컬럼이 추가될 테이블 목록.</summary>
    private static readonly string[] SampleTables =
    {
        "분석의뢰및결과",
        "폐수의뢰및결과",
    };

    /// <summary>추가될 컬럼 정의 (이름, DDL).</summary>
    private static readonly (string name, string ddl)[] FieldColumns =
    {
        ("현장_온도",        "DECIMAL(5,2)  DEFAULT NULL"),
        ("현장_pH",          "DECIMAL(4,2)  DEFAULT NULL"),
        ("현장_전기전도도",  "DECIMAL(8,2)  DEFAULT NULL"),
        ("현장_염분",        "DECIMAL(5,2)  DEFAULT NULL"),
        ("현장_잔류염소",    "DECIMAL(6,3)  DEFAULT NULL"),
        ("현장_DO",          "DECIMAL(5,2)  DEFAULT NULL"),
        ("현장_GPS_위도",    "DECIMAL(10,7) DEFAULT NULL"),
        ("현장_GPS_경도",    "DECIMAL(10,7) DEFAULT NULL"),
        ("현장_측정자",      "VARCHAR(50)   DEFAULT ''"),
        ("현장_측정시각",    "VARCHAR(30)   DEFAULT ''"),
    };

    /// <summary>한 번 호출되면 통과한 결과 캐싱 (재호출 시 SHOW COLUMNS 생략).</summary>
    private static bool _done;

    /// <summary>스키마 보장. 이미 끝났으면 즉시 반환.</summary>
    public static void EnsureSchema()
    {
        if (_done) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            int columnsAdded = 0;
            foreach (var table in SampleTables)
            {
                if (!DbConnectionFactory.TableExists(conn, table))
                {
                    Log($"테이블 없음, 스킵: {table}");
                    continue;
                }
                var existing = new HashSet<string>(
                    DbConnectionFactory.GetColumnNames(conn, table),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var (name, ddl) in FieldColumns)
                {
                    if (existing.Contains(name)) continue;
                    try
                    {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = $"ALTER TABLE `{table}` ADD COLUMN `{name}` {ddl}";
                        cmd.ExecuteNonQuery();
                        columnsAdded++;
                        Log($"  + {table}.{name}");
                    }
                    catch (Exception ex) { Log($"  ✗ {table}.{name}: {ex.Message}"); }
                }
            }

            EnsurePhotoTable(conn);

            _done = true;
            if (columnsAdded > 0) Log($"완료: 컬럼 {columnsAdded}개 추가");
        }
        catch (Exception ex)
        {
            Log($"오류 (재시도 대기): {ex.Message}");
            // _done = false 유지 → 다음 호출 때 재시도
        }
    }

    private static void EnsurePhotoTable(System.Data.Common.DbConnection conn)
    {
        if (DbConnectionFactory.TableExists(conn, "시료_사진")) return;
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE `시료_사진` (
                    `_id`           INT AUTO_INCREMENT PRIMARY KEY,
                    `sample_table`  VARCHAR(50)   NOT NULL,
                    `sample_id`     INT           NOT NULL,
                    `file_path`     TEXT          NOT NULL,
                    `taken_at`      VARCHAR(30)   DEFAULT '',
                    `gps_위도`      DECIMAL(10,7) DEFAULT NULL,
                    `gps_경도`      DECIMAL(10,7) DEFAULT NULL,
                    `uploader`      VARCHAR(50)   DEFAULT '',
                    `등록일시`      VARCHAR(30)   DEFAULT '',
                    INDEX `ix_sample` (`sample_table`, `sample_id`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
            cmd.ExecuteNonQuery();
            Log("  + 테이블 생성: 시료_사진");
        }
        catch (Exception ex) { Log($"  ✗ 시료_사진 생성: {ex.Message}"); }
    }

    private static void Log(string msg) =>
        Debug.WriteLine($"[FieldMeasurementMigration] {msg}");
}
