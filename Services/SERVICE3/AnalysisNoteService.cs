using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text.Json;
using ETA.Services.Common;

namespace ETA.Services.SERVICE3;

/// <summary>
/// 분석노트 컬럼 설정 서비스
/// — 분석정보 테이블의 note_columns 컬럼 관리
/// — MyTask Show2에서 Excel 노트 생성 시 사용
/// </summary>
public static class AnalysisNoteService
{
    private static bool _ensured = false;

    private static void EnsureOnce()
    {
        if (_ensured) return;
        _ensured = true;   // 재진입 방지: 먼저 설정
        EnsureNoteColumnsColumn();
        EnsureConcentrationFormulaColumn();
        EnsureSchemaOverrideColumn();
        EnsureVolumeConstantColumn();
        EnsureParserColumnMapColumn();
        BulkEnsureSchemaOverrides();  // 앱 시작 시 schema_override 일괄 초기화
    }

    // ── note_columns 컬럼 확보 (최초 1회 자동 호출) ───────────────────────────
    public static void EnsureNoteColumnsColumn()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "note_columns"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `note_columns` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureColumn 오류: {ex.Message}");
        }
    }

    // ── 항목별 노트 컬럼 목록 조회 ───────────────────────────────────────────
    public static List<string> GetNoteColumns(string analyte)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `note_columns` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var raw = cmd.ExecuteScalar()?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            return raw.Split(',', StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => s.Trim())
                      .Where(s => !string.IsNullOrEmpty(s))
                      .ToList();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetNoteColumns 오류: {ex.Message}");
            return new List<string>();
        }
    }

    // ── 노트 컬럼 저장 ────────────────────────────────────────────────────────
    public static void SaveNoteColumns(string analyte, IEnumerable<string> cols)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `분석정보` SET `note_columns` = @v WHERE `Analyte` = @a";
            cmd.Parameters.AddWithValue("@v", string.Join(",", cols));
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] SaveNoteColumns 오류: {ex.Message}");
        }
    }

    // ── 스키마 오버라이드 컬럼 확보 + 조회/저장 ────────────────────────────────
    private static void EnsureSchemaOverrideColumn()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "schema_override"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `schema_override` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureSchemaOverride 오류: {ex.Message}");
        }
    }

    public static string? GetSchemaOverride(string analyte)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `schema_override` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var val = cmd.ExecuteScalar()?.ToString() ?? "";
            return string.IsNullOrWhiteSpace(val) ? null : val;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetSchemaOverride 오류: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// schema_override가 비어 있으면 autoSchema를 저장. 이미 값이 있으면 무시.
    /// 앱 로드 시 자동 판단 결과를 초기 등록하는 용도.
    /// </summary>
    public static void EnsureSchemaOverride(string analyte, string autoSchema)
    {
        if (string.IsNullOrWhiteSpace(analyte) || string.IsNullOrWhiteSpace(autoSchema)) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            // 현재 값 조회
            cmd.CommandText = "SELECT `schema_override` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var existing = cmd.ExecuteScalar()?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(existing)) return;  // 이미 설정됨

            // 비어 있으면 자동 판단값 저장
            using var upd = conn.CreateCommand();
            upd.CommandText = "UPDATE `분석정보` SET `schema_override` = @v WHERE `Analyte` = @a";
            upd.Parameters.AddWithValue("@v", autoSchema);
            upd.Parameters.AddWithValue("@a", analyte);
            upd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureSchemaOverride 오류: {ex.Message}");
        }
    }

    /// <summary>
    /// 분석정보 전체를 읽어 schema_override가 비어 있는 항목을 일괄 초기화.
    /// DetermineSchema 기준: 각 행의 Category/Method/instrument 컬럼 사용.
    /// 이미 값이 있는 항목은 건드리지 않음 (수동 수정 보존).
    /// </summary>
    public static void BulkEnsureSchemaOverrides()
    {
        try
        {
            var rows = new List<(string Analyte, string Category, string Method, string Instrument)>();
            using (var conn = DbConnectionFactory.CreateConnection())
            {
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    SELECT `Analyte`,
                           COALESCE(`Category`,''),
                           COALESCE(`Method`,''),
                           COALESCE(`instrument`,'')
                    FROM `분석정보`
                    WHERE `Analyte` IS NOT NULL AND `Analyte` <> ''
                      AND (`schema_override` IS NULL OR `schema_override` = '')";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                    rows.Add((r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3)));
            }

            if (rows.Count == 0) return;

            using var conn2 = DbConnectionFactory.CreateConnection();
            conn2.Open();
            foreach (var (analyte, cat, method, inst) in rows)
            {
                string schema = ETA.Services.SERVICE2.WaterCenterDbMigration.DetermineSchema(analyte, cat, method, inst);
                if (string.IsNullOrWhiteSpace(schema)) continue;
                using var upd = conn2.CreateCommand();
                upd.CommandText = @"UPDATE `분석정보` SET `schema_override` = @v
                    WHERE `Analyte` = @a AND (`schema_override` IS NULL OR `schema_override` = '')";
                upd.Parameters.AddWithValue("@v", schema);
                upd.Parameters.AddWithValue("@a", analyte);
                upd.ExecuteNonQuery();
            }
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] BulkEnsureSchemaOverrides: {rows.Count}건 초기화 완료");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[BulkEnsureSchemaOverrides] 오류: {ex.Message}");
        }
    }

    public static void SaveSchemaOverride(string analyte, string? schema)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `분석정보` SET `schema_override` = @v WHERE `Analyte` = @a";
            cmd.Parameters.AddWithValue("@v", schema ?? "");
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] SaveSchemaOverride 오류: {ex.Message}");
        }
    }

    // ── 농도 계산 공식 컬럼 확보 ─────────────────────────────────────────────
    private static void EnsureConcentrationFormulaColumn()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "concentration_formula"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `concentration_formula` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureFormula 오류: {ex.Message}");
        }
    }

    // ── 시료량 계수 컬럼 확보 ────────────────────────────────────────────────
    private static void EnsureVolumeConstantColumn()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "volume_constant"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `volume_constant` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureVolumeConstant 오류: {ex.Message}");
        }
    }

    // ── 시료량 계수 조회 ──────────────────────────────────────────────────────
    /// <summary>항목별 시료량 계수 반환. 미설정 시 fallback(기본 60) 반환.</summary>
    public static double GetVolumeConstant(string analyte, double fallback = 60.0)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `volume_constant` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var val = cmd.ExecuteScalar()?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(val) && double.TryParse(val, out var v) && v > 0)
                return v;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetVolumeConstant 오류: {ex.Message}");
        }
        return fallback;
    }

    // ── 시료량 계수 저장 ──────────────────────────────────────────────────────
    public static void SaveVolumeConstant(string analyte, double value)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `분석정보` SET `volume_constant` = @v WHERE `Analyte` = @a";
            cmd.Parameters.AddWithValue("@v", value > 0 ? value.ToString() : "");
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] SaveVolumeConstant 오류: {ex.Message}");
        }
    }

    // ── 농도 계산 공식 조회 ──────────────────────────────────────────────────
    public static string GetFormula(string analyte)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `concentration_formula` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            return cmd.ExecuteScalar()?.ToString() ?? "";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetFormula 오류: {ex.Message}");
            return "";
        }
    }

    // ── 파서-컬럼 매핑 컬럼 확보 ────────────────────────────────────────────────
    private static void EnsureParserColumnMapColumn()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "parser_column_map"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `parser_column_map` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] EnsureParserColumnMap 오류: {ex.Message}");
        }
    }

    // ── 파서-컬럼 매핑 조회 (JSON → Dictionary) ──────────────────────────────
    public static Dictionary<string, string> GetParserColumnMap(string analyte)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `parser_column_map` FROM `분석정보` WHERE `Analyte` = @a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var raw = cmd.ExecuteScalar()?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(raw)) return new Dictionary<string, string>();
            return JsonSerializer.Deserialize<Dictionary<string, string>>(raw)
                   ?? new Dictionary<string, string>();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetParserColumnMap 오류: {ex.Message}");
            return new Dictionary<string, string>();
        }
    }

    // ── 파서-컬럼 매핑 저장 (Dictionary → JSON) ──────────────────────────────
    public static void SaveParserColumnMap(string analyte, Dictionary<string, string> map)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `분석정보` SET `parser_column_map` = @v WHERE `Analyte` = @a";
            cmd.Parameters.AddWithValue("@v", map.Count > 0 ? JsonSerializer.Serialize(map) : "");
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] SaveParserColumnMap 오류: {ex.Message}");
        }
    }

    // ── 농도 계산 공식 저장 ──────────────────────────────────────────────────
    public static void SaveFormula(string analyte, string formula)
    {
        EnsureOnce();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE `분석정보` SET `concentration_formula` = @v WHERE `Analyte` = @a";
            cmd.Parameters.AddWithValue("@v", formula);
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] SaveFormula 오류: {ex.Message}");
        }
    }

    // ── 농도 계산 공식 평가기 ─────────────────────────────────────────────────
    /// <summary>
    /// concentration_formula 토큰 문자열을 실제 값으로 평가합니다.
    /// vars 예: { "흡광도"=0.179, "절편"=0.001, "기울기"=0.133, "시료량"=50, "희석배수"=1 }
    /// 평가 실패 또는 공식 없으면 null 반환 → 호출자가 기존 로직으로 폴백.
    /// </summary>
    public static double? EvaluateFormula(string formula, System.Collections.Generic.Dictionary<string, double> vars)
    {
        if (string.IsNullOrWhiteSpace(formula)) return null;
        try
        {
            // 1. 토큰 정리
            var tokens = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(t => t != "=" && t != "농도" && t != "✓" && t != "✗")
                .Select(t => t == "×" ? "*" : t == "÷" ? "/" : t)
                .ToList();

            // 2. 변수 → 숫자 치환
            var expr = tokens.Select(t =>
                vars.TryGetValue(t, out var v)
                    ? v.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : t
            ).ToList();

            // 3. Shunting-Yard → 즉시 평가 (스택 기반)
            var vals = new System.Collections.Generic.Stack<double>();
            var ops  = new System.Collections.Generic.Stack<string>();

            int Prec(string o) => o == "+" || o == "-" ? 1 : 2;
            bool IsOp(string s) => s is "+" or "-" or "*" or "/";

            void Apply()
            {
                var op = ops.Pop();
                var b  = vals.Pop();
                var a  = vals.Pop();
                vals.Push(op switch
                {
                    "+" => a + b,
                    "-" => a - b,
                    "*" => a * b,
                    "/" => b != 0 ? a / b : double.NaN,
                    _   => 0
                });
            }

            foreach (var tok in expr)
            {
                if (double.TryParse(tok,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num))
                {
                    vals.Push(num);
                }
                else if (IsOp(tok))
                {
                    while (ops.Count > 0 && IsOp(ops.Peek()) && Prec(ops.Peek()) >= Prec(tok))
                        Apply();
                    ops.Push(tok);
                }
                else if (tok == "(") ops.Push(tok);
                else if (tok == ")")
                {
                    while (ops.Count > 0 && ops.Peek() != "(") Apply();
                    if (ops.Count > 0) ops.Pop();
                }
                // 알 수 없는 토큰(변수 치환 안 된 것)은 0 으로 처리
                else vals.Push(0);
            }
            while (ops.Count > 0) Apply();

            if (vals.Count != 1) return null;
            var result = vals.Pop();
            return double.IsNaN(result) || double.IsInfinity(result) ? null : result;
        }
        catch { return null; }
    }

    // ── DB 기준 표준형식 목록 조회 ───────────────────────────────────────────
    /// <summary>
    /// 분석정보 전체 항목을 읽어 실제 사용 중인 스키마 목록을 반환 (중복 제거, 오버라이드 우선).
    /// </summary>
    public static List<string> GetDistinctSchemas()
    {
        EnsureOnce();
        var result = new System.Collections.Generic.SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT DISTINCT Analyte,
                COALESCE(Category, ''),
                COALESCE(Method, ''),
                COALESCE(instrument, '')
                FROM `분석정보`
                WHERE Analyte IS NOT NULL AND Analyte <> ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string a   = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                string ca  = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                string me  = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                string ins = r.IsDBNull(3) ? "" : r.GetString(3).Trim();
                if (string.IsNullOrEmpty(a)) continue;
                string? ov = GetSchemaOverride(a);
                string schema = string.IsNullOrWhiteSpace(ov)
                    ? ETA.Services.SERVICE2.WaterCenterDbMigration.DetermineSchema(a, ca, me, ins)
                    : ov;
                result.Add(schema);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] GetDistinctSchemas 오류: {ex.Message}");
        }
        return new List<string>(result);
    }

    // ── 같은 표준형식 항목 전체에 공식 일괄 적용 ────────────────────────────
    /// <summary>
    /// targetSchema 와 일치하는 항목(자동 + 오버라이드 모두 고려)에 formula 를 일괄 저장.
    /// 반환값: 적용된 항목 수
    /// </summary>
    public static int BulkSaveFormulaBySchema(string targetSchema, string formula)
    {
        EnsureOnce();
        int count = 0;
        try
        {
            // 1단계: Analyte + 스키마 판단용 컬럼 조회
            var rows = new List<(string Analyte, string Category, string Method, string Instrument)>();
            using (var conn1 = DbConnectionFactory.CreateConnection())
            {
                conn1.Open();
                using var cmd = conn1.CreateCommand();
                cmd.CommandText = @"SELECT DISTINCT Analyte,
                    COALESCE(Category, ''),
                    COALESCE(Method, ''),
                    COALESCE(instrument, '')
                    FROM `분석정보`
                    WHERE Analyte IS NOT NULL AND Analyte <> ''";
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string a   = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                    string ca  = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                    string me  = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                    string ins = r.IsDBNull(3) ? "" : r.GetString(3).Trim();
                    if (!string.IsNullOrEmpty(a)) rows.Add((a, ca, me, ins));
                }
            }

            System.Diagnostics.Debug.WriteLine($"[BulkSave] 전체 {rows.Count}개 항목, 대상 스키마: {targetSchema}");

            // 2단계: 각 항목 스키마 확인 후 일치 시 UPDATE
            using var conn2 = DbConnectionFactory.CreateConnection();
            conn2.Open();
            foreach (var (analyte, cat, method, inst) in rows)
            {
                string? ov = GetSchemaOverride(analyte);
                string effectiveSchema = string.IsNullOrWhiteSpace(ov)
                    ? ETA.Services.SERVICE2.WaterCenterDbMigration.DetermineSchema(analyte, cat, method, inst)
                    : ov;

                System.Diagnostics.Debug.WriteLine($"[BulkSave]  {analyte}: {effectiveSchema}");

                if (!effectiveSchema.Equals(targetSchema, StringComparison.OrdinalIgnoreCase)) continue;

                using var upd = conn2.CreateCommand();
                upd.CommandText = "UPDATE `분석정보` SET `concentration_formula` = @v WHERE `Analyte` = @a";
                upd.Parameters.AddWithValue("@v", formula);
                upd.Parameters.AddWithValue("@a", analyte);
                upd.ExecuteNonQuery();
                count++;
            }
            System.Diagnostics.Debug.WriteLine($"[BulkSave] 완료: {count}개");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AnalysisNoteService] BulkSaveFormula 오류: {ex.Message}\n{ex.StackTrace}");
        }
        return count;
    }

    // ── 검정곡선 컬럼 판별 ───────────────────────────────────────────────────
    // 패턴 1: ^ST\d+_ 접두사  /  패턴 2: 고정 이름 HashSet (이름 추가 시 여기만 수정)
    private static readonly System.Collections.Generic.HashSet<string> _calColNames =
        new(System.StringComparer.OrdinalIgnoreCase)
        {
            "기울기", "절편", "R값", "R^2", "검량선_a", "검량선_b",
        };

    public static bool IsCalibrationCol(string col) =>
        System.Text.RegularExpressions.Regex.IsMatch(col, @"^ST\d+_",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase) ||
        _calColNames.Contains(col);
}
