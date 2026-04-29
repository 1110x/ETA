using ETA.Views.Pages;
using ETA.Views.Pages.PAGE1;
using ETA.Views.Pages.PAGE2;
using ETA.Views.Pages.Common;
using System.Data;
using System.Data.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// "수질분석센터_결과" 테이블 접근 서비스
/// </summary>
public static class AnalysisRequestService
{

    private static readonly string LogPath =
        Path.Combine(ETA.Services.Common.AppPaths.LogsDir, "Treeview.log");

    // field69~114 마이그레이션은 앱 세션당 1회만 실행
    private static volatile bool _fieldMigrationDone = false;

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AnalysisRequestService] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // =====================================================================
    //  전체 레코드 목록 (트리뷰용)
    // =====================================================================
    public static List<AnalysisRequestRecord> GetAllRecords()
    {
        var list = new List<AnalysisRequestRecord>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과"))
        {
            Log("테이블 없음: 수질분석센터_결과");
            return list;
        }

        // MariaDB: using var 중첩 시 커넥션 충돌 방지 — 단일 using 블록
        bool has정도보증 = DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", "정도보증");
        var samplerCandidates = new[] { "시료채취1", "채수담당자", "시료채취자1", "시료채취자-1" };
        string? samplerCol = samplerCandidates.FirstOrDefault(c => DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", c));

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`,     ''),
                       COALESCE(`시료명`,   ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, ''),
                       {(has정도보증 ? "COALESCE(`정도보증`, '')" : "''")},
                       {(samplerCol != null ? $"COALESCE(`{samplerCol}`, '')" : "''")}
                FROM   `수질분석센터_결과`
                ORDER  BY `채취일자` DESC";

            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    var date = rdr.IsDBNull(4) ? "" : rdr.GetString(4);
                    string 연도 = "", 월 = "";
                    if (date.Length >= 10 && (date[4] == '-' || date[4] == '/'))
                    { 연도 = date[..4]; 월 = date[5..7]; }
                    else if (date.Length >= 8 && int.TryParse(date[..4], out _))
                    { 연도 = date[..4]; 월 = date[4..6]; }
                    else if (date.Length >= 4)
                    { 연도 = date[..4]; }

                    list.Add(new AnalysisRequestRecord
                    {
                        Id         = Convert.ToInt32(rdr.GetValue(0)),
                        약칭       = rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                        시료명     = rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                        접수번호   = rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                        의뢰일     = date,
                        채취일자   = date,
                        연도       = 연도,
                        월         = 월,
                        정도보증   = rdr.IsDBNull(5) ? "" : rdr.GetString(5),
                        채수담당자 = rdr.IsDBNull(6) ? "" : rdr.GetString(6),
                    });
                }
            }
        }

        Log($"완료: {list.Count}건");
        return list;
    }

    // =====================================================================
    //  단일 레코드 전체 컬럼 (DetailPanel 표시용)
    // =====================================================================
    public static Dictionary<string, string> GetRecordRow(int rowId)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureFieldMeasurementColumns(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id";
        cmd.Parameters.AddWithValue("@id", rowId);

        using var rdr = cmd.ExecuteReader();
        if (rdr.Read())
            for (int i = 0; i < rdr.FieldCount; i++)
            {
                var col = rdr.GetName(i).Trim();
                var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString() ?? "";
                dict[col] = val;
            }
        return dict;
    }

    // =====================================================================
    //  분석의뢰 항목 일괄 갱신 — 체크된 항목은 'O', 해제된 항목은 NULL
    // =====================================================================
    public static bool UpdateAnalyteValues(
        int rowId,
        IEnumerable<string> allAnalytes,
        IEnumerable<string> checkedAnalytes)
    {
        var all        = allAnalytes.Select(a => a.Trim()).Where(a => a.Length > 0).ToList();
        var checkedSet = new HashSet<string>(checkedAnalytes.Select(a => a.Trim()),
                             StringComparer.OrdinalIgnoreCase);
        if (all.Count == 0) return true;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            var setParts = all.Select((a, i) => $"`{a}` = @v{i}");
            cmd.CommandText =
                $"UPDATE `수질분석센터_결과` SET {string.Join(", ", setParts)} " +
                $"WHERE {DbConnectionFactory.RowId} = @id";
            for (int i = 0; i < all.Count; i++)
                cmd.Parameters.AddWithValue($"@v{i}",
                    checkedSet.Contains(all[i]) ? (object)"O" : DBNull.Value);
            cmd.Parameters.AddWithValue("@id", rowId);
            int rows = cmd.ExecuteNonQuery();
            Log($"UpdateAnalyteValues: rowId={rowId} → 체크={checkedSet.Count}, 해제={all.Count - checkedSet.Count}, {rows}행");
        }
        catch (Exception ex) { Log($"UpdateAnalyteValues 오류: {ex.Message}"); return false; }

        // 연결된 견적서 수량 재계산 (실패해도 기본 저장은 성공으로 처리)
        try { RecalcQuotationQuantities(rowId, all); }
        catch (Exception ex) { Log($"RecalcQuotationQuantities 오류: {ex.Message}"); }
        return true;
    }

    // =====================================================================
    //  견적서 수량 재계산
    //  같은 견적번호를 가진 수질분석센터_결과 행들에서 각 항목이 'O'인 횟수를 집계해
    //  견적발행내역의 수량과 소계를 갱신
    // =====================================================================
    private static void RecalcQuotationQuantities(int rowId, List<string> analyteNames)
    {
        var allSet = new HashSet<string>(analyteNames, StringComparer.OrdinalIgnoreCase);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 1. 이 레코드의 견적번호 조회
        string quotationNo;
        using (var c = conn.CreateCommand())
        {
            c.CommandText = $"SELECT COALESCE(`견적번호`, '') FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id";
            c.Parameters.AddWithValue("@id", rowId);
            quotationNo = c.ExecuteScalar()?.ToString()?.Trim() ?? "";
        }
        if (string.IsNullOrWhiteSpace(quotationNo))
        {
            Log($"RecalcQuotationQuantities: 견적번호 없음 (rowId={rowId})");
            return;
        }

        // 2. 견적발행내역 행 ID 조회
        int quotRowId;
        using (var c = conn.CreateCommand())
        {
            c.CommandText = $"SELECT {DbConnectionFactory.RowId} FROM `견적발행내역` WHERE `견적번호` = @no LIMIT 1";
            c.Parameters.AddWithValue("@no", quotationNo);
            var r = c.ExecuteScalar();
            if (r == null)
            {
                Log($"RecalcQuotationQuantities: 견적발행내역 없음 (견적번호={quotationNo})");
                return;
            }
            quotRowId = Convert.ToInt32(r);
        }

        // 3. 같은 견적번호의 모든 분석의뢰 행에서 항목별 'O' 개수 집계
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        using (var c = conn.CreateCommand())
        {
            c.CommandText = "SELECT * FROM `수질분석센터_결과` WHERE `견적번호` = @no";
            c.Parameters.AddWithValue("@no", quotationNo);
            using var rdr = c.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    var col = rdr.GetName(i).Trim();
                    if (!allSet.Contains(col)) continue;
                    var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (string.Equals(val, "O", StringComparison.OrdinalIgnoreCase))
                        counts[col] = counts.TryGetValue(col, out var n) ? n + 1 : 1;
                }
            }
        }

        // 4. 견적발행내역의 현재 단가 조회
        var quotRow  = QuotationService.GetIssueRow(quotRowId);
        var quotCols = new HashSet<string>(
            DbConnectionFactory.GetColumnNames(conn, "견적발행내역"),
            StringComparer.OrdinalIgnoreCase);

        // 5. 수량·소계 UPDATE 빌드
        var setParts = new List<string>();
        var pvals    = new List<(string p, object v)>();
        int idx = 0;

        foreach (var a in analyteNames)
        {
            if (!quotCols.Contains(a)) continue;
            int qty = counts.TryGetValue(a, out var c2) ? c2 : 0;

            setParts.Add($"`{a}` = @p{idx}");
            pvals.Add(($"@p{idx}", qty > 0 ? (object)qty : DBNull.Value)); idx++;

            if (quotCols.Contains(a + "소계"))
            {
                object subVal;
                if (qty > 0)
                {
                    var priceStr = quotRow.TryGetValue(a + "단가", out var ep) ? ep ?? "0" : "0";
                    decimal.TryParse(
                        priceStr.Replace("\u20a9", "").Replace(",", "").Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var price);
                    subVal = qty * price;
                }
                else subVal = DBNull.Value;

                setParts.Add($"`{a}소계` = @p{idx}");
                pvals.Add(($"@p{idx}", subVal)); idx++;
            }
        }

        if (setParts.Count == 0) return;

        using var upd = conn.CreateCommand();
        upd.CommandText =
            $"UPDATE `견적발행내역` SET {string.Join(", ", setParts)} " +
            $"WHERE {DbConnectionFactory.RowId} = @id";
        foreach (var (p, v) in pvals)
            upd.Parameters.AddWithValue(p, v);
        upd.Parameters.AddWithValue("@id", quotRowId);
        int updRows = upd.ExecuteNonQuery();
        Log($"RecalcQuotationQuantities: 견적번호={quotationNo} → {updRows}행 갱신 ({counts.Count}개 항목 집계)");
    }

    // =====================================================================
    //  레코드 삭제
    // =====================================================================
    public static bool DeleteRecord(int id)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"DELETE FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@id", id);
            int rows = cmd.ExecuteNonQuery();
            Log($"DeleteRecord: id={id} → {rows}행 삭제");
            return rows > 0;
        }
        catch (Exception ex) { Log($"DeleteRecord 오류: {ex.Message}"); return false; }
    }

    // =====================================================================
    //  방류기준 — 목록 조회 + 적용유무 갱신
    // =====================================================================

    /// <summary>방류기준표의 구분값(기준 이름) 목록 반환</summary>
    public static List<string> GetDischargeStandardNames()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "방류기준표")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT `구분` FROM `방류기준표` ORDER BY `구분`";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var v = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(v)) list.Add(v);
            }
        }
        catch
        {
            // 구분 컬럼명이 다를 수 있으므로 첫 번째 컬럼으로 재시도
            try
            {
                using var conn = DbConnectionFactory.CreateConnection();
                conn.Open();
                var cols = DbConnectionFactory.GetColumnNames(conn, "방류기준표");
                if (cols.Count == 0) return list;
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"SELECT DISTINCT `{cols[0]}` FROM `방류기준표` ORDER BY `{cols[0]}`";
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var v = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(v)) list.Add(v);
                }
            }
            catch { }
        }
        return list;
    }

    // =====================================================================
    //  현장측정항목 — 채수 후 입력하는 5개 측정값
    //  (온도, pH, 용존산소, 전기전도도, 잔류염소)
    // =====================================================================
    public static readonly string[] FieldMeasurementCols =
        new[] { "현장_온도", "현장_pH", "현장_용존산소", "현장_전기전도도", "현장_잔류염소" };

    /// <summary>생태독성 담당자가 입력하는 추가 측정값 (염분/암모니아/경도)</summary>
    public static readonly string[] EcotoxFieldCols =
        new[] { "생태_염분", "생태_암모니아", "생태_경도" };

    /// <summary>수질분석센터_결과 테이블에 현장측정 + 생태독성 입력 컬럼 자가치유</summary>
    public static void EnsureFieldMeasurementColumns(DbConnection conn)
    {
        try
        {
            var existing = new HashSet<string>(
                DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과"),
                StringComparer.OrdinalIgnoreCase);
            foreach (var col in FieldMeasurementCols.Concat(EcotoxFieldCols))
            {
                if (existing.Contains(col)) continue;
                using var alt = conn.CreateCommand();
                alt.CommandText = $"ALTER TABLE `수질분석센터_결과` ADD COLUMN `{col}` TEXT DEFAULT NULL";
                try { alt.ExecuteNonQuery(); existing.Add(col); }
                catch (Exception ex) { Log($"FieldMeasurement {col} 추가 실패: {ex.Message}"); }
            }
        }
        catch (Exception ex) { Log($"EnsureFieldMeasurementColumns 오류: {ex.Message}"); }
    }

    /// <summary>현장측정 5개 항목 일괄 저장</summary>
    public static bool UpdateFieldMeasurements(int rowId, Dictionary<string, string> values)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureFieldMeasurementColumns(conn);

            using var cmd = conn.CreateCommand();
            var setParts = FieldMeasurementCols.Select((c, i) => $"`{c}` = @v{i}").ToList();
            cmd.CommandText =
                $"UPDATE `수질분석센터_결과` SET {string.Join(", ", setParts)} " +
                $"WHERE {DbConnectionFactory.RowId} = @id";
            for (int i = 0; i < FieldMeasurementCols.Length; i++)
            {
                values.TryGetValue(FieldMeasurementCols[i], out var v);
                cmd.Parameters.AddWithValue($"@v{i}",
                    string.IsNullOrWhiteSpace(v) ? (object)DBNull.Value : v.Trim());
            }
            cmd.Parameters.AddWithValue("@id", rowId);
            return cmd.ExecuteNonQuery() > 0;
        }
        catch (Exception ex) { Log($"UpdateFieldMeasurements 오류: {ex.Message}"); return false; }
    }

    /// <summary>생태독성 담당자 입력 3개 항목(염분/암모니아/경도) 일괄 저장</summary>
    public static bool UpdateEcotoxFields(int rowId, Dictionary<string, string> values)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureFieldMeasurementColumns(conn);

            using var cmd = conn.CreateCommand();
            var setParts = EcotoxFieldCols.Select((c, i) => $"`{c}` = @v{i}").ToList();
            cmd.CommandText =
                $"UPDATE `수질분석센터_결과` SET {string.Join(", ", setParts)} " +
                $"WHERE {DbConnectionFactory.RowId} = @id";
            for (int i = 0; i < EcotoxFieldCols.Length; i++)
            {
                values.TryGetValue(EcotoxFieldCols[i], out var v);
                cmd.Parameters.AddWithValue($"@v{i}",
                    string.IsNullOrWhiteSpace(v) ? (object)DBNull.Value : v.Trim());
            }
            cmd.Parameters.AddWithValue("@id", rowId);
            return cmd.ExecuteNonQuery() > 0;
        }
        catch (Exception ex) { Log($"UpdateEcotoxFields 오류: {ex.Message}"); return false; }
    }

    /// <summary>수질분석센터_결과의 방류허용기준 적용유무 업데이트</summary>
    public static bool UpdateDischargeStandard(int rowId, string standardName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"UPDATE `수질분석센터_결과` SET `방류허용기준 적용유무` = @val " +
                $"WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@val", standardName);
            cmd.Parameters.AddWithValue("@id",  rowId);
            return cmd.ExecuteNonQuery() > 0;
        }
        catch (Exception ex) { Log($"UpdateDischargeStandard 오류: {ex.Message}"); return false; }
    }

    // =====================================================================
    //  DB Migration — 약칭 기준 고유 시료명 조회 (4번째 컬럼 약칭, 6번째 컬럼 시료명)
    // =====================================================================
    public static List<string> GetDistinctSampleNames(string 약칭)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `시료명`
                FROM `수질분석센터_결과`
                WHERE `약칭` = @abbr AND `시료명` IS NOT NULL AND `시료명` <> ''
                ORDER BY `시료명` ASC";
            cmd.Parameters.AddWithValue("@abbr", 약칭);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
        }
        catch (Exception ex) { Log($"GetDistinctSampleNames 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  DB Migration — 시료명 일괄 갱신
    // =====================================================================
    /// <summary>
    /// 약칭이 일치하는 레코드 중 기존 시료명을 새 시료명으로 UPDATE한다.
    /// </summary>
    public static int RenameSampleName(string 약칭, string oldName, string newName)
    {
        if (string.IsNullOrWhiteSpace(newName)) return 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                UPDATE `수질분석센터_결과`
                SET `시료명` = @newName
                WHERE `약칭` = @abbr AND `시료명` = @oldName";
            cmd.Parameters.AddWithValue("@newName", newName.Trim());
            cmd.Parameters.AddWithValue("@abbr", 약칭);
            cmd.Parameters.AddWithValue("@oldName", oldName);
            int rows = cmd.ExecuteNonQuery();
            Log($"RenameSampleName: {약칭} / '{oldName}' → '{newName}' ({rows}행)");
            return rows;
        }
        catch (Exception ex) { Log($"RenameSampleName 오류: {ex.Message}"); return 0; }
    }

    // =====================================================================
    //  명칭 정리 — 전체 약칭 목록
    // =====================================================================
    public static List<string> GetDistinctAbbreviations()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `약칭`
                FROM `수질분석센터_결과`
                WHERE `약칭` IS NOT NULL AND `약칭` <> ''
                ORDER BY `약칭` ASC";
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
        }
        catch (Exception ex) { Log($"GetDistinctAbbreviations 오류: {ex.Message}"); }
        return list;
    }

    /// <summary>특정 약칭+시료명의 의뢰일 목록 (중복 제거, 최신순)</summary>
    public static List<string> GetRequestDatesBySample(string 약칭, string 시료명)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(시료명)) return list;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `채취일자`
                FROM `수질분석센터_결과`
                WHERE `약칭` = @abbr AND `시료명` = @name
                  AND `채취일자` IS NOT NULL AND `채취일자` <> ''
                ORDER BY `채취일자` DESC";
            cmd.Parameters.AddWithValue("@abbr", 약칭 ?? "");
            cmd.Parameters.AddWithValue("@name", 시료명);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var s = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                if (!string.IsNullOrEmpty(s)) list.Add(s);
            }
        }
        catch (Exception ex) { Log($"GetRequestDatesBySample 오류: {ex.Message}"); }
        return list;
    }

    /// <summary>약칭 → 대표 의뢰사업장(업체명) 매핑</summary>
    public static Dictionary<string, string> GetAbbrCompanyMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            // 각 약칭별 최빈 의뢰사업장을 선택 (가장 많이 나온 값)
            cmd.CommandText = @"
                SELECT `약칭`, `의뢰사업장`, COUNT(*) AS cnt
                FROM `수질분석센터_결과`
                WHERE `약칭` IS NOT NULL AND `약칭` <> ''
                  AND `의뢰사업장` IS NOT NULL AND `의뢰사업장` <> ''
                GROUP BY `약칭`, `의뢰사업장`
                ORDER BY `약칭`, cnt DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var abbr = r.GetString(0);
                var company = r.GetString(1);
                if (!map.ContainsKey(abbr)) map[abbr] = company;
            }
        }
        catch (Exception ex) { Log($"GetAbbrCompanyMap 오류: {ex.Message}"); }
        return map;
    }

    // =====================================================================
    //  분석정보 — ES 순서대로 항목 목록 반환 (Show4 정렬용)
    //    반환: (fullName=Analyte, shortName=약칭) ordered by ES ASC
    // =====================================================================
    public static List<(string fullName, string shortName)> GetOrderedAnalytes()
    {
        var result = new List<(string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Analyte, COALESCE(`약칭`,'') FROM `분석정보` WHERE Analyte IS NOT NULL AND Analyte <> '' ORDER BY ES ASC";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string full  = rdr.GetString(0).Trim();
                string alias = rdr.IsDBNull(1) ? "" : (rdr.GetString(1).Trim());
                string sn    = string.IsNullOrWhiteSpace(alias) ? full : alias;
                result.Add((full, sn));
            }
        }
        catch (Exception ex) { Log($"GetOrderedAnalytes 오류: {ex.Message}"); }
        return result;
    }

    // =====================================================================
    //  명칭 정리 — 분장표준처리 — 약칭 직접 조회
    //    항목명='약칭' 행을 찾아 컬럼헤드 → 약칭 매핑 반환
    //    예: "생물화학적 산소요구량" → "BOD", "수소이온농도" → "pH"
    // =====================================================================
    public static Dictionary<string, string> GetShortNames()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "분장표준처리")) return result;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = '약칭' LIMIT 1";
            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string colHeader = rdr.GetName(i).Trim();
                    if (colHeader.StartsWith("_") || colHeader == "항목명") continue;
                    string abbr      = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(colHeader) && !string.IsNullOrEmpty(abbr))
                        result[colHeader] = abbr;
                }
            }
            Log($"GetShortNames: {result.Count}개 로드");
        }
        catch (Exception ex) { Log($"GetShortNames 오류: {ex.Message}"); }
        return result;
    }

    // =====================================================================
    //  분장표준처리 — 항목 정보 조회
    //
    //  실제 테이블 구조:
    //    항목명        | 부유물질 | 생물화학적 산소요구량 | 수소이온농도 | ...
    //    표준처리기한   |    7    |          9            |      2       | ...
    //    약칭          |   SS    |         BOD           |     pH       | ...
    //    2025-07-01   |  김지은  |        정준하         |    유경지    | ...
    //    2025-07-02   |  김지은  |        정준하         |    유경지    | ...
    //
    //  반환: 컬럼 전체명 → (처리일수, 약식명)
    //    예: "생물화학적 산소요구량" → (9, "BOD")
    // =====================================================================
    public static Dictionary<string, (int days, string shortName)> GetStandardDaysInfo()
    {
        var result = new Dictionary<string, (int, string)>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            if (!DbConnectionFactory.TableExists(conn, "분장표준처리"))
            { Log("GetStandardDaysInfo: 분장표준처리 테이블 없음"); return result; }

            // 기타업무/담당계약업체 컬럼 자동 추가
            foreach (var col in new[] { "기타업무", "담당계약업체" })
            {
                if (!DbConnectionFactory.ColumnExists(conn, "분장표준처리", col))
                {
                    try
                    {
                        using var alt = conn.CreateCommand();
                        alt.CommandText = $"ALTER TABLE `분장표준처리` ADD COLUMN `{col}` TEXT DEFAULT ''";
                        alt.ExecuteNonQuery();
                        Log($"분장표준처리에 {col} 컬럼 추가");
                    }
                    catch { }
                }
            }

            // field69~field114 컬럼 삭제 마이그레이션 — 앱 세션당 1회만 실행
            if (!_fieldMigrationDone)
            {
                _fieldMigrationDone = true;
                var existingCols = new HashSet<string>(
                    DbConnectionFactory.GetColumnNames(conn, "분장표준처리"),
                    StringComparer.OrdinalIgnoreCase);
                foreach (var fieldName in existingCols.Where(c =>
                    c.StartsWith("field", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(c[5..], out int n) && n >= 69 && n <= 114).ToList())
                {
                    try
                    {
                        // 타임아웃 방지: 컬럼마다 별도 연결 사용
                        using var dropConn = DbConnectionFactory.CreateConnection();
                        dropConn.Open();
                        using var drop = dropConn.CreateCommand();
                        drop.CommandText = $"ALTER TABLE `분장표준처리` DROP COLUMN `{fieldName}`";
                        drop.ExecuteNonQuery();
                        Log($"분장표준처리에서 {fieldName} 컬럼 삭제");
                    }
                    catch (Exception ex) { Log($"분장표준처리 {fieldName} 삭제 실패: {ex.Message}"); }
                }
            }

            // 처음 3행: 표준처리기한, 약칭, (날짜행)
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` LIMIT 3";
            using var rdr = cmd.ExecuteReader();

            int      fc       = rdr.FieldCount;
            string[] headers  = new string[fc];
            int[]    daysArr  = new int[fc];
            string[] shortArr = new string[fc];
            bool     daysSet  = false;
            bool     shortSet = false;

            for (int i = 0; i < fc; i++)
                headers[i] = rdr.GetName(i).Trim();

            while (rdr.Read())
            {
                string label = rdr["항목명"]?.ToString()?.Trim() ?? "";

                if (!daysSet && label == "표준처리기한")
                {
                    for (int i = 0; i < fc; i++)
                    {
                        if (headers[i].StartsWith("_") || headers[i] == "항목명") continue;
                        var v = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        int.TryParse(v, out daysArr[i]);
                    }
                    daysSet = true;
                    continue;
                }

                if (!shortSet && label == "약칭")
                {
                    for (int i = 0; i < fc; i++)
                    {
                        if (headers[i].StartsWith("_") || headers[i] == "항목명") continue;
                        shortArr[i] = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    }
                    shortSet = true;
                    continue;
                }

                if (daysSet && shortSet) break;
            }

            for (int i = 0; i < fc; i++)
            {
                string fullName = headers[i];
                if (string.IsNullOrEmpty(fullName) || fullName.StartsWith("_") || fullName == "항목명") continue;
                string shortName = string.IsNullOrEmpty(shortArr[i]) ? fullName : shortArr[i];
                result[fullName] = (daysArr[i], shortName);
            }

            Log($"GetStandardDaysInfo: {result.Count}개 로드 (days={daysSet}, short={shortSet})");
        }
        catch (Exception ex) { Log($"GetStandardDaysInfo 오류: {ex.Message}"); }

        return result;
    }

    // =====================================================================
    //  분장표준처리 — 채취일자 기준 전체 담당자 한번에 조회
    //
    //  반환: 분석항목 전체명 → 담당자 이름
    //    예: "생물화학적 산소요구량" → "정준하"
    //
    //  동작: 항목명 컬럼이 채취일자와 일치하는 행을 찾아 각 컬럼 값(담당자) 반환
    // =====================================================================
    public static Dictionary<string, string> GetManagersByDate(string sampleDate)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 날짜 정규화: yyyy-MM-dd
            string dateKey = sampleDate.Trim();
            if (DateTime.TryParse(dateKey, out var dt))
                dateKey = dt.ToString("yyyy-MM-dd");

            // 1차: 정확한 날짜 조회
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT * FROM `분장표준처리` WHERE `항목명` = @date LIMIT 1";
                cmd.Parameters.AddWithValue("@date", dateKey);

                using var rdr = cmd.ExecuteReader();
                if (rdr.Read())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        string colName = rdr.GetName(i).Trim();
                        if (colName.StartsWith("_") || colName == "항목명") continue;
                        string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(colName))
                            result[colName] = manager;
                    }
                }
            }

            // 2차: 해당 날짜에 배정이 없으면 실제 배정이 있는 가장 최근 날짜로 fallback
            bool hasAny = result.Values.Any(v => !string.IsNullOrEmpty(v));
            if (!hasAny)
            {
                using var cmd2 = conn.CreateCommand();
                cmd2.CommandText =
                    "SELECT * FROM `분장표준처리` WHERE `항목명` <= @date AND `항목명` REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' ORDER BY `항목명` DESC";
                cmd2.Parameters.AddWithValue("@date", dateKey);

                using var rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    string fallbackDate = rdr2.IsDBNull(0) ? "" : rdr2.GetString(0);
                    var candidate = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < rdr2.FieldCount; i++)
                    {
                        string colName = rdr2.GetName(i).Trim();
                        if (colName.StartsWith("_") || colName == "항목명") continue;
                        string manager = rdr2.IsDBNull(i) ? "" : rdr2.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(colName))
                            candidate[colName] = manager;
                    }
                    // 실제 배정이 있는 행을 찾으면 적용 후 탈출
                    if (candidate.Values.Any(v => !string.IsNullOrEmpty(v)))
                    {
                        foreach (var kv in candidate) result[kv.Key] = kv.Value;
                        Log($"GetManagersByDate({dateKey}): 배정 없음 → {fallbackDate} 기준 적용");
                        break;
                    }
                }
            }
            Log($"GetManagersByDate({dateKey}): {result.Count}개 담당자 로드");
        }
        catch (Exception ex) { Log($"GetManagersByDate 오류: {ex.Message}"); }

        return result;
    }

    // ── 하위 호환용 ──────────────────────────────────────────────────────
    public static Dictionary<string, int> GetStandardDays()
    {
        var info   = GetStandardDaysInfo();
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in info)
            result[kv.Key] = kv.Value.days;
        return result;
    }

    // =====================================================================
    //  분장표준처리 — 특정 직원의 업무 분장 조회 (날짜 지정)
    //  queryDate: 조회할 날짜 (해당 월의 첫 날짜부터 마지막 날짜까지)
    //  반환: 분석항목 약식명 리스트
    // =====================================================================
    public static List<(string FullName, string ShortName)> GetAssignmentsForAgent(string employeeId, DateTime queryDate)
    {
        var assignments = new List<(string, string)>();

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 담당자 이름 조회
            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var result = cmd2.ExecuteScalar();
                if (result != null)
                    managerName = result.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return assignments;

            // 지정된 날짜 기준으로 할당된 항목 조회
            string queryDateStr = queryDate.ToString("yyyy-MM-dd");
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `분장표준처리` WHERE `항목명` = @date";
            cmd.Parameters.AddWithValue("@date", queryDateStr);

            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                // 분석항목 정보 가져오기
                var info = GetStandardDaysInfo();
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (manager == managerName)
                    {
                        // 약식명 찾기
                        foreach (var kv in info)
                        {
                            if (kv.Key == colName)
                            {
                                assignments.Add((kv.Key, kv.Value.shortName));
                                break;
                            }
                        }
                    }
                }
            }

            Log($"GetAssignmentsForAgent({employeeId}/{managerName}, {queryDateStr}): {assignments.Count}개 할당");
        }
        catch (Exception ex) { Log($"GetAssignmentsForAgent 오류: {ex.Message}"); }

        return assignments;
    }

    public static List<(string FullName, string ShortName)> GetAssignmentsForAgent(string employeeId)
        => GetAssignmentsForAgent(employeeId, DateTime.Today);

    // =====================================================================
    //  분장표준처리 — 기간 범위 조회 (중복 제거, 날짜 범위 내 할당된 항목 전체)
    // =====================================================================
    public static List<(string FullName, string ShortName)> GetAssignmentsForAgentRange(
        string employeeId, DateTime startDate, DateTime endDate)
    {
        var result = new List<(string, string)>();

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return result;

            var info = GetStandardDaysInfo();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", startDate.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", endDate.ToString("yyyy-MM-dd"));

            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (manager == managerName && seen.Add(colName))
                    {
                        if (info.TryGetValue(colName, out var meta))
                            result.Add((colName, meta.shortName));
                        else
                            result.Add((colName, colName));
                    }
                }
            }
            Log($"GetAssignmentsForAgentRange({employeeId}, {startDate:yyyy-MM-dd}~{endDate:yyyy-MM-dd}): {result.Count}개");
        }
        catch (Exception ex) { Log($"GetAssignmentsForAgentRange 오류: {ex.Message}"); }

        return result;
    }

    // =====================================================================
    //  분장표준처리 — 전체 기간, 중복 허용 (날짜별 1건씩 → 일수 집계용)
    // =====================================================================
    public static List<(string FullName, string ShortName)> GetAssignmentDaysForAgentAll(string employeeId)
    {
        var result = new List<(string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return result;

            var info = GetStandardDaysInfo();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` ORDER BY `항목명`";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (manager == managerName)
                    {
                        if (info.TryGetValue(colName, out var meta))
                            result.Add((colName, meta.shortName));
                        else
                            result.Add((colName, colName));
                    }
                }
            }
        }
        catch (Exception ex) { Log($"GetAssignmentDaysForAgentAll 오류: {ex.Message}"); }
        return result;
    }

    // =====================================================================
    //  분장표준처리 — 기간 내 특정 항목을 담당하는 직원 목록 (중복 제거)
    // =====================================================================
    /// <summary>
    /// 기간 내 특정 항목 담당자 목록 — 날짜별 중복 포함.
    /// 같은 사람이 여러 날 담당하면 여러 번 포함됩니다.
    /// </summary>
    public static List<string> GetAssigneesForAnalyteInRange(string analyteFullName, DateTime start, DateTime end)
    {
        var assignees = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 1. 범위 내 날짜별 할당자 맵 구축
            var dayMap = new Dictionary<string, string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
                cmd.Parameters.AddWithValue("@s", start.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@e", end.ToString("yyyy-MM-dd"));

                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var dateStr = rdr["항목명"]?.ToString()?.Trim() ?? "";
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        string cn = rdr.GetName(i).Trim();
                        if (cn.StartsWith("_") || cn == "항목명") continue;
                        if (!string.Equals(cn, analyteFullName,
                            StringComparison.OrdinalIgnoreCase)) continue;
                        var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val) && !string.IsNullOrEmpty(dateStr))
                            dayMap[dateStr] = val;
                        break;
                    }
                }
            }

            // 2. 범위 시작일 이전의 가장 최근 할당자 (자동 할당 fallback)
            string fallback = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText =
                    "SELECT * FROM `분장표준처리` WHERE `항목명` < @s ORDER BY `항목명` DESC";
                cmd2.Parameters.AddWithValue("@s", start.ToString("yyyy-MM-dd"));

                using var rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    for (int i = 1; i < rdr2.FieldCount; i++)
                    {
                        if (!string.Equals(rdr2.GetName(i).Trim(), analyteFullName,
                            StringComparison.OrdinalIgnoreCase)) continue;
                        var val = rdr2.IsDBNull(i) ? "" : rdr2.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val)) fallback = val;
                        break;
                    }
                    if (!string.IsNullOrEmpty(fallback)) break;
                }
            }

            // 3. 날짜별 순회 — 빈 날짜는 직전 할당자를 자동 적용
            string lastKnown = fallback;
            for (var d = start; d <= end; d = d.AddDays(1))
            {
                var key = d.ToString("yyyy-MM-dd");
                if (dayMap.TryGetValue(key, out var assignee))
                {
                    lastKnown = assignee;
                    assignees.Add(assignee);
                }
                else if (!string.IsNullOrEmpty(lastKnown))
                {
                    assignees.Add(lastKnown); // 자동 할당
                }
            }
        }
        catch (Exception ex) { Log($"GetAssigneesForAnalyteInRange 오류: {ex.Message}"); }
        return assignees;
    }

    // =====================================================================
    //  분장표준처리 — 특정 날짜에 특정 항목을 담당하는 직원 목록
    // =====================================================================
    public static List<string> GetAssigneesForAnalyteOnDate(string analyteFullName, DateTime date)
    {
        var assignees = new List<string>();

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 1. 해당 날짜에서 직접 조회
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = @date";
                cmd.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));

                using var rdr = cmd.ExecuteReader();
                if (rdr.Read())
                {
                    for (int i = 1; i < rdr.FieldCount; i++)
                    {
                        if (!string.Equals(rdr.GetName(i).Trim(), analyteFullName,
                            StringComparison.OrdinalIgnoreCase)) continue;
                        var val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val)) assignees.Add(val);
                        break;
                    }
                }
            }

            // 2. 할당자가 없으면 → 가장 최근 날짜의 할당자를 자동 적용
            if (assignees.Count == 0)
            {
                using var cmd2 = conn.CreateCommand();
                cmd2.CommandText =
                    "SELECT * FROM `분장표준처리` WHERE `항목명` < @date ORDER BY `항목명` DESC";
                cmd2.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));

                using var rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    for (int i = 1; i < rdr2.FieldCount; i++)
                    {
                        if (!string.Equals(rdr2.GetName(i).Trim(), analyteFullName,
                            StringComparison.OrdinalIgnoreCase)) continue;
                        var val = rdr2.IsDBNull(i) ? "" : rdr2.GetValue(i)?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(val)) assignees.Add(val);
                        break;
                    }
                    if (assignees.Count > 0) break; // 찾으면 즉시 종료
                }
            }
        }
        catch (Exception ex) { Log($"GetAssigneesForAnalyteOnDate 오류: {ex.Message}"); }

        return assignees;
    }

    // =====================================================================
    //  분장표준처리 — 특정 항목의 특정 직원 분장을 지정 날짜부터 제거
    // =====================================================================
    public static void ClearAnalyteFromAgent(string employeeId, string analyteFullName, DateTime from)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                UPDATE `분장표준처리`
                SET `{analyteFullName}` = NULL
                WHERE `{analyteFullName}` = @mgr
                  AND `항목명` >= @from";
            cmd.Parameters.AddWithValue("@mgr",  managerName);
            cmd.Parameters.AddWithValue("@from", from.ToString("yyyy-MM-dd"));
            cmd.ExecuteNonQuery();
            Log($"ClearAnalyteFromAgent: {employeeId}({managerName}) 항목='{analyteFullName}' from={from:yyyy-MM-dd}");
        }
        catch (Exception ex) { Log($"ClearAnalyteFromAgent 오류: {ex.Message}"); }
    }

    // =====================================================================
    //  분장표준처리 — 특정 직원의 기간 내 모든 분장을 NULL로 초기화
    // =====================================================================
    public static void ClearAssignmentsForAgent(string employeeId, DateTime startDate, DateTime endDate)
    {

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return;

            // 컬럼 목록 (항목명 제외)
            var columns = DbConnectionFactory.GetColumnNames(conn, "분장표준처리")
                .Where(c => c != "항목명").ToList();
            if (columns.Count == 0) return;

            // 해당 기간 내 이 직원이 담당인 셀을 NULL로
            var setClauses = columns.Select(
                c => $"`{c}` = CASE WHEN `{c}` = @mgr THEN NULL ELSE `{c}` END");
            var sql =
                $"UPDATE `분장표준처리` SET {string.Join(", ", setClauses)} " +
                $"WHERE `항목명` BETWEEN @start AND @end";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@mgr",   managerName);
            cmd.Parameters.AddWithValue("@start", startDate.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@end",   endDate.ToString("yyyy-MM-dd"));
            cmd.ExecuteNonQuery();

            Log($"ClearAssignmentsForAgent: {employeeId}({managerName}) " +
                $"{startDate:yyyy-MM-dd}~{endDate:yyyy-MM-dd}");
        }
        catch (Exception ex) { Log($"ClearAssignmentsForAgent 오류: {ex.Message}"); }
    }

    //  분장표준처리 — 업무 분장 추가 (기간별)
    //
    //  지정된 기간 동안 특정 분석항목을 특정 담당자로 설정
    // =====================================================================
    public static void AddAssignment(string employeeId, string analyte, DateTime startDate, DateTime endDate)
    {

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분석항목 컬럼명 확인 (전체명)
            var info = GetStandardDaysInfo();
            string? columnName = null;
            foreach (var kv in info)
            {
                if (kv.Value.shortName == analyte || kv.Key == analyte)
                {
                    columnName = kv.Key;
                    break;
                }
            }
            if (string.IsNullOrEmpty(columnName))
            {
                Log($"AddAssignment: 분석항목 '{analyte}' 컬럼을 찾을 수 없음");
                return;
            }

            // 담당자 이름 조회 (사번으로)
            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var result = cmd2.ExecuteScalar();
                if (result != null)
                    managerName = result.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName))
            {
                Log($"AddAssignment: 사번 '{employeeId}'의 이름을 찾을 수 없음");
                return;
            }

            // 기간 내 각 날짜에 대해 처리
            for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
            {
                string dateKey = date.ToString("yyyy-MM-dd");

                // 해당 날짜 행 존재 확인
                bool rowExists = false;
                using (var chk = conn.CreateCommand())
                {
                    chk.CommandText = "SELECT COUNT(*) FROM `분장표준처리` WHERE `항목명` = @date";
                    chk.Parameters.AddWithValue("@date", dateKey);
                    rowExists = Convert.ToInt64(chk.ExecuteScalar()!) > 0;
                }

                if (rowExists)
                {
                    // 업데이트
                    using var upd = conn.CreateCommand();
                    upd.CommandText = $"UPDATE `분장표준처리` SET `{columnName}` = @manager WHERE `항목명` = @date";
                    upd.Parameters.AddWithValue("@manager", managerName);
                    upd.Parameters.AddWithValue("@date", dateKey);
                    upd.ExecuteNonQuery();
                }
                else
                {
                    // 삽입 (새 행 생성)
                    // 먼저 컬럼 목록 가져오기
                    var columns = DbConnectionFactory.GetColumnNames(conn, "분장표준처리");

                    // INSERT 문 생성
                    var insertColumns = new List<string> { "`항목명`" };
                    var values = new List<string> { "@date" };
                    var parameters = new Dictionary<string, object> { { "@date", dateKey } };

                    foreach (var col in columns.Skip(1)) // 항목명 제외
                    {
                        insertColumns.Add($"`{col}`");
                        if (col == columnName)
                        {
                            values.Add("@manager");
                            parameters["@manager"] = managerName;
                        }
                        else
                        {
                            values.Add("NULL");
                        }
                    }

                    using var ins = conn.CreateCommand();
                    ins.CommandText = $"INSERT INTO `분장표준처리` ({string.Join(", ", insertColumns)}) VALUES ({string.Join(", ", values)})";
                    foreach (var param in parameters)
                        ins.Parameters.AddWithValue(param.Key, param.Value);
                    ins.ExecuteNonQuery();
                }
            }

            Log($"AddAssignment: {employeeId}({managerName})에게 {analyte}({columnName}) 분장 추가 ({startDate:yyyy-MM-dd} ~ {endDate:yyyy-MM-dd})");
        }
        catch (Exception ex) { Log($"AddAssignment 오류: {ex.Message}"); }
    }

    // =====================================================================
    //  타임라인 — 에이전트 분장 달력 (analyte → 날짜 리스트)
    // =====================================================================
    public record AssignmentCalendarEntry(string FullName, string ShortName, List<DateTime> Dates);

    public static List<AssignmentCalendarEntry> GetAssignmentCalendar(
        string employeeId, DateTime start, DateTime end)
    {
        var result = new Dictionary<string, AssignmentCalendarEntry>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            string managerName = "";
            using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = "SELECT `성명` FROM `Agent` WHERE `사번` = @id";
                cmd2.Parameters.AddWithValue("@id", employeeId);
                var r = cmd2.ExecuteScalar();
                if (r != null) managerName = r.ToString()?.Trim() ?? "";
            }
            if (string.IsNullOrEmpty(managerName)) return [];

            var info = GetStandardDaysInfo();

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", start.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", end.ToString("yyyy-MM-dd"));

            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string dateStr = rdr["항목명"]?.ToString()?.Trim() ?? "";
                if (!DateTime.TryParse(dateStr, out var dt)) continue;

                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
                    if (colName.StartsWith("_") || colName == "항목명") continue;
                    string mgr     = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (mgr != managerName) continue;

                    string shortName = info.TryGetValue(colName, out var meta) ? meta.shortName : colName;
                    if (!result.TryGetValue(colName, out var entry))
                    {
                        entry = new AssignmentCalendarEntry(colName, shortName, []);
                        result[colName] = entry;
                    }
                    entry.Dates.Add(dt);
                }
            }
            Log($"GetAssignmentCalendar({employeeId}, {start:yyyy-MM-dd}~{end:yyyy-MM-dd}): {result.Count}개 항목");
        }
        catch (Exception ex) { Log($"GetAssignmentCalendar 오류: {ex.Message}"); }

        return [.. result.Values];
    }

    // =====================================================================
    //  업무분장표 — 전체 분석항목의 담당자 변경 구간(span) 조회
    //
    //  반환: 항목별로 담당자가 연속된 날짜 구간 리스트
    //    예: (BOD, "정준하", 2025-07-01, 2025-07-15), (BOD, "김지은", 2025-07-16, 2025-07-31)
    // =====================================================================
    public record AssignmentSpan(string FullName, string ShortName, string Manager, DateTime Start, DateTime End);

    public static List<AssignmentSpan> GetAssignmentChartData(DateTime rangeStart, DateTime rangeEnd)
    {
        var spans = new List<AssignmentSpan>();

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "분장표준처리")) return spans;

            var info = GetStandardDaysInfo();

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", rangeStart.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", rangeEnd.ToString("yyyy-MM-dd"));

            var dateRows = new List<(DateTime Date, Dictionary<string, string> Managers)>();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string dateStr = rdr["항목명"]?.ToString()?.Trim() ?? "";
                if (!DateTime.TryParse(dateStr, out var dt)) continue;

                var managers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string col = rdr.GetName(i).Trim();
                    if (col.StartsWith("_") || col == "항목명") continue;
                    string mgr = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    managers[col] = mgr;
                }
                dateRows.Add((dt, managers));
            }

            // 항목별 연속 구간 생성
            foreach (var kv in info)
            {
                string col = kv.Key;
                string shortName = kv.Value.shortName;
                if (col.StartsWith("_") || col == "기타업무" || col == "담당계약업체" || col == "항목명") continue;

                string currentMgr = "";
                DateTime spanStart = rangeStart;

                foreach (var (date, managers) in dateRows)
                {
                    string mgr = managers.TryGetValue(col, out var m) ? m : "";
                    if (mgr != currentMgr)
                    {
                        if (!string.IsNullOrEmpty(currentMgr))
                            spans.Add(new AssignmentSpan(col, shortName, currentMgr, spanStart, date.AddDays(-1)));
                        currentMgr = mgr;
                        spanStart = date;
                    }
                }
                // 마지막 구간 닫기
                if (!string.IsNullOrEmpty(currentMgr) && dateRows.Count > 0)
                    spans.Add(new AssignmentSpan(col, shortName, currentMgr, spanStart, dateRows[^1].Date));
            }

            Log($"GetAssignmentChartData({rangeStart:yyyy-MM-dd}~{rangeEnd:yyyy-MM-dd}): {spans.Count}개 구간");
        }
        catch (Exception ex) { Log($"GetAssignmentChartData 오류: {ex.Message}"); }

        return spans;
    }

    // =====================================================================
    //  업무분장표 — 담당자 이름 직접 지정으로 분장 업데이트
    //  드래그로 경계를 이동할 때 사용
    // =====================================================================
    public static void UpdateAssignmentByName(string analyteFullName, string managerName, DateTime start, DateTime end)
    {
        if (analyteFullName == "항목명" || analyteFullName.StartsWith("_")) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            UpdateAssignmentByNameCore(conn, null, analyteFullName, managerName, start, end, null);
            Log($"UpdateAssignmentByName: {analyteFullName} → {managerName} ({start:yyyy-MM-dd}~{end:yyyy-MM-dd})");
        }
        catch (Exception ex) { Log($"UpdateAssignmentByName 오류: {ex.Message}"); }
    }

    /// <summary>배치용 오버로드 — 호출자가 연결/트랜잭션/컬럼캐시를 공유. 반영 버튼처럼 다량 호출 시 사용.</summary>
    public static void UpdateAssignmentByName(DbConnection conn, DbTransaction tx, string analyteFullName, string managerName, DateTime start, DateTime end, List<string>? cachedColumns = null)
    {
        if (analyteFullName == "항목명" || analyteFullName.StartsWith("_")) return;
        UpdateAssignmentByNameCore(conn, tx, analyteFullName, managerName, start, end, cachedColumns);
    }

    private static void UpdateAssignmentByNameCore(DbConnection conn, DbTransaction? tx, string analyteFullName, string managerName, DateTime start, DateTime end, List<string>? cachedColumns)
    {
        List<string>? columnsForInsert = cachedColumns;
        for (DateTime date = start; date <= end; date = date.AddDays(1))
        {
            string dateKey = date.ToString("yyyy-MM-dd");

            bool rowExists;
            using (var chk = conn.CreateCommand())
            {
                if (tx != null) chk.Transaction = tx;
                chk.CommandText = "SELECT COUNT(*) FROM `분장표준처리` WHERE `항목명` = @date";
                var p = chk.CreateParameter(); p.ParameterName = "@date"; p.Value = dateKey; chk.Parameters.Add(p);
                rowExists = Convert.ToInt64(chk.ExecuteScalar()!) > 0;
            }

            if (rowExists)
            {
                using var upd = conn.CreateCommand();
                if (tx != null) upd.Transaction = tx;
                upd.CommandText = $"UPDATE `분장표준처리` SET `{analyteFullName}` = @mgr WHERE `항목명` = @date";
                var pm = upd.CreateParameter(); pm.ParameterName = "@mgr"; pm.Value = managerName; upd.Parameters.Add(pm);
                var pd = upd.CreateParameter(); pd.ParameterName = "@date"; pd.Value = dateKey; upd.Parameters.Add(pd);
                upd.ExecuteNonQuery();
            }
            else
            {
                columnsForInsert ??= DbConnectionFactory.GetColumnNames(conn, "분장표준처리");
                var insertCols = new List<string> { "`항목명`" };
                var vals = new List<string> { "@date" };
                foreach (var c in columnsForInsert)
                {
                    if (c == "항목명" || c.StartsWith("_")) continue;
                    insertCols.Add($"`{c}`");
                    if (string.Equals(c, analyteFullName, StringComparison.OrdinalIgnoreCase))
                        vals.Add("@mgr");
                    else
                        vals.Add("NULL");
                }
                using var ins = conn.CreateCommand();
                if (tx != null) ins.Transaction = tx;
                ins.CommandText = $"INSERT INTO `분장표준처리` ({string.Join(", ", insertCols)}) VALUES ({string.Join(", ", vals)})";
                var pd = ins.CreateParameter(); pd.ParameterName = "@date"; pd.Value = dateKey; ins.Parameters.Add(pd);
                var pm = ins.CreateParameter(); pm.ParameterName = "@mgr"; pm.Value = managerName; ins.Parameters.Add(pm);
                ins.ExecuteNonQuery();
            }
        }
    }

    /// <summary>분장표에서 특정 항목의 날짜 범위를 NULL로 지움 (축소 시 사용)</summary>
    public static void ClearAssignmentByName(string analyteFullName, DateTime start, DateTime end)
    {
        if (analyteFullName == "항목명" || analyteFullName.StartsWith("_")) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            ClearAssignmentByNameCore(conn, null, analyteFullName, start, end);
            Log($"ClearAssignmentByName: {analyteFullName} ({start:yyyy-MM-dd}~{end:yyyy-MM-dd}) → NULL");
        }
        catch (Exception ex) { Log($"ClearAssignmentByName 오류: {ex.Message}"); }
    }

    /// <summary>배치용 오버로드 — 호출자가 연결/트랜잭션을 공유</summary>
    public static void ClearAssignmentByName(DbConnection conn, DbTransaction tx, string analyteFullName, DateTime start, DateTime end)
    {
        if (analyteFullName == "항목명" || analyteFullName.StartsWith("_")) return;
        ClearAssignmentByNameCore(conn, tx, analyteFullName, start, end);
    }

    private static void ClearAssignmentByNameCore(DbConnection conn, DbTransaction? tx, string analyteFullName, DateTime start, DateTime end)
    {
        for (DateTime date = start; date <= end; date = date.AddDays(1))
        {
            string dateKey = date.ToString("yyyy-MM-dd");
            using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = $"UPDATE `분장표준처리` SET `{analyteFullName}` = NULL WHERE `항목명` = @date";
            var pd = cmd.CreateParameter(); pd.ParameterName = "@date"; pd.Value = dateKey; cmd.Parameters.Add(pd);
            cmd.ExecuteNonQuery();
        }
    }

    /// <summary>분장표에서 항목명이 날짜가 아닌 깨진 행을 삭제</summary>
    public static int CleanupCorruptedRows()
    {
        int deleted = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            // 날짜 패턴(yyyy-MM-dd)이 아닌 행 삭제
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `분장표준처리` WHERE `항목명` NOT REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'";
            deleted = cmd.ExecuteNonQuery();
            if (deleted > 0)
                Log($"CleanupCorruptedRows: {deleted}개 깨진 행 삭제");
        }
        catch (Exception ex) { Log($"CleanupCorruptedRows 오류: {ex.Message}"); }
        return deleted;
    }

    /// <summary>특정 날짜 이후 행을 모두 삭제하고, 마지막 담당자로 재생성</summary>
    /// <param name="progress">진행률 콜백 (0.0~1.0, 항목명)</param>
    public static void RebuildAssignmentsFrom(DateTime fromDate, Action<double, string>? progress = null)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 1) fromDate 이후 행 전부 삭제
            progress?.Invoke(0, "기존 데이터 삭제 중...");
            using (var del = conn.CreateCommand())
            {
                del.CommandText = "DELETE FROM `분장표준처리` WHERE `항목명` >= @d";
                del.Parameters.AddWithValue("@d", fromDate.ToString("yyyy-MM-dd"));
                int cnt = del.ExecuteNonQuery();
                Log($"RebuildAssignmentsFrom: {fromDate:yyyy-MM-dd} 이후 {cnt}개 행 삭제");
            }

            // 2) fromDate 직전일의 담당자 정보 가져오기
            var prevDate = fromDate.AddDays(-1);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `분장표준처리` WHERE `항목명` = @d";
            cmd.Parameters.AddWithValue("@d", prevDate.ToString("yyyy-MM-dd"));
            using var rdr = cmd.ExecuteReader();

            var lastManagers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (rdr.Read())
            {
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string col = rdr.GetName(i).Trim();
                    if (col == "항목명" || col.StartsWith("_")) continue;
                    string val = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(val))
                        lastManagers[col] = val;
                }
            }
            rdr.Close();

            // 3) fromDate ~ 3개월 후 말까지 동일 담당자로 채우기 (4월→7월말)
            var yearEnd = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(4).AddDays(-1);
            var items = lastManagers
                .Where(kv => kv.Key != "기타업무" && kv.Key != "담당계약업체" && kv.Key != "항목명")
                .ToList();
            for (int i = 0; i < items.Count; i++)
            {
                var kv = items[i];
                progress?.Invoke((double)(i + 1) / items.Count, kv.Key);
                UpdateAssignmentByName(kv.Key, kv.Value, fromDate, yearEnd);
            }
            Log($"RebuildAssignmentsFrom: {fromDate:yyyy-MM-dd}~{yearEnd:yyyy-MM-dd} {lastManagers.Count}개 항목 재생성 완료");
        }
        catch (Exception ex) { Log($"RebuildAssignmentsFrom 오류: {ex.Message}"); }
    }

    // =====================================================================
    //  업무분장표 — 오늘까지 미배정 구간을 이전 담당자로 자동 연장
    //
    //  앱 시작 시 서버 연결 후 1회 실행.
    //  각 분석항목의 마지막 배정 이후 ~ 오늘 사이 빈 구간을
    //  직전 담당자로 채워 "오늘 미배정 항목 없음"을 보장.
    // =====================================================================
    /// <summary>
    /// 앱 시작 시 6개월 후까지 빈 날짜 행만 추가 (인원 배정은 하지 않음).
    /// 최초 1회: 오늘 이후 기존 배정 데이터를 NULL로 초기화.
    /// </summary>
    public static void AutoExtendAssignmentsToToday(Action<double, string>? progress = null)
    {

        progress?.Invoke(0, "데이터 정리 중...");
        CleanupCorruptedRows();

        var today = DateTime.Today;
        var extendTo = new DateTime(today.Year, today.Month, 1).AddMonths(7).AddDays(-1); // 6개월 후 말까지

        try
        {
            progress?.Invoke(0.1, "날짜 확인 중...");

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 이미 존재하는 날짜 행 조회
            var existingDates = new HashSet<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT `항목명` FROM `분장표준처리` WHERE `항목명` REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    existingDates.Add(reader.GetString(0));
            }

            // 오늘부터 extendTo까지 없는 날짜만 수집
            var missingDates = new List<string>();
            for (var d = today; d <= extendTo; d = d.AddDays(1))
            {
                var dk = d.ToString("yyyy-MM-dd");
                if (!existingDates.Contains(dk))
                    missingDates.Add(dk);
            }

            if (missingDates.Count == 0)
            {
                Log("AutoExtend: 추가할 날짜 없음");
                progress?.Invoke(1, "완료!");
                return;
            }

            // 빈 날짜 행 INSERT (트랜잭션)
            progress?.Invoke(0.3, $"날짜 행 {missingDates.Count}개 생성 중...");
            using var tx = conn.BeginTransaction();
            for (int i = 0; i < missingDates.Count; i++)
            {
                using var ins = conn.CreateCommand();
                ins.Transaction = tx;
                ins.CommandText = "INSERT INTO `분장표준처리` (`항목명`) VALUES (@d)";
                ins.Parameters.AddWithValue("@d", missingDates[i]);
                ins.ExecuteNonQuery();

                if (i % 20 == 0)
                    progress?.Invoke(0.3 + 0.65 * (i + 1) / missingDates.Count, missingDates[i]);
            }
            tx.Commit();

            Log($"AutoExtend: {missingDates.Count}개 날짜 행 추가 ({missingDates[0]}~{missingDates[^1]})");
            progress?.Invoke(1, "완료!");
        }
        catch (Exception ex) { Log($"AutoExtendAssignmentsToToday 오류: {ex.Message}"); }
    }

    // =====================================================================
    //  당월 분석항목별 의뢰 건수 — 시약 적정사용량 산정용
    // =====================================================================
    /// <summary>이번달 1일~오늘까지 특정 분석항목의 의뢰 건수 (값이 'O'인 행 수)</summary>
    public static int GetMonthlyAnalyteCount(string analyteName)
    {
        if (string.IsNullOrEmpty(analyteName)) return 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return 0;
            if (!DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", analyteName)) return 0;

            var today = DateTime.Today;
            var monthStart = new DateTime(today.Year, today.Month, 1).ToString("yyyy-MM-dd");
            var todayStr   = today.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT COUNT(*) FROM `수질분석센터_결과`
                WHERE `{analyteName}` = 'O'
                  AND `채취일자` >= @start AND `채취일자` <= @end";
            cmd.Parameters.AddWithValue("@start", monthStart);
            cmd.Parameters.AddWithValue("@end",   todayStr);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }
        catch (Exception ex) { Log($"GetMonthlyAnalyteCount 오류: {ex.Message}"); return 0; }
    }

    // =====================================================================
    //  채수담당자 등록 — 시료채취1/채수담당자/시료채취자1 컬럼 중 존재하는 것에 저장
    // =====================================================================
    public static void UpdateSamplers(int rowId, IEnumerable<string> names)
    {
        var nameStr = string.Join(", ", names);

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 여러 후보 컬럼 중 존재하는 첫 번째 컬럼에 저장
            var candidates = new[] { "시료채취1", "채수담당자", "시료채취자1", "시료채취자-1" };
            string? col = candidates.FirstOrDefault(c => DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", c));
            if (col == null) { Log($"UpdateSamplers: 채수담당자 컬럼 없음 (rowId={rowId})"); return; }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `수질분석센터_결과` SET `{col}` = @v WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@v", nameStr);
            cmd.Parameters.AddWithValue("@id", rowId);
            cmd.ExecuteNonQuery();
            Log($"UpdateSamplers: rowId={rowId} {col}='{nameStr}'");
        }
        catch (Exception ex) { Log($"UpdateSamplers 오류: {ex.Message}"); }
    }

    // =====================================================================
    //  날짜별 레코드 조회 (분석결과입력 자동분류용)
    // =====================================================================
    public static List<AnalysisRequestRecord> GetByDate(string date)
    {
        var list = new List<AnalysisRequestRecord>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`, ''),
                       COALESCE(`시료명`, ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, '')
                FROM `수질분석센터_결과`
                WHERE `채취일자` = @d
                ORDER BY {DbConnectionFactory.RowId}";
            cmd.Parameters.AddWithValue("@d", date);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new AnalysisRequestRecord
                {
                    Id       = Convert.ToInt32(rdr.GetValue(0)),
                    약칭     = rdr.GetString(1),
                    시료명   = rdr.GetString(2),
                    접수번호 = rdr.GetString(3),
                    채취일자 = rdr.GetString(4),
                });
            }
        }
        catch (Exception ex) { Log($"GetByDate 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  월별 의뢰 날짜 목록 (분석결과입력 트리뷰 용 — 수질분석센터 모드)
    // =====================================================================
    public static List<string> GetMonths()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT SUBSTR(`채취일자`, 1, 7) AS ym
                FROM `수질분석센터_결과`
                WHERE `채취일자` IS NOT NULL AND `채취일자` <> ''
                ORDER BY ym DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (r.IsDBNull(0)) continue;
                var ym = r.GetString(0);
                if (!string.IsNullOrWhiteSpace(ym)) list.Add(ym);
            }
        }
        catch (Exception ex) { Log($"AnalysisRequestService.GetMonths 오류: {ex.Message}"); }
        return list;
    }

    public static List<string> GetDatesByMonth(string yearMonth)
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `채취일자`
                FROM `수질분석센터_결과`
                WHERE SUBSTR(`채취일자`, 1, 7) = @ym
                ORDER BY `채취일자` DESC";
            cmd.Parameters.AddWithValue("@ym", yearMonth);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                if (rdr.IsDBNull(0)) continue;
                var d = rdr.GetString(0);
                if (!string.IsNullOrWhiteSpace(d)) list.Add(d);
            }
        }
        catch (Exception ex) { Log($"GetDatesByMonth 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  특정 일자에 의뢰된 분석항목 상태 (분석결과입력 배지용)
    //  → 수질분석센터_결과의 각 항목 컬럼을 스캔하여
    //     (항목명 → 결과입력여부) 맵 리턴
    //  - 'O' 값만 있으면 = 의뢰됐으나 결과 미입력 (requested=true, filled=false)
    //  - 숫자/문자 값이면 = 결과 입력됨            (requested=true, filled=true)
    //  - NULL 이면 = 의뢰 없음 (맵에 포함 안됨)
    // =====================================================================
    public static Dictionary<string, bool> GetRequestedItemsByDate(string date)
    {
        var result = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return result;

            // 항목 컬럼 목록 = 분석정보 Analyte
            var analyteCols = new List<string>();
            using (var c = conn.CreateCommand())
            {
                c.CommandText = "SELECT DISTINCT Analyte FROM `분석정보`";
                using var r = c.ExecuteReader();
                while (r.Read())
                {
                    var name = r.IsDBNull(0) ? "" : r.GetString(0);
                    if (!string.IsNullOrWhiteSpace(name) &&
                        DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", name))
                        analyteCols.Add(name);
                }
            }
            if (analyteCols.Count == 0) return result;

            using var cmd = conn.CreateCommand();
            var colList = string.Join(", ", analyteCols.Select(c => $"`{c}`"));
            cmd.CommandText = $"SELECT {colList} FROM `수질분석센터_결과` WHERE `채취일자` = @d";
            cmd.Parameters.AddWithValue("@d", date);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 0; i < analyteCols.Count; i++)
                {
                    if (rdr.IsDBNull(i)) continue;
                    var v = rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (v.Length == 0) continue;

                    var key = analyteCols[i];
                    // 'O' 는 의뢰만 있는 상태, 그 외 값은 결과 입력됨
                    bool filled = !v.Equals("O", StringComparison.OrdinalIgnoreCase);
                    // 같은 항목이 여러 시료에 걸쳐 있을 때 filled=true 를 유지 (OR)
                    if (result.TryGetValue(key, out var prev))
                        result[key] = prev || filled;
                    else
                        result[key] = filled;
                }
            }
        }
        catch (Exception ex) { Log($"GetRequestedItemsByDate 오류: {ex.Message}"); }
        return result;
    }

    // =====================================================================
    //  날짜별 의뢰 카테고리 상태 (수질분석센터 뱃지용)
    //  → 수질분석센터_결과의 항목 컬럼을 분석정보.Category 로 그룹핑
    //     (Category명 → 결과입력여부) 맵 리턴
    //  - 'O' 만 있으면 = 의뢰됐으나 결과 미입력
    //  - 숫자/문자 값이면 = 결과 입력됨 (한 카테고리 내 어느 항목이라도 입력됐으면 filled=true)
    //  - 어떤 항목도 의뢰 안 됐으면 = 맵에 포함 안됨
    // =====================================================================
    public static Dictionary<string, bool> GetRequestedCategoriesByDate(string date)
    {
        var result = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return result;
            if (!DbConnectionFactory.TableExists(conn, "분석정보")) return result;

            // Analyte → Category 매핑
            var analyteToCategory = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var c = conn.CreateCommand())
            {
                c.CommandText = "SELECT Analyte, COALESCE(Category, '') FROM `분석정보`";
                using var r = c.ExecuteReader();
                while (r.Read())
                {
                    var name = r.IsDBNull(0) ? "" : r.GetString(0);
                    var cat  = r.IsDBNull(1) ? "" : r.GetString(1);
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(cat)) continue;
                    analyteToCategory[name.Trim()] = cat.Trim();
                }
            }
            if (analyteToCategory.Count == 0) return result;

            // 수질분석센터_결과에 실제로 존재하는 컬럼만 필터
            var validCols = analyteToCategory.Keys
                .Where(name => DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", name))
                .ToList();
            if (validCols.Count == 0) return result;

            using var cmd = conn.CreateCommand();
            var colList = string.Join(", ", validCols.Select(c => $"`{c}`"));
            cmd.CommandText = $"SELECT {colList} FROM `수질분석센터_결과` WHERE `채취일자` = @d";
            cmd.Parameters.AddWithValue("@d", date);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                for (int i = 0; i < validCols.Count; i++)
                {
                    if (rdr.IsDBNull(i)) continue;
                    var v = rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (v.Length == 0) continue;

                    var category = analyteToCategory[validCols[i]];
                    bool filled = !v.Equals("O", StringComparison.OrdinalIgnoreCase);
                    if (result.TryGetValue(category, out var prev))
                        result[category] = prev || filled;
                    else
                        result[category] = filled;
                }
            }
        }
        catch (Exception ex) { Log($"GetRequestedCategoriesByDate 오류: {ex.Message}"); }
        return result;
    }

    // =====================================================================
    //  특정 월(YYYY-MM)의 의뢰 목록 — Show1 월별 lazy load 용
    // =====================================================================
    public static List<AnalysisRequestRecord> GetRecordsByMonth(string ym)
    {
        var list = new List<AnalysisRequestRecord>();
        if (string.IsNullOrWhiteSpace(ym) || ym.Length < 7) return list;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`, ''),
                       COALESCE(`시료명`, ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, '')
                FROM `수질분석센터_결과`
                WHERE `채취일자` LIKE @prefix
                ORDER BY `채취일자` DESC, {DbConnectionFactory.RowId} DESC";
            cmd.Parameters.AddWithValue("@prefix", ym.Substring(0, 7) + "%");
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new AnalysisRequestRecord
                {
                    Id       = Convert.ToInt32(rdr.GetValue(0)),
                    약칭     = rdr.GetString(1),
                    시료명   = rdr.GetString(2),
                    접수번호 = rdr.GetString(3),
                    채취일자 = rdr.GetString(4),
                });
            }
        }
        catch (Exception ex) { Log($"GetRecordsByMonth({ym}) 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  최근 N개월 의뢰 목록 조회 (수동 매칭 팝업용)
    // =====================================================================
    public static List<AnalysisRequestRecord> GetRecentRecords(int months = 1)
    {
        var list = new List<AnalysisRequestRecord>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;
            using var cmd = conn.CreateCommand();
            var cutoff = DateTime.Today.AddMonths(-months).ToString("yyyy-MM-dd");
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`, ''),
                       COALESCE(`시료명`, ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, '')
                FROM `수질분석센터_결과`
                WHERE `채취일자` >= @cutoff
                ORDER BY `채취일자` DESC";
            cmd.Parameters.AddWithValue("@cutoff", cutoff);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new AnalysisRequestRecord
                {
                    Id       = Convert.ToInt32(rdr.GetValue(0)),
                    약칭     = rdr.GetString(1),
                    시료명   = rdr.GetString(2),
                    접수번호 = rdr.GetString(3),
                    채취일자 = rdr.GetString(4),
                });
            }
        }
        catch (Exception ex) { Log($"GetRecentRecords 오류: {ex.Message}"); }
        return list;
    }

    /// <summary>생태독성이 포함된 의뢰 목록 (채취일자별 그룹)</summary>
    public static List<(string 채취일자, int Id, string 약칭, string 시료명, string 접수번호, string 결과)> GetEcotoxRecords(int months = 6)
    {
        var list = new List<(string, int, string, string, string, string)>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return list;
            if (!DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", "생태독성")) return list;
            using var cmd = conn.CreateCommand();
            var cutoff = DateTime.Today.AddMonths(-months).ToString("yyyy-MM-dd");
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`, ''),
                       COALESCE(`시료명`, ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, ''),
                       COALESCE(`생태독성`, '')
                FROM `수질분석센터_결과`
                WHERE `채취일자` >= @cutoff
                  AND `생태독성` IS NOT NULL AND `생태독성` <> ''
                ORDER BY `채취일자` DESC, {DbConnectionFactory.RowId}";
            cmd.Parameters.AddWithValue("@cutoff", cutoff);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add((
                    rdr.GetString(4),
                    Convert.ToInt32(rdr.GetValue(0)),
                    rdr.GetString(1),
                    rdr.GetString(2),
                    rdr.GetString(3),
                    rdr.GetString(5)));
            }
        }
        catch (Exception ex) { Log($"GetEcotoxRecords 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  단일 컬럼 결과값 업데이트 (분석결과입력용)
    // =====================================================================
    public static void UpdateResultValue(int rowId, string columnName, string value)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"UPDATE `수질분석센터_결과` SET `{columnName}` = @v WHERE {DbConnectionFactory.RowId} = @id";
            cmd.Parameters.AddWithValue("@v", value);
            cmd.Parameters.AddWithValue("@id", rowId);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Log($"UpdateResultValue 오류: {ex.Message}"); }
    }

    /// <summary>생태독성 저장 후 수질분석센터_결과.생태독성 컬럼에 TU값 업데이트</summary>
    public static void UpdateEcotoxResult(string 채취일자, string 시료명, string tuValue)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return;
            if (!DbConnectionFactory.ColumnExists(conn, "수질분석센터_결과", "생태독성")) return;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE `수질분석센터_결과`
                                   SET `생태독성` = @tu
                                 WHERE LEFT(`채취일자`, 10) = @d
                                   AND `시료명` = @nm";
            cmd.Parameters.AddWithValue("@tu", tuValue);
            cmd.Parameters.AddWithValue("@d",  채취일자);
            cmd.Parameters.AddWithValue("@nm", 시료명);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Log($"UpdateEcotoxResult 오류: {ex.Message}"); }
    }
}
