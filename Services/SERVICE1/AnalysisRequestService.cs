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
/// "분석의뢰및결과" 테이블 접근 서비스
/// TestReportService와 동일한 DbPathHelper.DbPath 사용
/// </summary>
public static class AnalysisRequestService
{

    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Treeview.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AnalysisRequestService] {msg}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    // =====================================================================
    //  전체 레코드 목록 (트리뷰용)
    // =====================================================================
    public static List<AnalysisRequestRecord> GetAllRecords()
    {
        var list = new List<AnalysisRequestRecord>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Log("DB 파일 없음"); return list; }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        if (!DbConnectionFactory.TableExists(conn, "분석의뢰및결과"))
        {
            Log("테이블 없음: 분석의뢰및결과");
            return list;
        }

        // MariaDB: using var 중첩 시 커넥션 충돌 방지 — 단일 using 블록
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $@"
                SELECT {DbConnectionFactory.RowId},
                       COALESCE(`약칭`,     ''),
                       COALESCE(`시료명`,   ''),
                       COALESCE(`견적번호`, ''),
                       COALESCE(`채취일자`, '')
                FROM   `분석의뢰및결과`
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
                        Id       = Convert.ToInt32(rdr.GetValue(0)),
                        약칭     = rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                        시료명   = rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                        접수번호 = rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                        의뢰일   = date,
                        연도     = 연도,
                        월       = 월,
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return dict;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM `분석의뢰및결과` WHERE {DbConnectionFactory.RowId} = @id";
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return false;
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
                $"UPDATE `분석의뢰및결과` SET {string.Join(", ", setParts)} " +
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
    //  같은 견적번호를 가진 분석의뢰및결과 행들에서 각 항목이 'O'인 횟수를 집계해
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
            c.CommandText = $"SELECT COALESCE(`견적번호`, '') FROM `분석의뢰및결과` WHERE {DbConnectionFactory.RowId} = @id";
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
            c.CommandText = "SELECT * FROM `분석의뢰및결과` WHERE `견적번호` = @no";
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return false;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"DELETE FROM `분석의뢰및결과` WHERE {DbConnectionFactory.RowId} = @id";
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;
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

    /// <summary>분析의뢰및결과의 방류허용기준 적용유무 업데이트</summary>
    public static bool UpdateDischargeStandard(int rowId, string standardName)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return false;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"UPDATE `분석의뢰및결과` SET `방류허용기준 적용유무` = @val " +
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `시료명`
                FROM `분석의뢰및결과`
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return 0;
        if (string.IsNullOrWhiteSpace(newName)) return 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                UPDATE `분석의뢰및결과`
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return list;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT DISTINCT `약칭`
                FROM `분석의뢰및결과`
                WHERE `약칭` IS NOT NULL AND `약칭` <> ''
                ORDER BY `약칭` ASC";
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(r.GetString(0));
        }
        catch (Exception ex) { Log($"GetDistinctAbbreviations 오류: {ex.Message}"); }
        return list;
    }

    // =====================================================================
    //  명칭 정리 — 분장표준처리 — 약칭 직접 조회
    //    항목명='약칭' 행을 찾아 컬럼헤드 → 약칭 매핑 반환
    //    예: "생물화학적 산소요구량" → "BOD", "수소이온농도" → "pH"
    // =====================================================================
    public static Dictionary<string, string> GetShortNames()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;
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
                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colHeader = rdr.GetName(i).Trim();
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) { Log("GetStandardDaysInfo: DB 없음"); return result; }

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            if (!DbConnectionFactory.TableExists(conn, "분장표준처리"))
            { Log("GetStandardDaysInfo: 분장표준처리 테이블 없음"); return result; }

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
                string label = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString()?.Trim() ?? "";

                if (!daysSet && label == "표준처리기한")
                {
                    for (int i = 1; i < fc; i++)
                    {
                        var v = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                        int.TryParse(v, out daysArr[i]);
                    }
                    daysSet = true;
                    continue;
                }

                if (!shortSet && label == "약칭")
                {
                    for (int i = 1; i < fc; i++)
                        shortArr[i] = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    shortSet = true;
                    continue;
                }

                if (daysSet && shortSet) break;
            }

            for (int i = 1; i < fc; i++)
            {
                string fullName = headers[i];
                if (string.IsNullOrEmpty(fullName)) continue;
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 날짜 정규화: yyyy-MM-dd
            string dateKey = sampleDate.Trim();
            if (DateTime.TryParse(dateKey, out var dt))
                dateKey = dt.ToString("yyyy-MM-dd");

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` = @date LIMIT 1";
            cmd.Parameters.AddWithValue("@date", dateKey);

            using var rdr = cmd.ExecuteReader();
            if (rdr.Read())
            {
                for (int i = 1; i < rdr.FieldCount; i++) // col0 = 항목명 건너뜀
                {
                    string colName = rdr.GetName(i).Trim();
                    string manager = rdr.IsDBNull(i) ? "" : rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(colName))
                        result[colName] = manager;
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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignments;

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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return result;

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
    //  분장표준처리 — 기간 내 특정 항목을 담당하는 직원 목록 (중복 제거)
    // =====================================================================
    /// <summary>
    /// 기간 내 특정 항목 담당자 목록 — 날짜별 중복 포함.
    /// 같은 사람이 여러 날 담당하면 여러 번 포함됩니다.
    /// </summary>
    public static List<string> GetAssigneesForAnalyteInRange(string analyteFullName, DateTime start, DateTime end)
    {
        var assignees = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignees;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT * FROM `분장표준처리` WHERE `항목명` BETWEEN @s AND @e ORDER BY `항목명`";
            cmd.Parameters.AddWithValue("@s", start.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@e", end.ToString("yyyy-MM-dd"));
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
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
        catch (Exception ex) { Log($"GetAssigneesForAnalyteInRange 오류: {ex.Message}"); }
        return assignees;
    }

    // =====================================================================
    //  분장표준처리 — 특정 날짜에 특정 항목을 담당하는 직원 목록
    // =====================================================================
    public static List<string> GetAssigneesForAnalyteOnDate(string analyteFullName, DateTime date)
    {
        var assignees = new List<string>();
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return assignees;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            using var cmd = conn.CreateCommand();
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
        catch (Exception ex) { Log($"GetAssigneesForAnalyteOnDate 오류: {ex.Message}"); }

        return assignees;
    }

    // =====================================================================
    // =====================================================================
    //  분장표준처리 — 특정 직원의 기간 내 모든 분장을 NULL로 초기화
    // =====================================================================
    public static void ClearAssignmentsForAgent(string employeeId, DateTime startDate, DateTime endDate)
    {
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return;

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
        if (!DbConnectionFactory.IsMariaDb && !File.Exists(DbPathHelper.DbPath)) return [];

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
                string dateStr = rdr.IsDBNull(0) ? "" : rdr.GetValue(0)?.ToString() ?? "";
                if (!DateTime.TryParse(dateStr, out var dt)) continue;

                for (int i = 1; i < rdr.FieldCount; i++)
                {
                    string colName = rdr.GetName(i).Trim();
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
}
