using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE2;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 시험기록부 Excel 생성 — 스캐폴드 단계.
/// 선택 시료의 각 분석항목별 {Analyte}_시험기록부 테이블 raw 레코드를
/// "자료" 시트에 덤프하고, 공유 포맷 시트(검정곡선/SS/BOD/생태독성) placeholder 생성.
/// 실제 포맷 채움 로직은 후속 단계.
/// </summary>
public static class TestRecordBookService
{
    public static string GetOutputDir()
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var dir = Path.Combine(desktop, "시험의뢰서 출력", DateTime.Today.ToString("yyyy-MM-dd"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static bool TemplatesExist() => true;

    // 처리시설/비용부담금용 시험기록부 대상 항목 — BOD, TOC, SS, TN, TP, 대장균, 페놀류, 노말헥산
    // 후보 테이블명은 여러 가능성을 순회하며 존재하는 테이블을 사용.
    // 페놀류는 직접법/추출법 각각 별도 시트로 분리 허용.
    private static readonly (string Label, string[] Candidates, bool AllowMultiple)[] FacilityWasteAnalytes =
    {
        ("BOD",       new[] { "생물학적_산소요구량_시험기록부", "생물화학적_산소요구량_시험기록부" }, false),
        ("TOC",       new[] { "총유기탄소_시험기록부", "TOC_시험기록부", "총_유기탄소_NPOC_시험기록부" }, false),
        ("SS",        new[] { "부유물질_시험기록부", "SS_시험기록부" }, false),
        ("TN",        new[] { "총질소_시험기록부", "TN_시험기록부", "T-N_시험기록부" }, false),
        ("TP",        new[] { "총인_시험기록부", "TP_시험기록부", "T-P_시험기록부" }, false),
        ("대장균",    new[] { "총대장균군_시험기록부", "대장균_시험기록부", "대장균군_시험기록부" }, false),
        ("페놀류",    new[] { "페놀류_직접법_시험기록부", "페놀류_추출법_시험기록부", "페놀류_시험기록부" }, true),
        ("노말헥산",  new[] { "노말헥산추출물질_시험기록부", "노말헥산_추출물질_시험기록부", "N-Hexan_시험기록부" }, false),
    };

    /// <summary>
    /// 처리시설/비용부담금용 시험기록부 생성.
    /// 8개 항목(BOD/TOC/SS/TN/TP/대장균/페놀류/노말헥산)의 `_시험기록부` 테이블을
    /// (분석일 LIKE date%) + (업체명 = companyFilter, optional) 조건으로 조회하여
    /// 항목당 1개 시트로 덤프. 페놀류는 직접법/추출법 각각 별 시트.
    /// </summary>
    /// <param name="date">채취/분석일 (yyyy-MM-dd)</param>
    /// <param name="companyFilter">업체명 필터 (null/빈 문자열이면 전체)</param>
    /// <param name="title">파일명/헤더에 들어갈 라벨 (예: "처리시설_여수", "비용부담금")</param>
    public static string GenerateFacilityWaste(string date, string? companyFilter, string title)
    {
        using var wb = new XLWorkbook();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        int sheetsAdded = 0;
        foreach (var (label, candidates, allowMultiple) in FacilityWasteAnalytes)
        {
            int matchedForLabel = 0;
            foreach (var tableName in candidates)
            {
                if (!DbConnectionFactory.TableExists(conn, tableName)) continue;
                var cols = DbConnectionFactory.GetColumnNames(conn, tableName);
                if (cols.Count == 0) continue;

                var rows = QueryFacilityWaste(conn, tableName, cols, date, companyFilter);

                // 페놀류처럼 여러 테이블 허용 시 각 테이블마다 별 시트
                string sheetName = matchedForLabel == 0
                    ? label
                    : $"{label}_{StripSuffix(tableName)}";
                var ws = wb.Worksheets.Add(SanitizeSheetName(sheetName));

                // 헤더
                for (int i = 0; i < cols.Count; i++)
                    ws.Cell(1, i + 1).Value = cols[i];
                ws.Range(1, 1, 1, cols.Count).Style.Font.Bold = true;

                // 데이터
                int r = 2;
                foreach (var row in rows)
                {
                    for (int c = 0; c < cols.Count && c < row.Count; c++)
                        ws.Cell(r, c + 1).Value = row[c] ?? "";
                    r++;
                }
                if (rows.Count == 0)
                    ws.Cell(2, 1).Value = $"(매칭 레코드 없음: date LIKE '{date}%'"
                        + (string.IsNullOrEmpty(companyFilter) ? "" : $", 업체명='{companyFilter}'") + ")";

                try { ws.Columns().AdjustToContents(); } catch { }
                sheetsAdded++;
                matchedForLabel++;

                if (!allowMultiple) break;
            }
        }

        if (sheetsAdded == 0)
        {
            var ws = wb.Worksheets.Add("빈_데이터");
            ws.Cell(1, 1).Value = $"[{title}] {date} — 대상 테이블 없음";
        }

        var safeTitle = SanitizeForFileName(title);
        var safeDate  = SanitizeForFileName(date);
        var stamp     = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var fileName  = $"시험기록부_{safeTitle}_{safeDate}_{stamp}.xlsx";
        var fullPath  = Path.Combine(GetOutputDir(), fileName);
        wb.SaveAs(fullPath);
        return fullPath;
    }

    private static List<List<string>> QueryFacilityWaste(
        DbConnection conn, string tableName, List<string> cols,
        string date, string? companyFilter)
    {
        var result = new List<List<string>>();
        bool hasDate    = cols.Contains("분석일", StringComparer.OrdinalIgnoreCase);
        bool hasCompany = !string.IsNullOrEmpty(companyFilter)
                      && cols.Contains("업체명", StringComparer.OrdinalIgnoreCase);

        var where = new List<string>();
        using var cmd = conn.CreateCommand();
        if (hasDate)
        {
            where.Add("`분석일` LIKE @d");
            cmd.Parameters.AddWithValue("@d", $"{date}%");
        }
        if (hasCompany)
        {
            where.Add("`업체명` = @c");
            cmd.Parameters.AddWithValue("@c", companyFilter!);
        }
        cmd.CommandText = $"SELECT * FROM `{tableName}`"
            + (where.Count > 0 ? " WHERE " + string.Join(" AND ", where) : "");

        try
        {
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var row = new List<string>();
                for (int i = 0; i < cols.Count && i < r.FieldCount; i++)
                    row.Add(r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "");
                result.Add(row);
            }
        }
        catch { }
        return result;
    }

    private static string StripSuffix(string tableName)
        => tableName.EndsWith("_시험기록부") ? tableName[..^"_시험기록부".Length] : tableName;

    private static string SanitizeSheetName(string name)
    {
        var invalid = new[] { ':', '/', '\\', '?', '*', '[', ']' };
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        return cleaned.Length > 31 ? cleaned[..31] : cleaned;
    }

    // ═════════════════════════════════════════════════════════════════════════
    // BOD 기본항목 시험기록부 생성
    // 템플릿: Data/Templates/기본항목 시험기록부/BOD 시험기록부.xlsx
    //   · BOD 시트 2페이지 (좌: A-H cols, 우: I-P cols), R8~R37 30행/페이지 × 2 = 60 샘플
    //   · 주입 대상: A/I=시료명, B/J=시료량, C/K=D1, D/L=D2, H/P=SN
    //   · E/F/G/M/N/O 수식은 템플릿 기존값 유지 (f(x/y)/P희석배수/Result 자동계산)
    // ═════════════════════════════════════════════════════════════════════════

    private record BodRow(string 시료명, string 시료량, string D1, string D2, string SN);

    /// <summary>
    /// BOD 시험기록부 생성.
    /// dateKind="채수일": 비용부담금_결과.채수일 → JOIN → 시험기록부 (BOD는 분석일=채수일+5일 시차 대응)
    /// dateKind="분석일": 직접 시험기록부.분석일 LIKE 조회
    /// </summary>
    public static string GenerateBodRecordBook(string date, string? companyFilter, string dateKind = "채수일")
    {
        var templateDir = TemplateConfiguration.Resolve("TestRecordBookFolder");
        var templatePath = Path.Combine(templateDir, "BOD 시험기록부.xlsx");
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"BOD 시험기록부 템플릿 없음: {templatePath}");

        string[] candidates = { "생물학적_산소요구량_시험기록부", "생물화학적_산소요구량_시험기록부" };
        List<BodRow> rows;
        using (var conn = DbConnectionFactory.CreateConnection())
        {
            conn.Open();
            string? activeTable = candidates.FirstOrDefault(t => DbConnectionFactory.TableExists(conn, t));
            if (activeTable == null)
                throw new InvalidOperationException(
                    $"BOD 시험기록부 테이블 없음 (후보: {string.Join(", ", candidates)})");

            rows = dateKind == "채수일"
                ? QueryBodRowsByCollection(conn, activeTable, date, companyFilter)
                : QueryBodRows(conn, activeTable, date, companyFilter);

            // 채수일 기반인데 매칭이 0건이면, 분석일 기반으로 fallback
            if (rows.Count == 0 && dateKind == "채수일")
                rows = QueryBodRows(conn, activeTable, date, companyFilter);
        }

        using var wb = new XLWorkbook(templatePath);
        var ws = wb.Worksheet("BOD");

        int truncated = 0;
        for (int i = 0; i < rows.Count; i++)
        {
            int row, nameCol, snCol;
            if (i < 30) { row = 8 + i;        nameCol = 1; snCol = 8;  }
            else if (i < 60) { row = 8 + (i - 30); nameCol = 9; snCol = 16; }
            else { truncated++; continue; }

            var r = rows[i];
            ws.Cell(row, nameCol + 0).Value = r.시료명;
            ws.Cell(row, nameCol + 1).Value = r.시료량;
            ws.Cell(row, nameCol + 2).Value = r.D1;
            ws.Cell(row, nameCol + 3).Value = r.D2;
            ws.Cell(row, snCol).Value       = r.SN;
        }

        if (truncated > 0)
            System.Diagnostics.Debug.WriteLine(
                $"[TestRecordBook] BOD {date}: 60개 초과 {truncated}건 잘림");

        var stamp     = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var titleLbl  = SanitizeForFileName(string.IsNullOrEmpty(companyFilter) ? "전체" : companyFilter);
        var dateLbl   = SanitizeForFileName(date);
        var fileName  = $"시험기록부_BOD_{titleLbl}_{dateLbl}_{stamp}.xlsx";
        var outPath   = Path.Combine(GetOutputDir(), fileName);
        wb.SaveAs(outPath);
        return outPath;
    }

    /// <summary>
    /// 채수일 기반 bridge 조회: 비용부담금_결과에서 (채수일, 업체명)으로 SN 목록을 얻은 뒤
    /// 시험기록부와 (SN, 업체명) JOIN. BOD처럼 분석일 != 채수일 시차 이슈 해결.
    /// </summary>
    private static List<BodRow> QueryBodRowsByCollection(
        DbConnection conn, string tableName, string date, string? companyFilter)
    {
        var list = new List<BodRow>();
        if (!DbConnectionFactory.TableExists(conn, "비용부담금_결과"))
            return list;  // bridge 불가 — 호출부에서 분석일 fallback

        var cols = DbConnectionFactory.GetColumnNames(conn, tableName);
        bool hasCompany = !string.IsNullOrEmpty(companyFilter)
                      && cols.Contains("업체명", StringComparer.OrdinalIgnoreCase);
        string nameCol = cols.Contains("시료명", StringComparer.OrdinalIgnoreCase) ? "시료명" : "업체명";

        var where = new List<string> { "r.`채수일` = @d" };
        using var cmd = conn.CreateCommand();
        cmd.Parameters.AddWithValue("@d", date);
        if (hasCompany)
        {
            where.Add("b.`업체명` = @c");
            cmd.Parameters.AddWithValue("@c", companyFilter!);
        }
        cmd.CommandText =
            $"SELECT b.`{nameCol}`, b.`시료량`, b.`D1`, b.`D2`, b.`SN`"
            + $" FROM `{tableName}` b"
            + $" INNER JOIN `비용부담금_결과` r ON r.`SN` = b.`SN` AND r.`업체명` = b.`업체명`"
            + $" WHERE {string.Join(" AND ", where)}"
            + $" ORDER BY b.`SN`";

        try
        {
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new BodRow(S(r, 0), S(r, 1), S(r, 2), S(r, 3), S(r, 4)));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[QueryBodRowsByCollection] {ex.Message}");
        }
        return list;

        static string S(DbDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
    }

    private static List<BodRow> QueryBodRows(
        DbConnection conn, string tableName, string date, string? companyFilter)
    {
        var list = new List<BodRow>();
        var cols = DbConnectionFactory.GetColumnNames(conn, tableName);

        bool hasCompany = !string.IsNullOrEmpty(companyFilter)
                      && cols.Contains("업체명", StringComparer.OrdinalIgnoreCase);
        string nameCol = cols.Contains("시료명", StringComparer.OrdinalIgnoreCase) ? "시료명" : "업체명";

        var where = new List<string> { "`분석일` LIKE @d" };
        using var cmd = conn.CreateCommand();
        cmd.Parameters.AddWithValue("@d", $"{date}%");
        if (hasCompany)
        {
            where.Add("`업체명` = @c");
            cmd.Parameters.AddWithValue("@c", companyFilter!);
        }
        cmd.CommandText =
            $"SELECT `{nameCol}`, `시료량`, `D1`, `D2`, `SN` FROM `{tableName}`"
            + $" WHERE {string.Join(" AND ", where)} ORDER BY `SN`";

        try
        {
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new BodRow(S(r, 0), S(r, 1), S(r, 2), S(r, 3), S(r, 4)));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[QueryBodRows] {ex.Message}");
        }
        return list;

        static string S(DbDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
    }

    public static string Generate(
        SampleRequest sample,
        List<AnalysisResultRow> rows,
        Dictionary<string, AnalysisItem>? meta = null)
    {
        if (sample == null) throw new ArgumentNullException(nameof(sample));
        rows ??= new List<AnalysisResultRow>();

        using var wb = new XLWorkbook();

        BuildDataSheet(wb, sample, rows);
        BuildPlaceholderSheet(wb, "검정곡선");
        BuildPlaceholderSheet(wb, "SS");
        BuildPlaceholderSheet(wb, "BOD");
        BuildPlaceholderSheet(wb, "생태독성");

        var safeY   = SanitizeForFileName(sample.약칭);
        var safeS   = SanitizeForFileName(sample.시료명);
        var stamp   = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var fileName = $"시험기록부_{safeY}_{safeS}_{stamp}.xlsx";
        var fullPath = Path.Combine(GetOutputDir(), fileName);
        wb.SaveAs(fullPath);
        return fullPath;
    }

    // ── "자료" 시트: 각 항목의 {Analyte}_시험기록부 레코드를 row-dump ───────
    private static void BuildDataSheet(
        XLWorkbook wb,
        SampleRequest sample,
        List<AnalysisResultRow> rows)
    {
        var ws = wb.Worksheets.Add("자료");

        DbConnection? conn = null;
        try { conn = DbConnectionFactory.CreateConnection(); conn.Open(); }
        catch { /* 연결 실패 시 빈 시트 */ }

        ws.Cell(1, 1).Value = "분석항목";
        ws.Cell(1, 2).Value = "테이블";
        int curRow = 2;

        if (conn == null)
        {
            ws.Cell(curRow, 1).Value = "(DB 연결 실패)";
            return;
        }

        foreach (var row in rows)
        {
            var analyte = (row.항목명 ?? "").Trim();
            if (string.IsNullOrEmpty(analyte)) continue;

            var tableName = $"{WaterCenterDbMigration.SafeName(analyte)}_시험기록부";
            if (!DbConnectionFactory.TableExists(conn, tableName))
            {
                ws.Cell(curRow, 1).Value = analyte;
                ws.Cell(curRow, 2).Value = tableName;
                ws.Cell(curRow, 3).Value = "(테이블 없음)";
                curRow++;
                continue;
            }

            var cols = DbConnectionFactory.GetColumnNames(conn, tableName);
            if (cols.Count == 0)
            {
                ws.Cell(curRow, 1).Value = analyte;
                ws.Cell(curRow, 2).Value = tableName;
                ws.Cell(curRow, 3).Value = "(컬럼 없음)";
                curRow++;
                continue;
            }

            // 헤더 행 (항목 블록별로 반복): 분석항목 | 테이블 | col1 | col2 | ...
            ws.Cell(curRow, 1).Value = analyte;
            ws.Cell(curRow, 2).Value = tableName;
            for (int i = 0; i < cols.Count; i++)
                ws.Cell(curRow, 3 + i).Value = cols[i];
            ws.Range(curRow, 1, curRow, 2 + cols.Count).Style.Font.Bold = true;
            curRow++;

            // 필터: 시료명 + 업체명 (컬럼이 있을 때만)
            bool hasSampleCol = cols.Contains("시료명", StringComparer.OrdinalIgnoreCase);
            bool hasCompanyCol = cols.Contains("업체명", StringComparer.OrdinalIgnoreCase);
            var where = new List<string>();
            using var cmd = conn.CreateCommand();
            if (hasSampleCol)
            {
                where.Add("`시료명` = @s");
                cmd.Parameters.AddWithValue("@s", sample.시료명 ?? "");
            }
            if (hasCompanyCol)
            {
                where.Add("`업체명` = @c");
                cmd.Parameters.AddWithValue("@c", sample.의뢰사업장 ?? "");
            }
            cmd.CommandText = $"SELECT * FROM `{tableName}`"
                + (where.Count > 0 ? " WHERE " + string.Join(" AND ", where) : "");

            int dumped = 0;
            try
            {
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    for (int i = 0; i < cols.Count && i < r.FieldCount; i++)
                    {
                        var v = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
                        ws.Cell(curRow, 3 + i).Value = v;
                    }
                    curRow++;
                    dumped++;
                }
            }
            catch (Exception ex)
            {
                ws.Cell(curRow, 3).Value = $"(조회 오류: {ex.Message})";
                curRow++;
            }

            if (dumped == 0)
            {
                ws.Cell(curRow, 3).Value = "(매칭 레코드 없음)";
                curRow++;
            }

            curRow++; // 블록 간 빈 줄
        }

        try { ws.Columns().AdjustToContents(); } catch { }
        conn.Dispose();
    }

    private static void BuildPlaceholderSheet(XLWorkbook wb, string name)
    {
        var ws = wb.Worksheets.Add(name);
        ws.Cell(1, 1).Value = $"[스캐폴드] {name} 포맷 채움 로직 미구현";
        ws.Cell(2, 1).Value = "후속 단계에서 항목 → 시트 매핑 및 셀 레이아웃 구현 예정";
        ws.Range(1, 1, 2, 1).Style.Font.Italic = true;
    }

    private static string SanitizeForFileName(string input)
    {
        if (string.IsNullOrEmpty(input)) return "_";
        var invalid = Path.GetInvalidFileNameChars();
        var chars = input.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        return new string(chars);
    }
}
