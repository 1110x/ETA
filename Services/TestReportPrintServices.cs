using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ETA.Models;
using Microsoft.Data.Sqlite;

namespace ETA.Services;

public static class TestReportPrintService
{
    private static string RootPath =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));

    public static string Template1Path =>
        Path.Combine(RootPath, "Data", "Templates", "시험성적서1_template.xlsx");

    public static string Template2Path =>
        Path.Combine(RootPath, "Data", "Templates", "시험성적서2_template.xlsx");

    private static string OutputDir =>
        Path.Combine(RootPath, "Data", "Reports");

    public static bool TemplateExists() =>
        File.Exists(Template1Path) || File.Exists(Template2Path);

    public static string TemplatePath => Template1Path;

    public static string FillAndSave(
        SampleRequest           sample,
        List<AnalysisResultRow> rows,
        Dictionary<string, AnalysisItem> meta,
        string reportNo   = "",
        string qualityMgr = "",
        bool   toPdf      = false,
        bool   openAfter  = true)
    {
        if (!TemplateExists())
            throw new FileNotFoundException(
                $"템플릿 없음\n경로: {Template1Path}\n\nData\\Templates\\ 폴더에 템플릿 파일을 넣어주세요.");

        Directory.CreateDirectory(OutputDir);

        var safe = sample.채취일자.Replace("-", "").Replace("/", "");
        var outPath = Path.Combine(OutputDir,
            $"성적서_{sample.약칭}_{safe}_{DateTime.Now:HHmmss}.xlsx");

        var no = string.IsNullOrEmpty(reportNo)
            ? $"WAC-{DateTime.Now:yyyyMMdd}-{sample.약칭}"
            : reportNo;

        var (companyName, representative) = GetContractInfo(sample.약칭);
        var standardMap    = BuildStandardMap(rows, sample.방류허용기준);

        int totalPages = Math.Max(1, (int)Math.Ceiling(rows.Count / 64.0));

        if (totalPages == 1)
        {
            // 단일 페이지: 템플릿 직접 복사 → 열너비/인쇄영역 100% 보존
            var pageRows  = rows.Take(64).ToList();
            var leftRows  = pageRows.Take(32).ToList();
            var rightRows = pageRows.Skip(32).Take(32).ToList();
            bool hasRight = rightRows.Count > 0;

            var tplPath = hasRight
                ? (File.Exists(Template2Path) ? Template2Path : Template1Path)
                : (File.Exists(Template1Path) ? Template1Path : Template2Path);

            File.Copy(tplPath, outPath, overwrite: true);

            using var wb = new XLWorkbook(outPath);
            var ws = wb.Worksheet(1);
            FillSheet(ws, sample, leftRows, rightRows,
                      no, qualityMgr, companyName, representative, standardMap,
                      0, hasRight);
            wb.Save();
        }
        else
        {
            // 복수 페이지: 첫 페이지 = 템플릿 복사, 이후 시트 추가
            var firstLeft  = rows.Take(32).ToList();
            var firstRight = rows.Skip(32).Take(32).ToList();
            bool firstHasRight = firstRight.Count > 0;

            var firstTpl = firstHasRight
                ? (File.Exists(Template2Path) ? Template2Path : Template1Path)
                : (File.Exists(Template1Path) ? Template1Path : Template2Path);

            File.Copy(firstTpl, outPath, overwrite: true);

            using var wb = new XLWorkbook(outPath);
            var ws0 = wb.Worksheet(1);
            ws0.Name = "시험성적서1";
            FillSheet(ws0, sample, firstLeft, firstRight,
                      no, qualityMgr, companyName, representative, standardMap,
                      0, firstHasRight);

            for (int page = 1; page < totalPages; page++)
            {
                var pageRows  = rows.Skip(page * 64).ToList();
                var leftRows  = pageRows.Take(32).ToList();
                var rightRows = pageRows.Skip(32).Take(32).ToList();
                bool hasRight = rightRows.Count > 0;

                var tplPath = hasRight
                    ? (File.Exists(Template2Path) ? Template2Path : Template1Path)
                    : (File.Exists(Template1Path) ? Template1Path : Template2Path);

                using var tplWb = new XLWorkbook(tplPath);
                tplWb.Worksheet(1).CopyTo(wb, $"시험성적서{page + 1}");
                var ws = wb.Worksheet($"시험성적서{page + 1}");
                FillSheet(ws, sample, leftRows, rightRows,
                          no, qualityMgr, companyName, representative, standardMap,
                          page, hasRight);
            }

            wb.Save();
        }
        Log($"저장: {outPath}");

        if (toPdf)
        {
            var pdfPath = ConvertToPdf(outPath);
            if (openAfter) OpenFile(pdfPath);
            return pdfPath;
        }

        if (openAfter) OpenFile(outPath);
        return outPath;
    }

    private static void FillSheet(
        IXLWorksheet ws,
        SampleRequest sample,
        List<AnalysisResultRow> left,
        List<AnalysisResultRow> right,
        string reportNo, string qualityMgr, string companyName, string representative,
        Dictionary<string,string> standardMap,
        int pageIndex, bool hasRight)
    {
        string suffix = pageIndex == 0 ? "" : $"-{pageIndex + 1}";

        // 좌측 헤더
        S(ws, "C1", reportNo + "-A" + suffix);
        S(ws, "C3", string.IsNullOrEmpty(companyName) ? sample.의뢰사업장 : companyName);
        S(ws, "F3", FormatDate(sample.채취일자));
        S(ws, "C4", representative);
        S(ws, "F4", $"{sample.시료채취자1} {sample.시료채취자2}".Trim());
        S(ws, "C5", sample.입회자);
        S(ws, "F5", FormatDate(sample.분석종료일));
        S(ws, "C7", sample.시료명);
        // 정도보증 여부에 따라 의뢰정보(용도) 및 하단 문구 결정
        bool isQC = sample.정도보증.Trim().ToUpper() == "O"
                 || sample.정도보증.Trim() == "정도보증 적용"
                 || sample.정도보증.Trim() == "Y";
        S(ws, "F7", isQC ? "정도보증 적용" : "참고용");
        S(ws, "A44", isQC
            ? "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분석결과 입니다."
            : "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분석결과입니다.");

        // 우측 헤더 (2페이지 템플릿)
        if (hasRight)
        {
            S(ws, "K1", reportNo + "-B" + suffix);
            S(ws, "K3", string.IsNullOrEmpty(companyName) ? sample.의뢰사업장 : companyName);
            S(ws, "N3", FormatDate(sample.채취일자));
            S(ws, "K4", representative);
            S(ws, "N4", $"{sample.시료채취자1} {sample.시료채취자2}".Trim());
            S(ws, "K5", sample.입회자);
            S(ws, "N5", FormatDate(sample.분석종료일));
            S(ws, "K7", sample.시료명);
            S(ws, "N7", isQC ? "정도보증 적용" : "참고용");
            S(ws, "I44", isQC
                ? "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분석결과 입니다."
                : "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분석결과입니다.");
        }

        // 좌측 항목: A=번호, B=구분, D=항목, E=ES번호, F=결과, G=단위, H=방류기준
        for (int i = 0; i < 32; i++)
        {
            int r = 10 + i;
            if (i < left.Count)
            {
                var row = left[i];
                ws.Cell(r, 1).Value = pageIndex * 64 + i + 1;
                ws.Cell(r, 2).Value = row.Category;
                ws.Cell(r, 4).Value = row.항목명;
                ws.Cell(r, 5).Value = row.ES;
                ws.Cell(r, 6).Value = ParseResult(row.결과값);
                ws.Cell(r, 7).Value = row.단위;
                ws.Cell(r, 8).Value = standardMap.GetValueOrDefault(row.항목명, "");
            }
            else
            {
                foreach (int col in new[]{ 1, 2, 4, 5, 6, 7, 8 })
                    ws.Cell(r, col).Value = "";
            }
        }

        // 우측 항목: I=번호, J=구분, L=항목, M=ES번호, N=결과, O=단위, P=방류기준
        if (hasRight)
        {
            for (int i = 0; i < 32; i++)
            {
                int r = 10 + i;
                if (i < right.Count)
                {
                    var row = right[i];
                    ws.Cell(r, 9).Value  = pageIndex * 64 + 32 + i + 1;
                    ws.Cell(r, 10).Value = row.Category;
                    ws.Cell(r, 12).Value = row.항목명;
                    ws.Cell(r, 13).Value = row.ES;
                    ws.Cell(r, 14).Value = ParseResult(row.결과값);
                    ws.Cell(r, 15).Value = row.단위;
                    ws.Cell(r, 16).Value = standardMap.GetValueOrDefault(row.항목명, "");
                }
                else
                {
                    foreach (int col in new[]{ 9, 10, 12, 13, 14, 15, 16 })
                        ws.Cell(r, col).Value = "";
                }
            }
        }

        // 서명
        if (!string.IsNullOrEmpty(qualityMgr))
        {
            var sig = $"품질책임 수질분야 환경측정분석사       {qualityMgr}       (서명)";
            ws.Cell(43, 1).Value = sig;
            if (hasRight) ws.Cell(43, 9).Value = sig;
        }
    }

    private static Dictionary<string, string> BuildStandardMap(
        List<AnalysisResultRow> rows, string 방류허용기준컬럼)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(방류허용기준컬럼)) return map;
        foreach (var r in rows)
        {
            if (!map.ContainsKey(r.항목명))
            {
                var val = TestReportService.GetStandardValue(r.항목명, 방류허용기준컬럼);
                if (!string.IsNullOrEmpty(val)) map[r.항목명] = val;
            }
        }
        Log($"방류기준 {map.Count}개 조회");
        return map;
    }

    // 계약 DB에서 회사이름(C_CompanyName), 대표자(C_Representative) 조회
    private static (string companyName, string representative) GetContractInfo(string 약칭)
    {
        if (string.IsNullOrEmpty(약칭)) return ("", "");
        var dbPath = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Data", "eta.db"));
        if (!File.Exists(dbPath)) { Debug.WriteLine("[Print] DB없음"); return ("", ""); }
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            // 컬럼 목록 먼저 확인
            using var pragma = conn.CreateCommand();
            pragma.CommandText = @"PRAGMA table_info(""계약 DB"")";
            var cols = new System.Collections.Generic.List<string>();
            using (var pr = pragma.ExecuteReader())
                while (pr.Read()) cols.Add(pr.GetString(1));
            Log($"계약DB 컬럼: {string.Join(", ", cols)}");

            // 샘플 데이터 확인 (첫 3행)
            using var sample = conn.CreateCommand();
            sample.CommandText = @"SELECT * FROM ""계약 DB"" LIMIT 3";
            using (var sr = sample.ExecuteReader())
                while (sr.Read())
                {
                    var row = string.Join(" | ", Enumerable.Range(0, sr.FieldCount)
                        .Select(i => $"{sr.GetName(i)}={sr.GetValue(i)}"));
                    Log($"계약DB 샘플: {row}");
                }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT C_CompanyName, C_Representative
                FROM ""계약 DB""
                WHERE C_Abbreviation = @exact
                   OR C_Abbreviation LIKE @name
                   OR C_CompanyName  LIKE @name
                LIMIT 1";
            cmd.Parameters.AddWithValue("@exact", 약칭);
            cmd.Parameters.AddWithValue("@name",  $"%{약칭}%");
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                var cn  = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                var rep = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString() ?? "";
                Log($"계약정보: '{약칭}' → 회사={cn}, 대표자={rep}");
                return (cn, rep);
            }
            Log($"계약정보 없음: '{약칭}'");
            return ("", "");
        }
        catch (Exception ex)
        {
            Log($"계약정보 조회 오류: {ex.Message}");
            return ("", "");
        }
    }

    private static void S(IXLWorksheet ws, string addr, string val)
    {
        try { ws.Cell(addr).Value = val ?? ""; }
        catch (Exception ex) { Log($"{addr}: {ex.Message}"); }
    }

    private static XLCellValue ParseResult(string val)
    {
        if (string.IsNullOrWhiteSpace(val)) return "";
        if (double.TryParse(val,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double d)) return d;
        return val;
    }

    private static string FormatDate(string val)
    {
        if (string.IsNullOrEmpty(val)) return "";
        if (DateTime.TryParse(val, out var dt)) return dt.ToString("yyyy년 M월 d일");
        return val;
    }

    public static string ConvertToPdf(string xlsxPath)
    {
        var dir = Path.GetDirectoryName(xlsxPath)!;
        string[] candidates = {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/libreoffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        };
        var soffice = candidates.FirstOrDefault(File.Exists);
        if (soffice == null) { Debug.WriteLine("[Print] LibreOffice 없음"); return xlsxPath; }
        using var proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = soffice,
            Arguments = $"--headless --convert-to pdf --outdir \"{dir}\" \"{xlsxPath}\"",
            UseShellExecute = false, RedirectStandardOutput = true,
            RedirectStandardError = true, CreateNoWindow = true,
        })!;
        proc.WaitForExit(60_000);
        var pdf = Path.ChangeExtension(xlsxPath, ".pdf");
        Debug.WriteLine(File.Exists(pdf) ? $"[Print] PDF: {pdf}" : "[Print] PDF 변환 실패");
        return File.Exists(pdf) ? pdf : xlsxPath;
    }

    public static void OpenFile(string path)
    {
        try { System.Diagnostics.Process.Start(
            new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { Log($"OpenFile: {ex.Message}"); }
    }

    public static void OpenTemplateFolder()
    {
        var dir = Path.GetDirectoryName(Template1Path)!;
        Directory.CreateDirectory(dir);
        OpenFile(dir);
    }

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [Print] {msg}";
        Debug.WriteLine(line);
        try
        {
            var logPath = Path.GetFullPath(
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "TestReportDebug.log"));
            File.AppendAllText(logPath, line + Environment.NewLine);
        }
        catch { }
    }
}