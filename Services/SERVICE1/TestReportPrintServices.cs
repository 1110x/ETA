using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ETA.Models;
using System.Data;
using System.Data.Common;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class TestReportPrintService
{
    private static string RootPath =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));

    public static string Template1Path =>
        Path.Combine(RootPath, "Data", "Templates", "시험성적서1_template.xlsx");

    public static string Template2Path =>
        Path.Combine(RootPath, "Data", "Templates", "시험성적서2_template.xlsx");

    private static string IntegratedTemplatePath =>
        Path.Combine(RootPath, "Data", "Templates", "시험성적서통합.xlsm");

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

        // 파일명: 채취일자_시료명  (파일시스템 금지 문자 제거)
        var dateStr  = sample.채취일자.Replace("-", "").Replace("/", "");
        var safeName = string.IsNullOrWhiteSpace(sample.시료명)
            ? sample.약칭
            : sample.시료명;
        // 파일명에 사용 불가한 문자 제거: \ / : * ? " < > |
        foreach (var c in Path.GetInvalidFileNameChars())
            safeName = safeName.Replace(c.ToString(), "");
        safeName = safeName.Trim();
        if (safeName.Length > 60) safeName = safeName[..60];   // 너무 긴 경우 자르기

        var outPath = Path.Combine(OutputDir,
            $"{dateStr}_{safeName}.xlsx");

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

            // ★ PDF 변환 완료 후 중간 xlsx 파일 삭제 (PDF만 남김)
            try
            {
                if (System.IO.File.Exists(outPath) && pdfPath != outPath)
                    System.IO.File.Delete(outPath);
            }
            catch { /* 삭제 실패해도 무시 */ }

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
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            // 컬럼 목록 먼저 확인
            var cols = DbConnectionFactory.GetColumnNames(conn, "계약 DB");
            Log($"계약DB 컬럼: {string.Join(", ", cols)}");

            // 샘플 데이터 확인 (첫 3행)
            using var sample = conn.CreateCommand();
            sample.CommandText = @"SELECT * FROM `계약 DB` LIMIT 3";
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
                FROM `계약 DB`
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
        var soffice = FindLibreOffice();

        // ── LibreOffice 없으면 자동 설치 시도 ────────────────────────────
        if (soffice == null)
        {
            Log("[PDF] LibreOffice 없음 → 설치 시도");
            bool installed = TryInstallLibreOffice();

            if (installed)
                soffice = FindLibreOffice();  // 설치 후 재탐색

            if (soffice == null)
            {
                Log("[PDF] LibreOffice 설치 실패 → xlsx 반환");
                return xlsxPath;  // 실패 시 엑셀 파일 그대로 반환
            }
        }

        // ── LibreOffice 로 PDF 변환 ───────────────────────────────────────
        var dir = System.IO.Path.GetDirectoryName(xlsxPath)!;
        try
        {
            // ★ 프린터 연결 대기 제거:
            //   --norestore  : 이전 세션 복구 안 함
            //   --nofirststartwizard : 초기 설정 마법사 스킵
            //   SAL_USE_VCLPLUGIN=svp : 화면/프린터 드라이버 없는 순수 소프트웨어 렌더러 사용
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName               = soffice,
                // --printer-name "": 가상 프린터 지정으로 대기 제거
                // -env:UserInstallation: 독립 프로필로 충돌 방지
                Arguments = $"--headless --norestore --nofirststartwizard " +
                             $"-env:UserInstallation=file:///tmp/libreoffice_eta " +
                             $"--convert-to pdf:writer_pdf_Export " +
                             $"--outdir \"{dir}\" \"{xlsxPath}\"",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
            };

            // 프린터 스풀러 연결 차단 — 핵심 설정
            psi.EnvironmentVariables["SAL_USE_VCLPLUGIN"]   = "svp";
            psi.EnvironmentVariables["PGPASSWORD"]          = "";       // DB 연결 타임아웃 방지
            psi.EnvironmentVariables["LIBO_HEADLESS"]       = "1";

            using var proc = System.Diagnostics.Process.Start(psi)!;
            proc.WaitForExit(60_000);

            Log($"[PDF] 변환 종료코드={proc.ExitCode}");
        }
        catch (Exception ex)
        {
            Log($"[PDF] 변환 오류: {ex.Message}");
            return xlsxPath;
        }

        var pdf = System.IO.Path.ChangeExtension(xlsxPath, ".pdf");
        Log(System.IO.File.Exists(pdf) ? $"[PDF] 완료: {pdf}" : "[PDF] 변환 실패");
        return System.IO.File.Exists(pdf) ? pdf : xlsxPath;
    }

    // ── LibreOffice 경로 탐색 ─────────────────────────────────────────────
    private static string? FindLibreOffice()
    {
        string[] candidates =
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            // winget 설치 후 환경변수 PATH 에서 탐색
            FindInPath("soffice.exe") ?? "",
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        };
        return candidates.FirstOrDefault(System.IO.File.Exists);
    }

    // PATH 환경변수에서 실행파일 탐색
    private static string? FindInPath(string fileName)
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(System.IO.Path.PathSeparator))
        {
            var full = System.IO.Path.Combine(dir, fileName);
            if (System.IO.File.Exists(full)) return full;
        }
        return null;
    }

    // ── LibreOffice 자동 설치 (winget) ────────────────────────────────────
    private static bool TryInstallLibreOffice()
    {
        // winget 사용 가능 여부 확인
        if (!IsWingetAvailable())
        {
            Log("[설치] winget 없음 → 브라우저로 다운로드 페이지 열기");
            OpenDownloadPage();
            return false;
        }

        Log("[설치] winget 으로 LibreOffice 설치 시작...");

        try
        {
            // 사용자에게 설치 진행 안내 다이얼로그 (UI 스레드에서 호출된 경우)
            // 여기선 콘솔 로그만 — 실제 UI 알림은 TestReportPage.ShowToast() 에서 처리
            using var proc = System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo
                {
                    FileName               = "winget",
                    Arguments              = "install --id TheDocumentFoundation.LibreOffice -e --silent --accept-package-agreements --accept-source-agreements",
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true,
                    CreateNoWindow         = true,
                })!;

            // 최대 5분 대기
            bool finished = proc.WaitForExit(300_000);
            int  exitCode = finished ? proc.ExitCode : -1;

            Log($"[설치] winget 종료코드={exitCode}  finished={finished}");
            return exitCode == 0;
        }
        catch (Exception ex)
        {
            Log($"[설치] winget 실행 오류: {ex.Message}");
            OpenDownloadPage();
            return false;
        }
    }

    // winget 사용 가능 여부
    private static bool IsWingetAvailable()
    {
        try
        {
            using var proc = System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo
                {
                    FileName               = "winget",
                    Arguments              = "--version",
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow         = true,
                })!;
            proc.WaitForExit(5_000);
            return proc.ExitCode == 0;
        }
        catch { return false; }
    }

    // winget 도 없을 때 브라우저로 다운로드 페이지 열기
    private static void OpenDownloadPage()
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName        = "https://www.libreoffice.org/download/libreoffice-fresh/",
                UseShellExecute = true,
            });
            Log("[설치] 브라우저에서 LibreOffice 다운로드 페이지를 열었습니다.");
        }
        catch { }
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
        if (App.EnableLogging)
        {
            try
            {
                var logPath = Path.GetFullPath(
                    Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs/TestReportDebug.log"));
                File.AppendAllText(logPath, line + Environment.NewLine);
            }
            catch { }
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  일괄 엑셀 — 시험성적서통합.xlsm 의 자료 시트에 횡으로 입력
    //  ClosedXML 저장 없이 ZIP XML 직접 조작 → 스타일/매크로 완전 보존
    // ══════════════════════════════════════════════════════════════════════
    public static string ExportToIntegratedTemplate(
        IEnumerable<SampleRequest>       samples,
        Dictionary<string, AnalysisItem> meta)
    {
        string tplPath = IntegratedTemplatePath;
        if (!File.Exists(tplPath))
            throw new FileNotFoundException($"통합 템플릿 없음: {tplPath}");

        Directory.CreateDirectory(OutputDir);
        string savePath = Path.Combine(OutputDir,
            $"시험성적서통합_{DateTime.Now:yyyyMMdd_HHmmss}.xlsm");
        File.Copy(tplPath, savePath, overwrite: true);

        // ── 1. ZIP에서 sharedStrings + sheet3 row1 직접 파싱으로 헤더 맵 구성 ──
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        string ssXmlOrig  = "";
        string[] ssArr    = Array.Empty<string>();

        using (var zipR = ZipFile.OpenRead(savePath))
        {
            // sharedStrings.xml → ssArr[] 배열
            var ssEntry = zipR.GetEntry("xl/sharedStrings.xml");
            if (ssEntry != null)
            {
                using var sr = new StreamReader(ssEntry.Open(), Encoding.UTF8);
                ssXmlOrig = sr.ReadToEnd();
                var siMatches = Regex.Matches(ssXmlOrig, @"<si>.*?</si>", RegexOptions.Singleline);
                ssArr = new string[siMatches.Count];
                for (int k = 0; k < siMatches.Count; k++)
                {
                    var tm = Regex.Match(siMatches[k].Value, @"<t[^>]*>([^<]*)</t>");
                    ssArr[k] = tm.Success ? System.Net.WebUtility.HtmlDecode(tm.Groups[1].Value) : "";
                }
            }

            // sheet3.xml row r="1" 파싱 → colMap
            var shEntry = zipR.GetEntry("xl/worksheets/sheet3.xml");
            if (shEntry != null)
            {
                using var sr2 = new StreamReader(shEntry.Open(), Encoding.UTF8);
                string shXml = sr2.ReadToEnd();
                var row1m = Regex.Match(shXml, @"<row r=""1""[^>]*>(.*?)</row>", RegexOptions.Singleline);
                if (row1m.Success)
                {
                    foreach (Match cm in Regex.Matches(row1m.Groups[1].Value,
                        @"<c r=""([A-Z]+)1""[^>]*>(.*?)</c>", RegexOptions.Singleline))
                    {
                        int colNum = ColLetterToNum(cm.Groups[1].Value);
                        string inner = cm.Groups[2].Value;
                        string hdr = "";
                        var vm = Regex.Match(inner, @"<v>(\d+)</v>");
                        if (vm.Success && int.TryParse(vm.Groups[1].Value, out int idx) && idx < ssArr.Length)
                            hdr = ssArr[idx].Trim();
                        else
                        {
                            var tm2 = Regex.Match(inner, @"<t[^>]*>([^<]*)</t>");
                            if (tm2.Success) hdr = System.Net.WebUtility.HtmlDecode(tm2.Groups[1].Value).Trim();
                        }
                        if (!string.IsNullOrEmpty(hdr) && !colMap.ContainsKey(hdr))
                            colMap[hdr] = colNum;
                    }
                }
            }
        }

        int ColOf(params string[] candidates)
        {
            foreach (var c in candidates)
                if (colMap.TryGetValue(c, out int n)) return n;
            return 0;
        }
        int colCompany  = ColOf("업체명", "사업장명", "의뢰사업장");
        int colDiv      = ColOf("구분", "시료구분");
        int colReportNo = ColOf("시험성적번호", "성적서 번호", "성적서번호", "견적번호");
        int colDate     = ColOf("채수일자", "채취일자");
        int colSampler   = ColOf("시료채취1", "채수담당자", "시료채취자1", "시료채취자-1");
        int colSampler2  = ColOf("시료채취2", "시료채취자2", "시료채취자-2");
        int colSampleNm  = ColOf("시료명", "시료명칭");
        int colZone     = ColOf("특례지역", "방류기준");
        int colWitness  = ColOf("채수입회자", "입회자");
        int colEndDate   = ColOf("분석완료일", "분석완료일자", "분석종료일");
        int colPurpose  = ColOf("의뢰용도");
        int colQuality  = ColOf("정도보증 적용", "정도보증");
        int colRepresentative = ColOf("대표자");

        // ── 2. 모든 셀 값을 sharedStrings 에 추가, t="s" 참조 방식으로 행 XML 생성 ──
        var sampleList = samples.ToList();

        // 약칭별 대표자 캐시 (계약 DB 조회)
        var representativeCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 기존 SS 목록 끝에 새 문자열 추가 (중복 제거)
        var ssList = new List<string>(ssArr);
        int GetOrAddSs(string val)
        {
            int existing = ssList.IndexOf(val);
            if (existing >= 0) return existing;
            ssList.Add(val);
            return ssList.Count - 1;
        }

        // 먼저 모든 셀 데이터 구성
        // 견적번호별 A,B,C... suffix 카운터
        var reportNoCounter = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        string NextSuffix(string reportNo)
        {
            if (string.IsNullOrWhiteSpace(reportNo)) return "A";
            if (!reportNoCounter.TryGetValue(reportNo, out int n)) n = 0;
            reportNoCounter[reportNo] = n + 1;
            if (n < 26) return ((char)('A' + n)).ToString();
            return ((char)('A' + (n / 26) - 1)).ToString() + ((char)('A' + n % 26)).ToString();
        }

        var allRows = new List<(int rowNum, SortedDictionary<int, int> cells)>();
        for (int i = 0; i < sampleList.Count; i++)
        {
            var s      = sampleList[i];
            int rowNum = i + 2;
            string suffix = NextSuffix(s.견적번호);
            // 대표자: 약칭으로 계약 DB 조회 (캐시)
            if (!representativeCache.TryGetValue(s.의뢰사업장, out string? repVal))
            {
                repVal = ContractService.GetRepresentativeByCompany(s.의뢰사업장);
                representativeCache[s.의뢰사업장] = repVal;
            }
            var cells  = new SortedDictionary<int, int>(); // col → ssIndex

            void Add(int col, string val)
            {
                if (col > 0 && !string.IsNullOrWhiteSpace(val))
                    cells[col] = GetOrAddSs(val);
            }

            Add(colCompany,  s.의뢰사업장);
            Add(colDiv,      s.방류허용기준);
            Add(colReportNo, string.IsNullOrWhiteSpace(s.견적번호) ? "" : s.견적번호 + suffix);
            Add(colDate,     s.채취일자);
            Add(colSampler,  s.시료채취자1);
            Add(colSampleNm, s.시료명);
            Add(colSampler2, s.시료채취자2);
            Add(colZone,     s.견적구분);
            Add(colWitness,  s.입회자);
            Add(colEndDate,  s.분석종료일);
            Add(colPurpose,  s.약칭);
            Add(colQuality,  s.정도보증);
            Add(colRepresentative, repVal ?? "");

            foreach (var kv in s.분석결과)
            {
                if (string.IsNullOrWhiteSpace(kv.Value)) continue;
                if (colMap.TryGetValue(kv.Key, out int ac))
                    cells[ac] = GetOrAddSs(kv.Value);
            }

            allRows.Add((rowNum, cells));
        }

        // 행 XML 빌드
        var rowsXml = new StringBuilder();
        foreach (var (rowNum, cells) in allRows)
        {
            if (cells.Count == 0) continue;
            int minCol = cells.Keys.Min();
            int maxCol = cells.Keys.Max();
            rowsXml.Append($"<row r=\"{rowNum}\" spans=\"{minCol}:{maxCol}\">");
            foreach (var kv in cells)
            {
                string cellRef = ColNumToLetter(kv.Key) + rowNum;
                rowsXml.Append($"<c r=\"{cellRef}\" t=\"s\"><v>{kv.Value}</v></c>");
            }
            rowsXml.Append("</row>");
        }

        // ── 3. ZIP 안의 sharedStrings.xml + sheet3.xml 교체, calcChain 삭제 ──
        const string sheetEntry = "xl/worksheets/sheet3.xml";
        const string ssEntry2   = "xl/sharedStrings.xml";
        using (var zip = ZipFile.Open(savePath, ZipArchiveMode.Update))
        {
            // calcChain.xml 삭제
            zip.GetEntry("xl/calcChain.xml")?.Delete();

            // sharedStrings.xml 갱신
            {
                var ssZipEntry = zip.GetEntry(ssEntry2)
                    ?? throw new InvalidOperationException("sharedStrings.xml 을 찾을 수 없습니다.");
                int totalCount = ssList.Count;
                var newSsXml = new StringBuilder();
                newSsXml.Append($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                newSsXml.Append($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{totalCount}\" uniqueCount=\"{totalCount}\">");
                foreach (var sv in ssList)
                    newSsXml.Append($"<si><t>{XmlEscape(sv)}</t></si>");
                newSsXml.Append("</sst>");
                ssZipEntry.Delete();
                var newSsEntry = zip.CreateEntry(ssEntry2, CompressionLevel.Optimal);
                using var ssw = new StreamWriter(newSsEntry.Open(), new UTF8Encoding(false));
                ssw.Write(newSsXml.ToString());
            }

            // sheet3.xml 갱신
            {
                var entry = zip.GetEntry(sheetEntry)
                    ?? throw new InvalidOperationException($"{sheetEntry} 를 찾을 수 없습니다.");
                string sheetXml;
                using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                    sheetXml = sr.ReadToEnd();

                // 기존 Row2 이상 모두 제거 (r="2"~r="9", r="10" 이상 모두 포함)
                sheetXml = Regex.Replace(sheetXml,
                    @"<row r=""(?:[2-9]|\d{2,})""[^>]*>.*?</row>",
                    "", RegexOptions.Singleline);

                // 새 행 삽입
                sheetXml = sheetXml.Replace("</sheetData>", rowsXml.ToString() + "</sheetData>");

                // dimension 갱신
                int newLastRow = Math.Max(1, sampleList.Count + 1);
                sheetXml = Regex.Replace(sheetXml,
                    @"ref=""A1:[A-Z]+\d+""",
                    $@"ref=""A1:BX{newLastRow}""");

                entry.Delete();
                var newEntry = zip.CreateEntry(sheetEntry, CompressionLevel.Optimal);
                using var sw = new StreamWriter(newEntry.Open(), new UTF8Encoding(false));
                sw.Write(sheetXml);
            }
        }

        Log($"일괄 통합 저장: {savePath}  ({sampleList.Count}건)");
        return savePath;
    }

    private static int ColLetterToNum(string col)
    {
        int n = 0;
        foreach (char ch in col.ToUpperInvariant())
            n = n * 26 + (ch - 'A' + 1);
        return n;
    }

    private static string ColNumToLetter(int col)
    {
        var sb = new StringBuilder();
        while (col > 0)
        {
            col--;
            sb.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        return sb.ToString();
    }

    private static string XmlEscape(string val) =>
        val.Replace("&", "&amp;")
           .Replace("<", "&lt;")
           .Replace(">", "&gt;")
           .Replace("\"", "&quot;");
}