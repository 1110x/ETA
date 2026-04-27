using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 시험성적서 Word(.docx) 출력 — 템플릿-치환 방식.
/// `Data/Templates/시험성적서/시험성적서_template.docx` 의 `{{키}}` 치환자를 시료 데이터로 채움.
///
/// 헤더 치환자(문서 어디서나 단일 치환):
///   {{성적서번호}} {{회사명}} {{대표자}} {{시료명}} {{채취일자}} {{채취자}}
///   {{입회자}}    {{분석종료일}} {{용도}} {{비고문구}} {{서명}}
///
/// 항목 행 치환자(`{{번호}}` 가 들어있는 표 행 1개를 반복 복제):
///   {{번호}} {{구분}} {{항목}} {{ES}} {{결과}} {{단위}} {{기준}}
/// </summary>
public static class TestReportWordPrintService
{
    /// <summary>분석 항목 수가 이 값 이하면 단일페이지 양식, 초과면 다중페이지 양식 사용.</summary>
    private const int SinglePageRowThreshold = 28;

    public static string TemplatePath1Page => TemplateConfiguration.Resolve("TestReportWord1Page");
    public static string TemplatePathMulti => TemplateConfiguration.Resolve("TestReportWordMulti");

    private static string OutputDir =>
        Path.Combine(AppPaths.RootPath, "Data", "Reports");

    public static bool TemplateExists() =>
        File.Exists(TemplatePath1Page) && File.Exists(TemplatePathMulti);

    /// <summary>두 양식(단일/다중) 자동 생성. 이미 있으면 덮어씀.</summary>
    public static string GenerateDefaultTemplate()
    {
        TestReportWordTemplateGenerator.GenerateSinglePage(TemplatePath1Page);
        TestReportWordTemplateGenerator.GenerateMultiPage(TemplatePathMulti);
        Log($"기본 템플릿 생성: {TemplatePath1Page}, {TemplatePathMulti}");
        return TemplatePath1Page;
    }

    /// <summary>분석 항목 수에 따라 적절한 양식 경로 선택.</summary>
    private static string PickTemplate(int rowCount) =>
        rowCount <= SinglePageRowThreshold ? TemplatePath1Page : TemplatePathMulti;

    public static string FillAndSave(
        SampleRequest                    sample,
        List<AnalysisResultRow>          rows,
        Dictionary<string, AnalysisItem> meta,
        string reportNo       = "",
        string qualityMgr     = "",
        bool   includeStandard = true,
        bool   openAfter       = true)
    {
        // 템플릿이 없으면 시험기록부 시각언어로 기본 템플릿 자동 생성
        if (!TemplateExists()) GenerateDefaultTemplate();

        Directory.CreateDirectory(OutputDir);

        // 출력 파일명: 채취일자_시료명.docx
        var dateStr  = sample.채취일자.Replace("-", "").Replace("/", "");
        var safeName = string.IsNullOrWhiteSpace(sample.시료명) ? sample.약칭 : sample.시료명;
        foreach (var c in Path.GetInvalidFileNameChars())
            safeName = safeName.Replace(c.ToString(), "");
        safeName = safeName.Trim();
        if (safeName.Length > 60) safeName = safeName[..60];

        var outPath = Path.Combine(OutputDir, $"{dateStr}_{safeName}.docx");
        var templatePath = PickTemplate(rows.Count);
        Log($"양식 선택: {(rows.Count <= SinglePageRowThreshold ? "단일페이지" : "다중페이지")} ({rows.Count}건) → {templatePath}");
        File.Copy(templatePath, outPath, overwrite: true);

        var no = string.IsNullOrEmpty(reportNo)
            ? $"WAC-{DateTime.Now:yyyyMMdd}-{sample.약칭}"
            : reportNo;

        var (companyName, representative) = GetContractInfo(sample.약칭);
        var standardMap = BuildStandardMap(rows, sample.방류허용기준);

        bool isQC = sample.정도보증.Trim().ToUpper() == "O"
                 || sample.정도보증.Trim() == "정도보증 적용"
                 || sample.정도보증.Trim() == "Y";

        // 품질책임자 — 호출자가 지정 안 했으면 Agent 테이블에서 직급 '품책' 보유자 자동 조회
        if (string.IsNullOrWhiteSpace(qualityMgr))
            qualityMgr = GetQualityManagerName();

        var headerMap = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["성적서번호"] = no,
            ["회사명"]     = string.IsNullOrEmpty(companyName) ? sample.의뢰사업장 : companyName,
            ["대표자"]     = representative,
            ["시료명"]     = sample.시료명 ?? "",
            ["채취일자"]   = FormatDate(sample.채취일자),
            ["채취자"]     = $"{sample.시료채취자1} {sample.시료채취자2}".Trim(),
            ["입회자"]     = sample.입회자 ?? "",
            ["분석종료일"] = FormatDate(sample.분석종료일),
            ["용도"]       = isQC ? "정도보증 적용" : "참고용",
            ["비고문구"]   = isQC
                ? "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분석결과 입니다."
                : "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분석결과입니다.",
            ["품질책임자"] = qualityMgr ?? "",
        };

        using (var doc = WordprocessingDocument.Open(outPath, isEditable: true))
        {
            var body = doc.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Word 본문 없음");

            // 1) 방류기준 토글 OFF — 항목 표 헤더의 "기준" 텍스트를 "비고" 로 변경.
            //    값은 ExpandItemRows 단계에서 standardMap 을 비워 빈 문자열로 채움.
            if (!includeStandard)
            {
                ReplaceHeaderText(body, "기준", "비고");
                standardMap.Clear();
            }

            // 2) 항목 행 반복 — `{{번호}}` 포함 행이 있는 첫 표를 찾아 복제
            ExpandItemRows(body, rows, standardMap);

            // 3) 헤더 등 문서 전체 치환자 채우기 (반복 후에 실행 — 복제된 행도 같이 처리)
            foreach (var p in body.Descendants<Paragraph>().ToList())
                ReplacePlaceholders(p, headerMap);

            doc.MainDocumentPart!.Document!.Save();
        }

        Log($"저장: {outPath}");
        if (openAfter) OpenFile(outPath);
        return outPath;
    }

    // ── 헤더 텍스트 치환 — `{{번호}}` 포함 표의 헤더 행에서 oldText → newText 교체 ─
    private static void ReplaceHeaderText(Body body, string oldText, string newText)
    {
        foreach (var table in body.Descendants<Table>())
        {
            var rows = table.Elements<TableRow>().ToList();
            // `{{번호}}` 가 들어있는 행의 바로 위(또는 그 행 자신 이전) = 헤더 행
            int templateIdx = rows.FindIndex(r => r.InnerText.Contains("{{번호}}"));
            if (templateIdx <= 0) continue;
            var headerRow = rows[templateIdx - 1];
            foreach (var t in headerRow.Descendants<Text>())
            {
                if (t.Text == oldText) { t.Text = newText; break; }
            }
            return;
        }
    }

    // ── 표 항목 행 반복 처리 ────────────────────────────────────────────────
    private static void ExpandItemRows(
        Body body,
        List<AnalysisResultRow> rows,
        Dictionary<string, string> standardMap)
    {
        foreach (var table in body.Descendants<Table>())
        {
            var template = table.Elements<TableRow>()
                .FirstOrDefault(r => r.InnerText.Contains("{{번호}}"));
            if (template == null) continue;

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var clone = (TableRow)template.CloneNode(true);
                var rowMap = new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["번호"] = (i + 1).ToString(),
                    ["구분"] = row.Category ?? "",
                    ["항목"] = row.항목명 ?? "",
                    ["ES"]   = row.ES ?? "",
                    ["결과"] = row.결과값 ?? "",
                    ["단위"] = row.단위 ?? "",
                    ["기준"] = standardMap.GetValueOrDefault(row.항목명 ?? "", ""),
                };
                foreach (var p in clone.Descendants<Paragraph>())
                    ReplacePlaceholders(p, rowMap);
                template.InsertBeforeSelf(clone);
            }
            template.Remove();
            return; // 첫 매칭 표만 처리
        }
    }

    // ── 단락 내 `{{key}}` 치환 ────────────────────────────────────────────────
    // 1차: 각 Text 단위 치환 — 단일 Text 안에 들어있는 치환자는 스타일 보존하며 처리.
    //      예: "(서명)" 같은 이탤릭 Run 이 같은 단락에 있어도 그 Run 의 스타일은 그대로.
    // 2차: 단락에 아직 `{{` 가 남아있으면 Run 분할로 인한 잘림이라
    //      모든 Text 를 합쳐 첫 노드에 몰아넣음 (스타일은 첫 Run 기준).
    private static void ReplacePlaceholders(Paragraph p, IDictionary<string, string> map)
    {
        foreach (var t in p.Descendants<Text>())
        {
            if (t.Text.IndexOf("{{", StringComparison.Ordinal) < 0) continue;
            var s = t.Text;
            foreach (var kv in map)
                s = s.Replace("{{" + kv.Key + "}}", kv.Value ?? "");
            t.Text  = s;
            t.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
        }

        var texts = p.Descendants<Text>().ToList();
        if (texts.Count == 0) return;
        var combined = string.Concat(texts.Select(t => t.Text));
        if (combined.IndexOf("{{", StringComparison.Ordinal) < 0) return;
        foreach (var kv in map)
            combined = combined.Replace("{{" + kv.Key + "}}", kv.Value ?? "");
        texts[0].Text  = combined;
        texts[0].Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
        for (int i = 1; i < texts.Count; i++)
            texts[i].Text = "";
    }

    // ── 방류허용기준 매핑 (Excel 익스포터와 동일 로직) ─────────────────────
    private static Dictionary<string, string> BuildStandardMap(
        List<AnalysisResultRow> rows, string 방류허용기준컬럼)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(방류허용기준컬럼)) return map;
        foreach (var r in rows)
        {
            if (map.ContainsKey(r.항목명)) continue;
            var val = TestReportService.GetStandardValue(r.항목명, 방류허용기준컬럼);
            if (!string.IsNullOrEmpty(val)) map[r.항목명] = val;
        }
        return map;
    }

    // ── Agent 에서 품책 직급 보유자 성명 조회 ──────────────────────────────
    private static string GetQualityManagerName()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            // 직급이 정확히 '품책' 인 사람 우선, 없으면 '품질책임' 포함자 폴백
            cmd.CommandText = @"
                SELECT `성명` FROM `Agent`
                WHERE TRIM(`직급`) = '품책'
                   OR `직급` LIKE '%품질책임%'
                ORDER BY (TRIM(`직급`) = '품책') DESC
                LIMIT 1";
            var result = cmd.ExecuteScalar();
            var name = result?.ToString() ?? "";
            Log($"품책 자동조회: '{name}'");
            return name;
        }
        catch (Exception ex)
        {
            Log($"품책 조회 오류: {ex.Message}");
            return "";
        }
    }

    // ── 계약 DB 에서 회사명/대표자 조회 ────────────────────────────────────
    private static (string companyName, string representative) GetContractInfo(string 약칭)
    {
        if (string.IsNullOrEmpty(약칭)) return ("", "");
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
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
                return (cn, rep);
            }
            return ("", "");
        }
        catch (Exception ex)
        {
            Log($"계약정보 조회 오류: {ex.Message}");
            return ("", "");
        }
    }

    private static string FormatDate(string val)
    {
        if (string.IsNullOrEmpty(val)) return "";
        if (DateTime.TryParse(val, out var dt)) return dt.ToString("yyyy년 M월 d일");
        return val;
    }

    public static void OpenFile(string path)
    {
        try
        {
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex) { Log($"OpenFile: {ex.Message}"); }
    }

    public static void OpenTemplateFolder()
    {
        var dir = Path.GetDirectoryName(TemplatePath1Page)!;
        Directory.CreateDirectory(dir);
        OpenFile(dir);
    }

    private static void Log(string msg)
    {
        if (!App.EnableLogging) return;
        try
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] [PrintWord] {msg}";
            File.AppendAllText(
                Path.Combine(AppPaths.LogsDir, "TestReportDebug.log"),
                line + Environment.NewLine);
        }
        catch { }
    }
}
