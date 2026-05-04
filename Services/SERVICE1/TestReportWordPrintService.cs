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
///   {{입회자}}    {{분석종료일}} {{용도}} {{비고문구}} {{시험성적서서명}}
///
/// 항목 행 치환자(`{{번호}}` 가 들어있는 표 행 1개를 반복 복제):
///   {{번호}} {{구분}} {{항목}} {{ES}} {{결과}} {{단위}} {{기준}}
/// </summary>
public static class TestReportWordPrintService
{
    /// <summary>분석 항목 수가 이 값 이하면 단일페이지 양식, 초과면 다중페이지 양식 사용.</summary>
    private const int SinglePageRowThreshold = 32;

    public static string TemplatePath1Page => TemplateConfiguration.Resolve("TestReportWord1Page");
    public static string TemplatePathMulti => TemplateConfiguration.Resolve("TestReportWordMulti");

    /// <summary>출력 폴더 — 출력보관함(ReportsPanel) 이 보는 위치와 동일하게 `Data/Reports`.
    /// 출력 즉시 Show4 출력보관함 패널에 표시되도록 통일.</summary>
    private static string OutputDir
    {
        get
        {
            var dir = Path.Combine(AppPaths.RootPath, "Data", "Reports");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

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

    /// <summary>옛 스키마가 들어있는 구버전 템플릿이면 true. 다음 마커 중 하나라도 해당.
    ///   1) 본문에 `{{시험성적서서명}}` 치환자가 없음 (신버전 필수 마커)
    ///   2) 본문에 옛 `{{품질책임자}}` 치환자가 살아있음
    ///   3) 본문에 `{{성적서번호}}` 가 있음 — 신버전은 헤더로 이동
    ///   4) 헤더에 PAGE 필드가 없음 — 신버전은 헤더에 페이지번호 표시
    ///   5) 결재라인 anchored 표의 TablePositionY 가 옛 값(13200) — 신버전은 14000</summary>
    private static bool IsLegacyTemplate(string path)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(path, isEditable: false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return false;
            var bodyText = string.Concat(body.Descendants<Text>().Select(t => t.Text));
            if (!bodyText.Contains("{{시험성적서서명}}")) return true;
            if (bodyText.Contains("{{품질책임자}}"))     return true;
            if (bodyText.Contains("{{성적서번호}}"))     return true;

            // 헤더에 PAGE 필드(자동 페이지번호) 가 없으면 옛버전.
            bool hdrHasPage = doc.MainDocumentPart!.HeaderParts.Any(hp =>
                hp.Header?.Descendants<FieldCode>().Any(f => f.Text.Contains("PAGE")) == true);
            if (!hdrHasPage) return true;

            if (!bodyText.Contains("용도") || !bodyText.Contains("입회자")) return true;

            // 시료유형 분기 placeholder 부재 — 신버전 메타 표 3행은 동적 라벨 사용
            if (!bodyText.Contains("{{라벨_일자}}")) return true;

            // 메타 표 행 수 검사 — 신버전은 3행, 옛버전은 4행.
            // 본문 첫 번째 표 = BuildHeaderTable() (메타 표).
            var firstTable = body.Descendants<Table>().FirstOrDefault();
            int metaRows = firstTable?.Elements<TableRow>().Count() ?? 0;
            if (metaRows == 4) return true;

            // 결재라인 anchored 표의 TablePositionY 가 옛 값(13200) 이면 옛버전.
            var oldAnchorY = body.Descendants<TablePositionProperties>()
                .Any(tp => tp.TablePositionY?.Value == 13200);
            if (oldAnchorY) return true;

            // 페이지 상단 마진이 옛 값(1701, 3cm) 이면 옛버전.
            // 신버전(1134, 2cm) 은 32항목 + 결재라인 한 페이지 수용.
            var oldTopMargin = body.Descendants<PageMargin>()
                .Any(pm => pm.Top?.Value == 1701);
            if (oldTopMargin) return true;

            return false;
        }
        catch { return false; }
    }

    /// <summary>레거시 템플릿이 디스크에 있으면 자동 재생성. 결재정보 변경이 즉시 반영되도록 보장.</summary>
    private static void AutoMigrateTemplates()
    {
        bool needRegen = false;
        if (File.Exists(TemplatePath1Page) && IsLegacyTemplate(TemplatePath1Page)) needRegen = true;
        if (File.Exists(TemplatePathMulti) && IsLegacyTemplate(TemplatePathMulti)) needRegen = true;
        if (needRegen)
        {
            Log("레거시 템플릿 감지 → 자동 재생성");
            GenerateDefaultTemplate();
        }
    }

    /// <summary>분석 항목 수에 따라 적절한 양식 경로 선택.</summary>
    private static string PickTemplate(int rowCount) =>
        rowCount <= SinglePageRowThreshold ? TemplatePath1Page : TemplatePathMulti;

    /// <summary>분석항목 _meta 의 Parts 컬럼에 "특정32종" 라벨이 들어 있으면 32종 항목으로 분류.
    /// "특정32종 13", "특정32종13", "특정32종 13" 등 공백/포맷 변형 허용.
    /// 끝 숫자를 자연정렬 키로 사용. 메타 매칭 실패해도 항목명 자체로 한번 더 시도.</summary>
    private static readonly System.Text.RegularExpressions.Regex SpecialRe =
        new(@"특정\s*32\s*종\s*(\d+)?", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static (bool isSpecial, int order) ClassifySpecial(
        AnalysisResultRow row, Dictionary<string, AnalysisItem> meta)
    {
        if (row?.항목명 == null) return (false, int.MaxValue);

        string parts = "";
        if (meta.TryGetValue(row.항목명, out var item) && item != null)
            parts = item.Parts ?? "";

        // 1차: meta.Parts. 2차: meta.Category, AliasX 폴백 (사용자가 다른 컬럼에 입력했을 수 있음).
        string[] candidates = item == null
            ? new[] { "" }
            : new[] { parts, item.Category ?? "", item.AliasX ?? "" };

        foreach (var c in candidates)
        {
            if (string.IsNullOrEmpty(c)) continue;
            var m = SpecialRe.Match(c);
            if (!m.Success) continue;
            int n = m.Groups[1].Success && int.TryParse(m.Groups[1].Value, out var v)
                ? v : int.MaxValue;
            return (true, n);
        }
        return (false, int.MaxValue);
    }

    /// <summary>분류 결과 카운트만 반환 — UI 토스트 표시용.</summary>
    public static (int special, int standard) CountClassification(
        List<AnalysisResultRow> rows, Dictionary<string, AnalysisItem> meta)
    {
        int s = 0, t = 0;
        foreach (var r in rows)
            if (ClassifySpecial(r, meta).isSpecial) s++; else t++;
        return (s, t);
    }

    /// <summary>모드별 정렬/필터 적용. ES 번호 정렬은 문자열 자연정렬.</summary>
    private static List<AnalysisResultRow> ApplySortMode(
        List<AnalysisResultRow> rows,
        Dictionary<string, AnalysisItem> meta,
        SortMode mode)
    {
        var classified = rows
            .Select(r => (row: r, cls: ClassifySpecial(r, meta)))
            .ToList();

        int specialCnt = classified.Count(x => x.cls.isSpecial);
        int unmatched  = classified.Count(x => !meta.ContainsKey(x.row.항목명 ?? ""));
        Log($"[정렬] mode={mode}, total={rows.Count}, 특정32종={specialCnt}, 나머지={rows.Count - specialCnt}, _meta매칭실패={unmatched}");

        // 진단: 매칭 실패 항목 일부 + Parts 비교 출력 (32종 0건일 때 원인 추적)
        if (specialCnt == 0 && (mode == SortMode.SpecialFirst || mode == SortMode.SpecialOnly))
        {
            Log("⚠ 특정32종 0건 — 진단:");
            foreach (var c in classified.Take(8))
            {
                bool inMeta = meta.TryGetValue(c.row.항목명 ?? "", out var it);
                Log($"  - 항목명='{c.row.항목명}' _meta매칭={inMeta}, Parts='{(it?.Parts ?? "")}', Category='{(it?.Category ?? "")}'");
            }
        }

        IEnumerable<(AnalysisResultRow row, (bool isSpecial, int order) cls)> filtered = mode switch
        {
            SortMode.SpecialOnly  => classified.Where(x => x.cls.isSpecial),
            SortMode.StandardOnly => classified.Where(x => !x.cls.isSpecial),
            _                     => classified,
        };

        var result = mode switch
        {
            // ES 기준 — ES 문자열 자연정렬 (없으면 항목명).
            SortMode.ByEs =>
                filtered.OrderBy(x => x.row.ES ?? "", StringComparer.Ordinal)
                        .ThenBy(x => x.row.항목명 ?? "", StringComparer.Ordinal),
            // 32종 먼저 (order 순) → 일반 (ES 순)
            SortMode.SpecialFirst =>
                filtered.OrderBy(x => x.cls.isSpecial ? 0 : 1)
                        .ThenBy(x => x.cls.order)
                        .ThenBy(x => x.row.ES ?? "", StringComparer.Ordinal),
            // SpecialOnly: 32종 order 순
            SortMode.SpecialOnly =>
                filtered.OrderBy(x => x.cls.order),
            // StandardOnly: ES 순
            SortMode.StandardOnly =>
                filtered.OrderBy(x => x.row.ES ?? "", StringComparer.Ordinal)
                        .ThenBy(x => x.row.항목명 ?? "", StringComparer.Ordinal),
            _ => filtered,
        };

        return result.Select(x => x.row).ToList();
    }

    /// <summary>시험성적서 항목 정렬/필터 모드.</summary>
    public enum SortMode
    {
        /// <summary>ES 번호 기준 (기본).</summary>
        ByEs,
        /// <summary>분석항목.Parts 가 "특정32종" 으로 시작하는 항목을 앞으로(번호순), 나머지는 ES 기준.</summary>
        SpecialFirst,
        /// <summary>특정32종 항목만 출력.</summary>
        SpecialOnly,
        /// <summary>특정32종 외 나머지 항목만 출력.</summary>
        StandardOnly,
    }

    private const string SpecialPartsPrefix = "특정32종";

    public static string FillAndSave(
        SampleRequest                    sample,
        List<AnalysisResultRow>          rows,
        Dictionary<string, AnalysisItem> meta,
        string reportNo       = "",
        string qualityMgr     = "",
        bool   includeStandard = true,
        bool   openAfter       = true,
        SortMode sortMode     = SortMode.ByEs,
        string fileNameSuffix = "")
    {
        // 템플릿이 없으면 시험기록부 시각언어로 기본 템플릿 자동 생성
        if (!TemplateExists()) GenerateDefaultTemplate();
        // 옛 스키마 템플릿이면 자동으로 새 양식으로 마이그레이션 (사용자 조작 불필요)
        else AutoMigrateTemplates();

        // 정렬/필터 적용 — 원본 rows 는 보존
        rows = ApplySortMode(rows, meta, sortMode);

        Directory.CreateDirectory(OutputDir);

        // 출력 파일명: 채취일자_시료명[_접미사].docx
        var dateStr  = sample.채취일자.Replace("-", "").Replace("/", "");
        var safeName = string.IsNullOrWhiteSpace(sample.시료명) ? sample.약칭 : sample.시료명;
        foreach (var c in Path.GetInvalidFileNameChars())
            safeName = safeName.Replace(c.ToString(), "");
        safeName = safeName.Trim();
        if (safeName.Length > 60) safeName = safeName[..60];
        var fileName = string.IsNullOrEmpty(fileNameSuffix)
            ? $"{dateStr}_{safeName}.docx"
            : $"{dateStr}_{safeName}_{fileNameSuffix}.docx";

        var outPath = Path.Combine(OutputDir, fileName);
        var templatePath = PickTemplate(rows.Count);
        Log($"양식 선택: {(rows.Count <= SinglePageRowThreshold ? "단일페이지" : "다중페이지")} ({rows.Count}건, mode={sortMode}) → {templatePath}");
        File.Copy(templatePath, outPath, overwrite: true);

        var no = string.IsNullOrEmpty(reportNo)
            ? $"WAC-{DateTime.Now:yyyyMMdd}-{sample.약칭}"
            : reportNo;

        var (companyName, representative) = GetContractInfo(sample.약칭);
        var standardMap = BuildStandardMap(rows, sample.방류허용기준);

        bool isQC = sample.정도보증.Trim().ToUpper() == "O"
                 || sample.정도보증.Trim() == "정도보증 적용"
                 || sample.정도보증.Trim() == "Y";

        // 시험성적서 서명란 텍스트 — 호출자 지정 우선 → 사용자 결재정보(설정→결재정보) → Agent 테이블 자동조회.
        // 결재정보에는 "품질책임 수질분야 환경측정분석사 박은지" 같은 풀텍스트를 자유롭게 입력.
        if (string.IsNullOrWhiteSpace(qualityMgr))
            qualityMgr = UserPrefsService.TestReportSignerQualityMgr;
        if (string.IsNullOrWhiteSpace(qualityMgr))
            qualityMgr = GetQualityManagerName();

        // 시료유형 분기 — "접수" / "채수" / "" (미지정 → 채수 취급)
        bool isIntake = string.Equals(sample.시료유형?.Trim(), "접수", StringComparison.Ordinal);
        // 채수 모드 기본 매핑 (기존 채취일자/시료채취자/견적번호 재사용)
        string sampleDate    = isIntake
            ? FormatDate(sample.접수일자)
            : FormatDate(sample.채취일자);
        string samplePerson  = isIntake
            ? (sample.접수담당자 ?? "")
            : $"{sample.시료채취자1} {sample.시료채취자2}".Trim();
        string sampleExtra   = isIntake
            ? (sample.업체담당자 ?? "")
            : (sample.견적번호 ?? "");
        string labelDate    = isIntake ? "접수일자"    : "채수일자";
        string labelPerson  = isIntake ? "담당자"      : "채수담당자";
        string labelExtra   = isIntake ? "업체담당자"  : "채수의뢰서";

        var headerMap = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["성적서번호"] = no,
            ["회사명"]     = string.IsNullOrEmpty(companyName) ? sample.의뢰사업장 : companyName,
            ["대표자"]     = representative,
            ["시료명"]     = sample.시료명 ?? "",
            // 기존 placeholder — 채수 기본 (구버전 템플릿 호환)
            ["채취일자"]   = FormatDate(sample.채취일자),
            ["채취자"]     = $"{sample.시료채취자1} {sample.시료채취자2}".Trim(),
            ["입회자"]     = sample.입회자 ?? "",
            ["분석종료일"] = FormatDate(sample.분석종료일),
            ["용도"]       = isQC ? "정도보증 적용" : "참고용",
            // 비고문구는 결재정보(설정→결재정보) 에서 사용자가 편집 가능. 미설정이면 기본 문구.
            ["비고문구"]   = isQC
                ? UserPrefsService.TestReportRemarkQc
                : UserPrefsService.TestReportRemarkRef,
            ["시험성적서서명"] = qualityMgr ?? "",
            // 시료유형 분기 placeholder (신버전 템플릿 — 라벨/값을 동적으로 출력)
            ["시료유형"]    = isIntake ? "접수" : "채수",
            ["라벨_일자"]   = labelDate,
            ["값_일자"]     = sampleDate,
            ["라벨_담당"]   = labelPerson,
            ["값_담당"]     = samplePerson,
            ["라벨_추가"]   = labelExtra,
            ["값_추가"]     = sampleExtra,
            // 직접 사용 가능한 모드별 placeholder (사용자 커스텀 양식용)
            ["접수일자"]    = FormatDate(sample.접수일자),
            ["접수담당자"]  = sample.접수담당자 ?? "",
            ["업체담당자"]  = sample.업체담당자 ?? "",
            ["채수일자"]    = FormatDate(sample.채취일자),
            ["채수담당자"]  = $"{sample.시료채취자1} {sample.시료채취자2}".Trim(),
            ["채수의뢰서"]  = sample.견적번호 ?? "",
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

            // 3) 본문 치환자 채우기 (반복 후에 실행 — 복제된 행도 같이 처리)
            foreach (var p in body.Descendants<Paragraph>().ToList())
                ReplacePlaceholders(p, headerMap);

            // 3-1) 헤더 파트(들) 도 치환 — {{성적서번호}} 등 헤더 영역 치환자 처리
            foreach (var hp in doc.MainDocumentPart!.HeaderParts)
            {
                var hdrRoot = hp.Header;
                if (hdrRoot == null) continue;
                foreach (var p in hdrRoot.Descendants<Paragraph>().ToList())
                    ReplacePlaceholders(p, headerMap);
                hp.Header.Save();
            }

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
