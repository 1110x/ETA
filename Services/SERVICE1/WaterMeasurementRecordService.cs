using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 수질측정기록부 Excel 생성
/// VBA 매크로 '수질측정기록부자료' / '법정양식' 로직 포팅
/// 템플릿 1 (≤22개 항목, 1페이지) / 템플릿 2 (>22개, 2페이지)
/// </summary>
public static class WaterMeasurementRecordService
{
    private const int ItemThreshold = 22;
    private static readonly string Template1 =
        Path.Combine("Data", "Templates", "수질측정기록부1_template.xlsx");
    private static readonly string Template2 =
        Path.Combine("Data", "Templates", "수질측정기록부2_template.xlsx");

    public static bool TemplatesExist() => File.Exists(Template1) && File.Exists(Template2);

    /// <summary>Excel 파일 생성 후 경로 반환</summary>
    public static string Generate(SampleRequest sample,
                                  List<AnalysisResultRow> rows,
                                  Dictionary<string, AnalysisItem>? meta = null)
    {
        if (sample == null) throw new ArgumentNullException(nameof(sample));

        // 항목 수 기준 템플릿 선택
        var itemCount = rows?.Count ?? 0;
        var templatePath = itemCount <= ItemThreshold ? Template1 : Template2;
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"템플릿 없음: {templatePath}");

        using var wb = new XLWorkbook(templatePath);
        var ws = wb.Worksheet("수질측정기록부");

        // ── 업체 정보 조회 (계약 DB) ─────────────────────────────
        var company = sample.의뢰사업장 ?? "";
        var contracts = ContractService.GetAllContracts();
        var contract = contracts.FirstOrDefault(c =>
            string.Equals(c.C_CompanyName, company, StringComparison.OrdinalIgnoreCase));

        // ① 의뢰인
        ws.Cell(2, 4).Value = company;                           // 상호
        ws.Cell(3, 4).Value = contract?.C_Address ?? "";         // 소재지
        ws.Cell(4, 4).Value = contract?.C_Representative ?? ""; // 대표자
        ws.Cell(5, 4).Value = sample.입회자 ?? "";               // 환경기술인 (입회자)

        // ② 일반항목
        ws.Cell(2, 9).Value = contract?.C_FacilityType ?? "";    // 시설별
        ws.Cell(3, 9).Value = contract?.C_CategoryType ?? "";    // 종류별
        ws.Cell(4, 9).Value = contract?.C_MainProduct ?? "";     // 주생산품

        // ③ 의뢰용도 — Agent 존재 여부로 결정
        var agentNames = AgentService.GetAllNames();
        var agentSet = new HashSet<string>(agentNames, StringComparer.OrdinalIgnoreCase);
        bool sampler1IsAgent = !string.IsNullOrEmpty(sample.시료채취자1)
                            && agentSet.Contains(sample.시료채취자1.Trim());
        bool sampler2IsAgent = !string.IsNullOrEmpty(sample.시료채취자2)
                            && agentSet.Contains(sample.시료채취자2.Trim());
        bool anyAgent = sampler1IsAgent || sampler2IsAgent;
        ws.Cell(6, 4).Value = anyAgent ? "제출 또는 보고용" : "참고용";

        // ③ 대상 명칭, 의뢰항목
        ws.Cell(7, 4).Value = sample.시료명 ?? "";
        var items = rows?.Select(r => r.항목명).Where(n => !string.IsNullOrEmpty(n)).ToList()
                    ?? new List<string>();
        if (items.Count > 0)
        {
            var firstItem = items[0];
            var extra = items.Count - 1;
            ws.Cell(8, 4).Value = extra > 0 ? $"{firstItem} 외 {extra}건" : firstItem;
        }

        // ④ 채취일시 / 시료채취자
        ws.Cell(11, 4).Value = sample.채취일자 ?? "";
        ws.Cell(12, 4).Value = sample.채취시간 ?? "";
        if (!string.IsNullOrEmpty(sample.시료채취자1))
            ws.Cell(11, 8).Value = $"{sample.시료채취자1}      (서명)";
        if (!string.IsNullOrEmpty(sample.시료채취자2))
            ws.Cell(12, 8).Value = $"{sample.시료채취자2}      (서명)";

        // ⑤ 측정분석 결과 하단 - 분석기간 & 날짜
        // 시작일 = 채수일, 종료일 = 각 항목 시험기록부의 분석일 MAX
        var 채취일자 = sample.채취일자 ?? "";
        var 분석종료일 = FindMaxAnalysisDate(sample);
        if (string.IsNullOrEmpty(분석종료일)) 분석종료일 = sample.분석종료일 ?? "";

        if (!string.IsNullOrEmpty(채취일자) || !string.IsNullOrEmpty(분석종료일))
            ws.Cell(79, 4).Value = $"{채취일자} ~ {분석종료일}";

        if (!string.IsNullOrEmpty(분석종료일) && 분석종료일 != "분석중"
            && DateTime.TryParse(분석종료일, out var dt))
        {
            ws.Cell(84, 1).Value = $"{dt:yyyy년 MM월 dd일}";
        }

        // ⑤ 측정분석 결과 (14행부터)
        // 방류기준 캐시 (시료의 방류허용기준 기준)
        var stdMap = TestReportService.GetStandardValueMap(sample.방류허용기준 ?? "");

        int rowIdx = 14;
        int maxRow = itemCount <= ItemThreshold ? 35 : 78;
        foreach (var r in rows ?? Enumerable.Empty<AnalysisResultRow>())
        {
            if (rowIdx > maxRow) break;
            var itemLabel = string.IsNullOrEmpty(r.단위) ? r.항목명 : $"{r.항목명}({r.단위})";
            ws.Cell(rowIdx, 2).Value = itemLabel;
            var stdVal = stdMap.TryGetValue(r.항목명.Trim(), out var sv) && !string.IsNullOrEmpty(sv)
                ? sv : "해당없음";
            ws.Cell(rowIdx, 4).Value = stdVal;
            ws.Cell(rowIdx, 6).Value = r.결과값 ?? "";
            string method = "";
            if (meta != null && meta.TryGetValue(r.항목명, out var m))
                method = m.ES ?? "";
            if (string.IsNullOrEmpty(method)) method = r.분석방법 ?? "";
            ws.Cell(rowIdx, 8).Value = method;
            rowIdx++;
        }

        // 저장
        var outDir = Path.Combine("Data", "Exports");
        Directory.CreateDirectory(outDir);
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var safeSample = SanitizeFileName(sample.시료명 ?? "sample");
        var outPath = Path.Combine(outDir, $"수질측정기록부_{safeSample}_{ts}.xlsx");
        wb.SaveAs(outPath);
        return outPath;
    }

    /// <summary>
    /// 각 분석항목의 *_시험기록부에서 해당 시료의 분석일 MAX를 구함
    /// </summary>
    private static string FindMaxAnalysisDate(SampleRequest sample)
    {
        if (sample == null || sample.분석결과 == null || sample.분석결과.Count == 0)
            return "";
        var 시료명 = sample.시료명 ?? "";
        if (string.IsNullOrWhiteSpace(시료명)) return "";

        DateTime? maxDate = null;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            foreach (var 항목 in sample.분석결과.Keys)
            {
                var tableName = $"{항목}_시험기록부";
                if (!DbConnectionFactory.TableExists(conn, tableName)) continue;

                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"SELECT MAX(`분석일`) FROM `{tableName}` WHERE TRIM(`시료명`)=@s";
                cmd.Parameters.AddWithValue("@s", 시료명.Trim());
                try
                {
                    var result = cmd.ExecuteScalar();
                    if (result == null || result == DBNull.Value) continue;
                    var s = result.ToString() ?? "";
                    if (DateTime.TryParse(s, out var dt))
                    {
                        if (maxDate == null || dt > maxDate) maxDate = dt;
                    }
                }
                catch { /* 컬럼 없는 테이블은 스킵 */ }
            }
        }
        catch { }

        return maxDate.HasValue ? maxDate.Value.ToString("yyyy-MM-dd") : "";
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c));
        return safe.Length > 50 ? safe.Substring(0, 50) : safe;
    }
}
