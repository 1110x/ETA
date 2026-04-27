using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace ETA.Services.Common;

/// <summary>
/// 양식(엑셀 템플릿) 파일 경로 설정.
/// 사용자가 "설정 → 양식폴더" 메뉴에서 각 양식별 경로를 덮어쓸 수 있고,
/// 미설정 항목은 기본 경로로 폴백. 저장 위치: &lt;repo&gt;/Data/template_paths.json.
/// </summary>
public static class TemplateConfiguration
{
    private const string ConfigFileName = "template_paths.json";

    /// <summary>
    /// 템플릿 슬롯 1개의 메타데이터.
    /// Key=식별자, Label=한글 표시명, DefaultRelative=기본 상대경로, UsedBy=사용 서비스명.
    /// </summary>
    public sealed record Slot(string Key, string Label, string DefaultRelative, string UsedBy);

    /// <summary>전체 템플릿 슬롯 목록 — 설정 UI에 표시되는 순서대로.</summary>
    public static readonly IReadOnlyList<Slot> Slots = new List<Slot>
    {
        // 시험성적서 / 수질측정기록부
        new("TestReport1",          "시험성적서 (1장)",          "Data/Templates/시험성적서/시험성적서1_template.xlsx",       "TestReportPrintService"),
        new("TestReport2",          "시험성적서 (2장)",          "Data/Templates/시험성적서/시험성적서2_template.xlsx",       "TestReportPrintService"),
        new("TestReportIntegrated", "시험성적서 통합(배치)",     "Data/Templates/시험성적서/시험성적서통합.xlsm",            "TestReportPrintService"),
        new("TestReportWord",       "시험성적서 (Word, 레거시)", "Data/Templates/시험성적서/시험성적서_template.docx",       "TestReportWordPrintService"),
        new("TestReportWord1Page",  "시험성적서 Word 단일페이지","Data/Templates/시험성적서/시험성적서_1page_template.docx", "TestReportWordPrintService"),
        new("TestReportWordMulti",  "시험성적서 Word 다중페이지","Data/Templates/시험성적서/시험성적서_multi_template.docx", "TestReportWordPrintService"),
        new("WaterRecord1",         "수질측정기록부 (1장)",      "Data/Templates/시험성적서/수질측정기록부1_template.xlsx",  "WaterMeasurementRecordService"),
        new("WaterRecord2",         "수질측정기록부 (2장)",      "Data/Templates/시험성적서/수질측정기록부2_template.xlsx",  "WaterMeasurementRecordService"),
        // GS 발송양식
        new("GsReport",             "GS 발송양식",               "Data/Templates/GS_발송양식_template.xlsx",                 "GsReportService"),
        // 견적 / 거래명세서
        new("Quotation1",           "견적서 (1장)",              "Data/Templates/견적서/견적서1.xlsx",                        "QuotationService"),
        new("Quotation2",           "견적서 (2장)",              "Data/Templates/견적서/견적서2.xlsx",                        "QuotationService"),
        new("TradeStatement1",      "거래명세서 (1장)",          "Data/Templates/거래명세서/거래명세서1.xlsx",                "QuotationService"),
        new("TradeStatement2",      "거래명세서 (2장)",          "Data/Templates/거래명세서/거래명세서2.xlsx",                "QuotationService"),
        // 기록부 / 업무용
        new("AnalysisRecord",       "분석기록부",                "Data/Templates/분석NOTE/분석기록부.xlsx",                   "AnalysisRecordService"),
        new("TestRecordBookFolder", "기본항목 시험기록부 폴더",  "Data/Templates/기본항목 시험기록부",                       "TestRecordBookService"),
        new("WeeklySchedule",       "요일별 분석항목 정리",      "Data/Templates/요일별 분석항목 정리.xlsx",                 "FacilityResultService"),
        new("RiskAssessment",       "리스크 청하",               "Data/Templates/청하 파일/2025 리스크 청하-2.xlsm",         "RiskExcelService"),
        new("PreDbSnapshot",        "DB 저장 전 스냅샷",         "Data/Templates/PreDbSnapshot_template.xlsm",               "PreDbSnapshotService"),
    };

    // 개발·설치 환경 모두 대응 (AppPaths 참고)
    public static string RepoRoot => AppPaths.RootPath;

    // 설정 파일은 쓰기 가능한 위치 (설치 환경에선 %LOCALAPPDATA%\ETA)
    private static string ConfigPath => Path.Combine(AppPaths.WritableDataRoot, "Data", ConfigFileName);

    private static Dictionary<string, string>? _cache;

    private static Dictionary<string, string> Load()
    {
        if (_cache != null) return _cache;
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                _cache = JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                         ?? new Dictionary<string, string>();
            }
            else _cache = new Dictionary<string, string>();
        }
        catch
        {
            _cache = new Dictionary<string, string>();
        }
        return _cache;
    }

    /// <summary>설정을 JSON 파일로 저장. 캐시도 갱신.</summary>
    public static bool Save(Dictionary<string, string> map)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            var json = JsonSerializer.Serialize(map,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
            _cache = new Dictionary<string, string>(map);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>현재 저장된 오버라이드 맵 (복사본). 설정 UI에서 사용.</summary>
    public static Dictionary<string, string> Snapshot() => new(Load());

    /// <summary>
    /// 키에 대응하는 최종 경로를 반환. 사용자 오버라이드 → 기본 상대경로(RepoRoot 기준) → fallback 순.
    /// 반환 경로는 항상 절대경로. 파일 존재 여부는 호출자가 검사.
    /// </summary>
    public static string Resolve(string key, string? fallbackAbsolute = null)
    {
        var map = Load();
        if (map.TryGetValue(key, out var overridePath) && !string.IsNullOrWhiteSpace(overridePath))
            return Path.IsPathRooted(overridePath) ? overridePath : Path.Combine(RepoRoot, overridePath);

        var slot = Slots.FirstOrDefault(s => s.Key == key);
        if (slot != null)
            return Path.Combine(RepoRoot, slot.DefaultRelative);

        return fallbackAbsolute ?? "";
    }

    /// <summary>설정 UI 강제 재로드용.</summary>
    public static void InvalidateCache() => _cache = null;
}
