using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 자료TO측정인 — 분석결과 행 구성 및 날짜 계산 서비스
/// CDP 자동입력은 DataToMeasurerWindow 가 담당.
/// </summary>
public static class DataToMeasurerService
{
    /// <summary>
    /// 선택된 시료 → MeasurerRow 목록 빌드 (Show2 패널 / DataToMeasurerWindow 공용)
    /// </summary>
    public static List<MeasurerRow> BuildRows(SampleRequest sample)
    {
        var rows = new List<MeasurerRow>();

        // fallback: 분장표준처리에 할당이 없는 항목(현장측정 등)은 시료채취자
        // 시료채취자1/2 모두에서 이름을 모아 사번 낮은 1명만 선택
        var candidates = new List<string>();
        foreach (var raw in new[] { sample.시료채취자1 ?? "", sample.시료채취자2 ?? "" })
        {
            foreach (var tok in raw.Split(new[] { ',', '/', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = tok.Trim();
                if (t.Length > 0 && !candidates.Contains(t)) candidates.Add(t);
            }
        }

        List<Agent> allAgents;
        try   { allAgents = AgentService.GetAllItems(); }
        catch { allAgents = new(); }

        string fallbackAnalyst = candidates
            .Select(nm =>
            {
                var ag = allAgents.FirstOrDefault(a =>
                    !string.IsNullOrEmpty(a.성명) &&
                    a.성명.Equals(nm, StringComparison.OrdinalIgnoreCase));
                return (name: nm, sabun: ag?.사번 ?? "");
            })
            .OrderByDescending(x => x.sabun, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.name)
            .FirstOrDefault() ?? "";

        if (string.IsNullOrWhiteSpace(fallbackAnalyst))
            fallbackAnalyst = (sample.입회자 ?? "").Trim();

        Debug.WriteLine($"[DataToMeasurer] sample={sample.시료명} 후보=[{string.Join(",", candidates)}] 입회자='{sample.입회자}' → fallback='{fallbackAnalyst}' (사번 높은 1명)");

        List<AnalysisItem> allItems;
        try   { allItems = AnalysisService.GetAllItems(); }
        catch { allItems = new(); }
        var meta = allItems.ToDictionary(x => x.Analyte, x => x, StringComparer.OrdinalIgnoreCase);

        Dictionary<string, string>              managers;
        Dictionary<string, (int days, string)>  stdDays;
        try   { managers = AnalysisRequestService.GetManagersByDate(sample.채취일자); }
        catch { managers = new(); }
        try   { stdDays = AnalysisRequestService.GetStandardDaysInfo(); }
        catch { stdDays = new(); }

        // 방류기준표: 구분(=sample.방류허용기준) 행의 {항목명: 기준값}
        string 기준유형 = (sample.방류허용기준 ?? "").Trim();
        Dictionary<string, string> stdMap;
        try   { stdMap = TestReportService.GetStandardValueMap(기준유형); }
        catch { stdMap = new(StringComparer.OrdinalIgnoreCase); }
        Debug.WriteLine($"[DataToMeasurer] 기준유형='{기준유형}' stdMap.Count={stdMap.Count}");

        // 현장측정 항목(분장 할당 없이 시료채취자 사용): pH / 수소이온 / 온도 / 수온 / DO / 용존산소 / 잔류염소 / 전기전도도 / 투명도 / 탁도
        static bool IsFieldMeasurement(string analyte)
        {
            var n = (analyte ?? "").Replace(" ", "").Replace("-", "").ToLowerInvariant();
            if (string.IsNullOrEmpty(n)) return false;
            return n.Contains("수소이온") || n.StartsWith("ph") || n == "ph" || n.EndsWith("ph")
                || n.Contains("온도") || n.Contains("수온")
                || n.Contains("용존산소") || n.StartsWith("do")
                || n.Contains("잔류염소") || n.Contains("전기전도") || n.Contains("투명도") || n.Contains("탁도");
        }

        foreach (var kv in sample.분석결과)
        {
            var analyte = kv.Key;
            var result  = kv.Value;
            if (string.IsNullOrWhiteSpace(result)) continue;

            meta.TryGetValue(analyte, out var m);

            string 종료일 = !string.IsNullOrWhiteSpace(sample.분석종료일)
                ? sample.분석종료일
                : (stdDays.TryGetValue(analyte, out var si) && si.days > 0
                    ? AddBusinessDays(sample.채취일자, si.days)
                    : sample.채취일자);

            bool isField = IsFieldMeasurement(analyte);
            string 담당자;
            string src;
            if (isField)
            {
                담당자 = fallbackAnalyst; src = "field-fallback";
            }
            else if (managers.TryGetValue(analyte, out var mgr) && !string.IsNullOrWhiteSpace(mgr))
            {
                담당자 = mgr.Trim(); src = "분장";
            }
            else
            {
                담당자 = fallbackAnalyst; src = "no-분장-fallback";
            }
            Debug.WriteLine($"[DataToMeasurer] analyte='{analyte}' field={isField} src={src} 담당자='{담당자}'");

            string 법적기준 = "";
            if (stdMap.TryGetValue(analyte.Trim(), out var v) && !string.IsNullOrWhiteSpace(v))
                법적기준 = v.Trim();
            else
            {
                var normA = analyte.Replace("-", "").Replace(" ", "").ToLowerInvariant();
                foreach (var skv in stdMap)
                {
                    var normK = skv.Key.Replace("-", "").Replace(" ", "").ToLowerInvariant();
                    if (normK == normA && !string.IsNullOrWhiteSpace(skv.Value)) { 법적기준 = skv.Value.Trim(); break; }
                }
            }

            rows.Add(new MeasurerRow
            {
                항목명   = analyte,
                법적기준 = 법적기준,
                결과값   = result,
                측정방법 = m?.Method     ?? "",
                장비명   = m?.instrument ?? "",
                담당자   = 담당자,
                시작일   = sample.채취일자,
                시작시간 = "09:00",
                종료일   = 종료일,
                약칭     = sample.약칭,
                시료명   = sample.시료명,
                채취일자 = sample.채취일자,
            });
        }

        return rows;
    }

    public static string AddBusinessDays(string startDateStr, int days)
    {
        if (!DateTime.TryParse(startDateStr, out var date) || days <= 0) return startDateStr;
        int added = 0;
        while (added < days)
        {
            date = date.AddDays(1);
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                added++;
        }
        return date.ToString("yyyy-MM-dd");
    }
}
