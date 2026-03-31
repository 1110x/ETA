using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 자료TO측정인 — 분析결과 행 구성 및 날짜 계산 서비스
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

        string analyst = CurrentUserManager.Instance.CurrentUserId;

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

            rows.Add(new MeasurerRow
            {
                항목명   = analyte,
                법적기준 = "",   // TODO: 방류기준표 연동
                결과값   = result,
                측정방법 = m?.Method     ?? "",
                장비명   = m?.instrument ?? "",
                담당자   = managers.TryGetValue(analyte, out var mgr) && !string.IsNullOrEmpty(mgr) ? mgr : analyst,
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
