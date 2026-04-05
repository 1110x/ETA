using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.Common;

public record ReagentInfo(
    string ITEM_NO, string 국문명, string 영문명, string 화학식, string CAS번호,
    string 규격, string 단위, string 등급);

public record ReagentDailyRecord(
    int ReagentIndex, DateTime 일자, int 입고, int 출고, int 재고, int 사용중);

public static class RiskExcelService
{
    private static List<ReagentInfo>? _reagentCache;
    private static List<ReagentDailyRecord>? _historyCache;

    public static IReadOnlyList<ReagentInfo> GetReagents()
        => _reagentCache ??= LoadReagentMaster();

    public static IReadOnlyList<ReagentDailyRecord> GetReagentHistory()
        => _historyCache ??= LoadReagentHistory();

    public static void InvalidateCache()
    {
        _reagentCache = null;
        _historyCache = null;
    }

    // ── 시약 마스터 정보 (시약자료 시트 R1~R6) ──────────────────────────────
    private static List<ReagentInfo> LoadReagentMaster()
    {
        var path = FindExcelPath();
        if (!File.Exists(path)) return new();

        try
        {
            using var wb = new XLWorkbook(path);
            if (!wb.TryGetWorksheet("시약자료", out var ws)) return new();

            int maxC = ws.RangeUsed()?.LastColumn().ColumnNumber() ?? 0;
            var list = new List<ReagentInfo>();

            for (int c = 2; c <= maxC; c += 4)
            {
                string itemNo = S(ws, 1, c);
                string name   = S(ws, 3, c);

                // 빈 열 건너뛰기 (5열 간격일 수 있음)
                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(itemNo))
                {
                    if (c + 1 <= maxC)
                    {
                        itemNo = S(ws, 1, c + 1);
                        name   = S(ws, 3, c + 1);
                        if (!string.IsNullOrEmpty(name)) c++;
                        else continue;
                    }
                    else continue;
                }

                list.Add(new ReagentInfo(
                    ITEM_NO: itemNo,
                    국문명:  name,
                    영문명:  S(ws, 2, c),
                    화학식:  S(ws, 4, c),
                    CAS번호: S(ws, 5, c),
                    규격:    S(ws, 6, c),
                    단위:    S(ws, 6, c + 1),
                    등급:    S(ws, 6, c + 2)
                ));
            }
            return list;
        }
        catch { return new(); }
    }

    // ── 일별 입고/출고/재고/사용중 이력 (시약자료 시트 R8~) ─────────────────
    private static List<ReagentDailyRecord> LoadReagentHistory()
    {
        var path = FindExcelPath();
        if (!File.Exists(path)) return new();

        try
        {
            using var wb = new XLWorkbook(path);
            if (!wb.TryGetWorksheet("시약자료", out var ws)) return new();

            int maxC = ws.RangeUsed()?.LastColumn().ColumnNumber() ?? 0;
            int maxR = ws.RangeUsed()?.LastRow().RowNumber() ?? 0;

            // 시약 시작 컬럼 수집
            var reagentCols = new List<int>();
            for (int c = 2; c <= maxC; c += 4)
            {
                string name = S(ws, 3, c);
                string itemNo = S(ws, 1, c);
                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(itemNo))
                {
                    if (c + 1 <= maxC)
                    {
                        name = S(ws, 3, c + 1);
                        if (!string.IsNullOrEmpty(name)) { c++; reagentCols.Add(c); }
                    }
                    continue;
                }
                reagentCols.Add(c);
            }

            var list = new List<ReagentDailyRecord>();

            for (int r = 8; r <= maxR; r++)
            {
                // 날짜 파싱
                var dateVal = ws.Cell(r, 1).Value;
                DateTime date;
                if (dateVal.IsDateTime)
                    date = dateVal.GetDateTime().Date;
                else
                {
                    var ds = dateVal.ToString()?.Trim() ?? "";
                    if (!TryParseKoreanDate(ds, out date)) continue;
                }

                for (int i = 0; i < reagentCols.Count; i++)
                {
                    int c = reagentCols[i];
                    int inV  = N(ws, r, c);
                    int outV = N(ws, r, c + 1);
                    int stk  = N(ws, r, c + 2);
                    int use  = N(ws, r, c + 3);

                    // 전부 0이면 건너뛰기 (데이터 없음)
                    if (inV == 0 && outV == 0 && stk == 0 && use == 0) continue;

                    list.Add(new ReagentDailyRecord(i, date, inV, outV, stk, use));
                }
            }
            return list;
        }
        catch { return new(); }
    }

    // ── 유해자료 시트도 동일 구조 ───────────────────────────────────────────
    public static IReadOnlyList<ReagentInfo> GetHazardousReagents()
    {
        var path = FindExcelPath();
        if (!File.Exists(path)) return new List<ReagentInfo>();

        try
        {
            using var wb = new XLWorkbook(path);
            if (!wb.TryGetWorksheet("유해자료", out var ws)) return new List<ReagentInfo>();

            int maxC = ws.RangeUsed()?.LastColumn().ColumnNumber() ?? 0;
            var list = new List<ReagentInfo>();

            for (int c = 2; c <= maxC; c += 4)
            {
                string itemNo = S(ws, 1, c);
                string name   = S(ws, 3, c);
                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(itemNo))
                {
                    if (c + 1 <= maxC)
                    {
                        itemNo = S(ws, 1, c + 1);
                        name   = S(ws, 3, c + 1);
                        if (!string.IsNullOrEmpty(name)) c++;
                        else continue;
                    }
                    else continue;
                }

                list.Add(new ReagentInfo(
                    ITEM_NO: itemNo,
                    국문명:  name,
                    영문명:  S(ws, 2, c),
                    화학식:  S(ws, 4, c),
                    CAS번호: S(ws, 5, c),
                    규격:    S(ws, 6, c),
                    단위:    S(ws, 6, c + 1),
                    등급:    S(ws, 6, c + 2)
                ));
            }
            return list;
        }
        catch { return new List<ReagentInfo>(); }
    }

    // ── 엑셀 파일 경로 ─────────────────────────────────────────────────────
    private static string FindExcelPath()
    {
        foreach (var root in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            // 새 파일명 우선
            var p1 = Path.Combine(root, "Data", "Templates", "2025 리스크 청하-2.xlsm");
            if (File.Exists(p1)) return p1;
            // 기존 경로 폴백
            var p2 = Path.Combine(root, "Data", "Risk 청하", "2025 리스크 청하.xlsm");
            if (File.Exists(p2)) return p2;
        }
        return Path.Combine(Environment.CurrentDirectory, "Data", "Templates", "2025 리스크 청하-2.xlsm");
    }

    private static string S(IXLWorksheet ws, int r, int c)
        => ws.Cell(r, c).Value.ToString()?.Trim().Replace("\n", " ") ?? "";

    private static int N(IXLWorksheet ws, int r, int c)
    {
        var v = ws.Cell(r, c).Value;
        if (v.IsNumber) return (int)v.GetNumber();
        if (v.IsText && int.TryParse(v.GetText(), out int n)) return n;
        return 0;
    }

    private static bool TryParseKoreanDate(string s, out DateTime date)
    {
        date = default;
        // "2026. 3. 1. 오전 12:00:00" 형식
        try
        {
            var parts = s.Split('.', StringSplitOptions.TrimEntries);
            if (parts.Length >= 3 &&
                int.TryParse(parts[0], out int y) &&
                int.TryParse(parts[1], out int m) &&
                int.TryParse(parts[2], out int d))
            {
                date = new DateTime(y, m, d);
                return true;
            }
        }
        catch { }
        return false;
    }
}
