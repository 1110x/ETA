using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ETA.Services.Common;

public record ReagentInfo(string 국문명, string 영문명, string 화학식, string CAS번호);

public static class RiskExcelService
{
    private static List<ReagentInfo>? _cache;

    public static IReadOnlyList<ReagentInfo> GetReagents()
        => _cache ??= LoadReagents();

    public static void InvalidateCache() => _cache = null;

    private static List<ReagentInfo> LoadReagents()
    {
        var path = FindExcelPath();
        if (!File.Exists(path)) return new();

        try
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet("시약자료");
            var seen = new HashSet<string>();
            var list = new List<ReagentInfo>();
            int maxCol = ws.RangeUsed()?.LastColumnUsed()?.ColumnNumber() ?? 600;

            for (int c = 2; c <= maxCol; c++)
            {
                string name = ws.Cell(3, c).Value.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(name) || name == "국문") continue;
                if (!seen.Add(name)) continue;

                list.Add(new ReagentInfo(
                    국문명: name.Replace("\n", " "),
                    영문명: ws.Cell(2, c).Value.ToString()?.Trim().Replace("\n", " ") ?? "",
                    화학식: ws.Cell(4, c).Value.ToString()?.Trim() ?? "",
                    CAS번호: ws.Cell(5, c).Value.ToString()?.Trim() ?? ""
                ));
            }
            return list;
        }
        catch { return new(); }
    }

    private static string FindExcelPath()
    {
        foreach (var root in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            var p = Path.Combine(root, "Data", "Risk 청하", "2025 리스크 청하.xlsm");
            if (File.Exists(p)) return p;
        }
        return Path.Combine(Environment.CurrentDirectory, "Data", "Risk 청하", "2025 리스크 청하.xlsm");
    }
}
