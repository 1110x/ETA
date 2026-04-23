using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record WasteCompanyRow(string 프로젝트명, string 업체명, string 관리번호);

public static class WasteCompanyListService
{
    public static List<WasteCompanyRow> GetAll()
    {
        var list = new List<WasteCompanyRow>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT COALESCE(`프로젝트명`,'') AS 프로젝트명,
                       COALESCE(`업체명`,'')    AS 업체명,
                       COALESCE(`관리번호`,'')  AS 관리번호
                FROM   `폐수배출업소`
                ORDER  BY `프로젝트명` ASC, `관리번호` ASC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                list.Add(new WasteCompanyRow(
                    r.GetString(0), r.GetString(1), r.GetString(2)));
            }
        }
        catch { }
        return list;
    }

    public static List<(string Group, List<WasteCompanyRow> Rows)> GetGrouped()
    {
        var all = GetAll();
        return all
            .GroupBy(x => NormalizeGroup(x.프로젝트명))
            .OrderBy(g => GroupOrder(g.Key))
            .Select(g => (g.Key, g.ToList()))
            .ToList();
    }

    private static string NormalizeGroup(string projName)
    {
        if (string.IsNullOrWhiteSpace(projName)) return "기타";
        if (projName.Contains("여수")) return "여수";
        if (projName.Contains("율촌")) return "율촌";
        if (projName.Contains("세풍")) return "세풍";
        return projName;
    }

    private static int GroupOrder(string g) => g switch
    {
        "여수" => 0, "율촌" => 1, "세풍" => 2, _ => 9,
    };
}
