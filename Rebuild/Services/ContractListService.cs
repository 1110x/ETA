using System;
using System.Collections.Generic;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record ContractRow(
    string 업체명, string 약칭, string 계약종류,
    string 계약시작, string 계약종료, string 시설구분,
    string 카테고리, string 대표자);

public static class ContractListService
{
    public static List<ContractRow> GetAll()
    {
        var list = new List<ContractRow>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT C_CompanyName, C_Abbreviation, C_ContractType,
                   C_ContractStart, C_ContractEnd, C_FacilityType,
                   C_CategoryType, C_Representative
            FROM   `계약 DB`
            ORDER  BY C_ContractEnd DESC, C_CompanyName ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ContractRow(
                Safe(r, 0), Safe(r, 1), Safe(r, 2),
                Safe(r, 3), Safe(r, 4), Safe(r, 5),
                Safe(r, 6), Safe(r, 7)));
        }
        return list;
    }

    public static bool IsActive(string start, string end)
    {
        if (!DateTime.TryParse(start, out var s)) return false;
        if (!DateTime.TryParse(end,   out var e)) return false;
        var today = DateTime.Today;
        return s.Date <= today && today <= e.Date;
    }

    private static string Safe(System.Data.Common.DbDataReader r, int i)
        => r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
}
