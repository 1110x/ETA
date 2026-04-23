using System.Collections.Generic;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record StaffRow(string 사번, string 성명, string 부서, string 상태);

public static class StaffListService
{
    public static List<StaffRow> GetAll()
    {
        var list = new List<StaffRow>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `사번`, `성명`, `부서`, `상태` FROM `Agent` ORDER BY `사번` ASC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new StaffRow(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3)));
        }
        return list;
    }
}
