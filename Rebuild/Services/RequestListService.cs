using System.Collections.Generic;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record RequestRow(
    string 견적번호, string 채취일자, string 시료명,
    string 약칭, string 정도보증);

public static class RequestListService
{
    public static List<RequestRow> GetRecent(int limit = 200)
    {
        var list = new List<RequestRow>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT COALESCE(`견적번호`, '') AS 견적번호,
                   COALESCE(`채취일자`, '') AS 채취일자,
                   COALESCE(`시료명`,   '') AS 시료명,
                   COALESCE(`약칭`,     '') AS 약칭,
                   COALESCE(`정도보증`, '') AS 정도보증
            FROM   `수질분석센터_결과`
            ORDER  BY `채취일자` DESC
            LIMIT  {limit}";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new RequestRow(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4)));
        }
        return list;
    }
}

public record ReportGroup(string 견적번호, string 채취일자, string 약칭, int 시료수);

public static class ReportListService
{
    public static List<ReportGroup> GetGroups(int limit = 120)
    {
        var list = new List<ReportGroup>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT COALESCE(`견적번호`,'')  AS 견적번호,
                   MAX(COALESCE(`채취일자`,'')) AS 채취일자,
                   MAX(COALESCE(`약칭`,''))     AS 약칭,
                   COUNT(*)                     AS 시료수
            FROM   `수질분석센터_결과`
            WHERE  COALESCE(`견적번호`,'') <> ''
            GROUP  BY `견적번호`
            ORDER  BY MAX(`채취일자`) DESC
            LIMIT  {limit}";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ReportGroup(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? 0  : System.Convert.ToInt32(r.GetValue(3))));
        }
        return list;
    }
}
