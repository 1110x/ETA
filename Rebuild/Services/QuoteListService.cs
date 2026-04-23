using System.Collections.Generic;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record QuoteRow(
    string 견적번호, string 견적발행일자, string 업체명,
    string 약칭, string 시료명, string 담당자,
    decimal 합계금액);

public static class QuoteListService
{
    public static List<QuoteRow> GetAll(int limit = 200)
    {
        var list = new List<QuoteRow>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT COALESCE(`견적번호`,'')       AS 견적번호,
                   COALESCE(`견적발행일자`,'')   AS 견적발행일자,
                   COALESCE(`업체명`,'')         AS 업체명,
                   COALESCE(`약칭`,'')           AS 약칭,
                   COALESCE(`시료명`,'')         AS 시료명,
                   COALESCE(`담당자`,'')         AS 담당자,
                   COALESCE(`합계 금액`, 0)      AS 합계금액
            FROM   `견적발행내역`
            ORDER  BY `견적발행일자` DESC, `견적번호` DESC
            LIMIT  {limit}";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            decimal amount = 0;
            try { amount = r.IsDBNull(6) ? 0 : System.Convert.ToDecimal(r.GetValue(6)); } catch { }
            list.Add(new QuoteRow(
                r.IsDBNull(0) ? "" : r.GetString(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4),
                r.IsDBNull(5) ? "" : r.GetString(5),
                amount));
        }
        return list;
    }
}
