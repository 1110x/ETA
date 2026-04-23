using System.Collections.Generic;
using ETA.Services.Common;

namespace ETA.Rebuild.Services;

public record AliasMappingRow(string 원문, string 약칭, string 소스);

public class AliasSnapshot
{
    public List<string> Standards { get; set; } = new();
    public List<AliasMappingRow> Mappings { get; set; } = new();
}

public static class CompoundAliasService
{
    public static AliasSnapshot GetSnapshot()
    {
        var snap = new AliasSnapshot();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT COALESCE(`별명`,'')     AS 원문,
                       COALESCE(`분석항목`,'') AS 약칭,
                       COALESCE(`표준코드`,'') AS 소스
                FROM   `화합물별명`
                ORDER  BY `분석항목` ASC, `별명` ASC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var alias = r.GetString(0);
                var std = r.GetString(1);
                var code = r.GetString(2);
                snap.Mappings.Add(new AliasMappingRow(alias, std, code));
            }
        }
        catch { }

        var seen = new HashSet<string>();
        foreach (var m in snap.Mappings)
        {
            if (string.IsNullOrWhiteSpace(m.약칭)) continue;
            if (seen.Add(m.약칭)) snap.Standards.Add(m.약칭);
        }
        return snap;
    }
}
