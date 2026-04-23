using System;
using System.Collections.Generic;
using System.Text.Json;
using ETA.Services.Common;
using ETA.Services.SERVICE4;

namespace ETA.Services.SERVICE1;

/// <summary>
/// 분석정보.분석조건 컬럼(TEXT, JSON Key/Value 배열) 로드/저장.
/// 시험기록부형 UI(수질분석센터) 상단 블록에서 Analyte 단위로 Lookup.
/// </summary>
public static class AnalysisConditionService
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = false,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>분석항목(Analyte)에 저장된 분석조건 목록 로드.</summary>
    public static List<AnalysisCondition> Load(string analyte)
    {
        var list = new List<AnalysisCondition>();
        if (string.IsNullOrWhiteSpace(analyte)) return list;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(`분석조건`, '') FROM `분석정보` WHERE `Analyte`=@a LIMIT 1";
        var p = cmd.CreateParameter(); p.ParameterName = "@a"; p.Value = analyte;
        cmd.Parameters.Add(p);

        var obj = cmd.ExecuteScalar();
        var json = obj?.ToString() ?? "";
        if (string.IsNullOrWhiteSpace(json)) return list;

        try
        {
            var parsed = JsonSerializer.Deserialize<List<AnalysisCondition>>(json, JsonOpts);
            if (parsed != null) list.AddRange(parsed);
        }
        catch { /* 구형/손상 JSON은 빈 목록 */ }
        return list;
    }

    /// <summary>분석항목(Analyte)에 분석조건 목록 저장 (전체 덮어쓰기).</summary>
    public static void Save(string analyte, IEnumerable<AnalysisCondition> conditions)
    {
        if (string.IsNullOrWhiteSpace(analyte)) return;
        var json = JsonSerializer.Serialize(conditions ?? Array.Empty<AnalysisCondition>(), JsonOpts);

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `분석정보` SET `분석조건`=@j WHERE `Analyte`=@a";
        var pj = cmd.CreateParameter(); pj.ParameterName = "@j"; pj.Value = json; cmd.Parameters.Add(pj);
        var pa = cmd.CreateParameter(); pa.ParameterName = "@a"; pa.Value = analyte; cmd.Parameters.Add(pa);
        cmd.ExecuteNonQuery();
    }

    /// <summary>모든 분석항목의 Analyte 목록 (드롭다운용).</summary>
    public static List<string> GetAnalyteList()
    {
        var list = new List<string>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT `Analyte` FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte`<>'' ORDER BY `ES` ASC";
        using var rd = cmd.ExecuteReader();
        while (rd.Read())
        {
            var a = rd.IsDBNull(0) ? "" : rd.GetString(0);
            if (!string.IsNullOrWhiteSpace(a)) list.Add(a);
        }
        return list;
    }
}
