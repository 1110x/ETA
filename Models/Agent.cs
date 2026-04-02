using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Models;

public class Agent
{
    public string   성명             { get; set; } = string.Empty;
    public string   Original성명     { get; set; } = string.Empty;
    public string   직급             { get; set; } = string.Empty;
    public string   직무             { get; set; } = string.Empty;
    public DateOnly 입사일           { get; set; } = DateOnly.MinValue;
    public string   사번             { get; set; } = string.Empty;
    public string   자격사항         { get; set; } = string.Empty;
    public string   Email            { get; set; } = string.Empty;
    public string   기타             { get; set; } = string.Empty;
    public string   측정인고유번호   { get; set; } = string.Empty;

    // 사진 파일 경로 (Data/Photos/ 기준 상대경로 또는 절대경로)
    public string   PhotoPath        { get; set; } = string.Empty;

    public List<Agent> Children { get; set; } = new();

    // 담당 분석항목/계약업체 (콤마 구분 문자열)
    public string 담당항목 { get; set; } = "";
    public string 담당업체 { get; set; } = "";

    public List<string> 담당항목목록 =>
        담당항목.Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
    public List<string> 담당업체목록 =>
        담당업체.Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim()).Where(s => s.Length > 0).ToList();

    public string 입사일표시 =>
        입사일 == DateOnly.MinValue ? "" : 입사일.ToString("yyyy-MM-dd");
}