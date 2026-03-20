using System;
using System.Collections.Generic;

namespace ETA.Models;

public class Agent
{
    public string 성명 { get; set; } = string.Empty;
    public string Original성명 { get; set; } = string.Empty;
    public string 직급 { get; set; } = string.Empty;
    public string 직무 { get; set; } = string.Empty;
    public DateOnly 입사일 { get; set; } = DateOnly.MinValue;

    public string 사번 { get; set; } = string.Empty;


   

    public string 자격사항 { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string 기타 { get; set; } = string.Empty;
    public string 측정인고유번호 { get; set; } = string.Empty;

    public List<Agent> Children { get; set; } = new();

    public string 입사일표시 =>
        입사일 == DateOnly.MinValue ? "" : 입사일.ToString("yyyy-MM-dd");
}