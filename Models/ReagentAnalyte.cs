namespace ETA.Models;

/// <summary>시약 ↔ 분석항목 연결 — 시료당 소요량 정보</summary>
public class ReagentAnalyte
{
    public int    Id           { get; set; }
    public int    시약Id       { get; set; }
    public string 분석항목     { get; set; } = "";   // 분장표준처리 컬럼명 (fullName)
    public double 시료당소요량  { get; set; } = 0;    // 시료 1건당 소요량 (단위: 시약의 단위와 동일)
}
