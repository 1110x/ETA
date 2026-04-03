namespace ETA.Models;

/// <summary>시약 월별 입출고 이력</summary>
public class ReagentHistory
{
    public int    Id       { get; set; }
    public int    시약Id   { get; set; }
    public string 일자     { get; set; } = "";  // yyyy-MM-dd (월말 기준)
    public int    입고     { get; set; } = 0;
    public int    출고     { get; set; } = 0;
    public int    재고     { get; set; } = 0;
    public int    사용중   { get; set; } = 0;
}
