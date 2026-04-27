namespace ETA.Models;

public class AnalysisItem
{
    public string Category      { get; set; } = string.Empty;
    public string Analyte       { get; set; } = string.Empty;
    public string 약칭          { get; set; } = string.Empty;
    public string Parts         { get; set; } = string.Empty;
    public int    DecimalPlaces { get; set; }
    public string unit          { get; set; } = string.Empty;     // DB와 일치 (소문자)
    public string ES            { get; set; } = string.Empty;
    public string Method        { get; set; } = string.Empty;
    public string instrument    { get; set; } = string.Empty;    // DB와 일치 (소문자)
    public string AliasX        { get; set; } = string.Empty;    // 쉼표 구분 별칭 목록 (파서 키워드 매핑용)
    public double? 정량한계      { get; set; }                    // LoQ — 결과 < LoQ 면 "ND" 표시
}