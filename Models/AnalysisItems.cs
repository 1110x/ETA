namespace ETA.Models;

public class AnalysisItem
{
    public string Category      { get; set; } = string.Empty;
    public string Analyte       { get; set; } = string.Empty;
    public string Parts         { get; set; } = string.Empty;
    public int    DecimalPlaces { get; set; }
    public string unit          { get; set; } = string.Empty;     // DB와 일치 (소문자)
    public string ES            { get; set; } = string.Empty;
    public string Method        { get; set; } = string.Empty;
    public string instrument    { get; set; } = string.Empty;    // DB와 일치 (소문자)
}