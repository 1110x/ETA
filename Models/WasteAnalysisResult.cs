namespace ETA.Models;

public class WasteAnalysisResult
{
    public string  채수일    { get; set; } = "";
    public double? BOD      { get; set; }
    public double? TOC_TCIC { get; set; }   // TOC(TC-IC)
    public double? TOC_NPOC { get; set; }   // TOC(NPOC)
    public double? SS       { get; set; }
    public double? TN       { get; set; }
    public double? TP       { get; set; }
    public double? Phenols  { get; set; }
    public double? NHexan   { get; set; }
}
