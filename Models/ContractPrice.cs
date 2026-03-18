namespace ETA.Models;

public class ContractPrice
{
    public string ES { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Analyte { get; set; } = string.Empty;
    public decimal? FS100     { get; set; }
    public decimal? FS100Plus { get; set; }
    public decimal? FS56      { get; set; }
    public decimal? NFS56     { get; set; }
    public decimal? FS55      { get; set; }
    public decimal? FS52      { get; set; }
    public decimal? FSHN52    { get; set; }
    public decimal? NFS50     { get; set; }
    public decimal? NFS45     { get; set; }
    public decimal? NFS39     { get; set; }
    public decimal? NFS36     { get; set; }
    public decimal? NFS36RE   { get; set; }
    public decimal? FS25      { get; set; }
}