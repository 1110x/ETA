namespace ETA.Models;

public class WasteSample
{
    public int    Id        { get; set; }
    public string 채수일    { get; set; } = "";   // YYYY-MM-DD
    public string 구분      { get; set; } = "여수"; // 여수 / 세풍 / 율촌
    public int    순서      { get; set; }
    public string SN        { get; set; } = "";   // 03-31-01 / [세풍]03-31-01
    public string 업체명    { get; set; } = "";
    public string 관리번호  { get; set; } = "";
    public string BOD       { get; set; } = "";
    public string TOC       { get; set; } = "";
    public string SS        { get; set; } = "";
    public string TN        { get; set; } = "";   // T-N
    public string TP        { get; set; } = "";   // T-P
    public string NHexan    { get; set; } = "";   // N-Hexan
    public string Phenols   { get; set; } = "";
    public string CN        { get; set; } = "";   // 시안
    public string CR6       { get; set; } = "";   // 6가크롬
    public string COLOR     { get; set; } = "";   // 색도
    public string ABS       { get; set; } = "";   // ABS
    public string FLUORIDE  { get; set; } = "";   // 불소
    public string 비고      { get; set; } = "";
    public string 확인자    { get; set; } = "";

    // SN 자동 생성: 채수일 + 구분 + 순서
    public static string BuildSN(string 채수일, string 구분, int 순서)
    {
        if (!System.DateTime.TryParse(채수일, out var d)) return "";
        string base_ = $"{d.Month:D2}-{d.Day:D2}-{순서:D2}";
        return 구분 switch
        {
            "세풍" => $"[세풍]{base_}",
            "율촌" => $"[율촌]{base_}",
            _      => base_,   // 여수
        };
    }
}
