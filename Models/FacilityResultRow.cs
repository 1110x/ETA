namespace ETA.Models;

public class FacilityResultRow
{
    public int    Id        { get; set; }   // 0 = 미저장
    public int    마스터Id   { get; set; }
    public string 시료명     { get; set; } = "";
    public string 비고마스터  { get; set; } = "";  // 처리시설_마스터.비고 (주1회 목요일 등)

    // 활성 항목 (처리시설_마스터에서 "O"/"O(MLSS)" 등인 항목)
    public bool BOD활성      { get; set; }
    public bool TOC활성      { get; set; }
    public bool SS활성       { get; set; }
    public bool TN활성       { get; set; }
    public bool TP활성       { get; set; }
    public bool 총대장균군활성 { get; set; }
    public bool COD활성      { get; set; }
    public bool 염소이온활성   { get; set; }
    public bool 영양염류활성   { get; set; }
    public bool 함수율활성    { get; set; }
    public bool 중금속활성    { get; set; }

    // 측정 결과값
    public string BOD      { get; set; } = "";
    public string TOC      { get; set; } = "";
    public string SS       { get; set; } = "";
    public string TN       { get; set; } = "";
    public string TP       { get; set; } = "";
    public string 총대장균군 { get; set; } = "";
    public string COD      { get; set; } = "";
    public string 염소이온  { get; set; } = "";
    public string 영양염류  { get; set; } = "";
    public string 함수율   { get; set; } = "";
    public string 중금속   { get; set; } = "";
    public string 비고     { get; set; } = "";
}
