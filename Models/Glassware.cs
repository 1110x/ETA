namespace ETA.Models;

/// <summary>초자 — 실험실 초자류 재고 관리</summary>
public class Glassware
{
    public int     Id     { get; set; }
    public string  품목명 { get; set; } = "";
    public string  용도   { get; set; } = "";
    public string  규격   { get; set; } = "";   // 예) 100mL, 250mL
    public string  재질   { get; set; } = "유리"; // 유리/플라스틱/기타
    public int     수량   { get; set; } = 0;
    public decimal 단가   { get; set; } = 0;
    public string  비고   { get; set; } = "";
    public string  등록일 { get; set; } = "";   // yyyy-MM-dd
    public string  상태   { get; set; } = "정상"; // 정상/파손/폐기
}
