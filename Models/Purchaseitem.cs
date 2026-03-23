using System;
 
namespace ETA.Models;
 
public class PurchaseItem
{
    public int    Id       { get; set; }           // 번호 (DB auto)
    public string 구분     { get; set; } = "";     // 소모품 / 장비 / 기타
    public string 품목     { get; set; } = "";
    public int    수량     { get; set; } = 1;
    public string 비고     { get; set; } = "";
    public string 요청자   { get; set; } = "";
    public DateTime 요청일 { get; set; } = DateTime.Today;
    public string 상태     { get; set; } = "대기"; // 대기 / 승인 / 완료 / 반려
}
 