using System;

namespace ETA.Models;

public class RepairItem
{
    public int       Id       { get; set; }
    public string    구분     { get; set; } = "";   // 장비/시설/차량/IT/기타
    public string    장비명   { get; set; } = "";
    public string    증상     { get; set; } = "";
    public string    위치     { get; set; } = "";
    public string    요청자   { get; set; } = "";
    public DateTime  요청일   { get; set; } = DateTime.Today;
    public DateTime? 완료예정일 { get; set; }
    public string    처리내용 { get; set; } = "";
    public string    비고     { get; set; } = "";
    public string    상태     { get; set; } = "대기";

    public string 상태배지 => 상태 switch
    {
        "대기"   => "⏳",
        "진행중" => "🔧",
        "완료"   => "✅",
        "반려"   => "❌",
        _        => "❓",
    };
}
