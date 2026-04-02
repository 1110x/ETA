namespace ETA.Models;

/// <summary>기타업무 모델 — 직원에게 배정하는 일반 업무</summary>
public class MiscTask
{
    public int    Id       { get; set; }
    public string 업무명   { get; set; } = "";
    public string 내용     { get; set; } = "";
    public string 배정자   { get; set; } = "";   // 배정한 사람 사번
    public string 담당자id { get; set; } = "";   // 배정받은 직원 사번
    public string 담당자명 { get; set; } = "";
    public string 상태     { get; set; } = "대기";  // 대기/진행/완료
    public string 마감일   { get; set; } = "";   // yyyy-MM-dd
    public string 등록일시 { get; set; } = "";
    public string 완료일시 { get; set; } = "";
}
