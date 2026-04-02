namespace ETA.Models;

public class ScheduleEntry
{
    public int    Id       { get; set; }
    public string 날짜     { get; set; } = "";   // YYYY-MM-DD
    public string 직원명   { get; set; } = "";
    public string 직원id   { get; set; } = "";
    public string 분류     { get; set; } = "출장";  // 출장/휴일근무/연차/반차/공가/기타
    public string 사이트   { get; set; } = "";   // 여수/율촌/세풍
    public string 업체약칭 { get; set; } = "";   // 출장 대상 계약업체 약칭
    public string 제목     { get; set; } = "";
    public string 내용     { get; set; } = "";
    public string 시작시간 { get; set; } = "";   // HH:mm
    public string 종료시간 { get; set; } = "";   // HH:mm
    public string 첨부파일 { get; set; } = "";
    public string 등록일시 { get; set; } = "";
    public string 등록자   { get; set; } = "";
}
