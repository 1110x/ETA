namespace ETA.Models;

public class WasteRequest
{
    public int    Id        { get; set; }
    public string 의뢰번호  { get; set; } = "";
    public string 구분      { get; set; } = "";   // 여수 | 율촌 | 세풍
    public string 채취일자  { get; set; } = "";   // yyyy-MM-dd
    public string 업체명    { get; set; } = "";
    public string 관리번호  { get; set; } = "";
    public string 상태      { get; set; } = "대기"; // 대기 | 진행중 | 완료
    public string 등록자    { get; set; } = "";
    public string 등록일시  { get; set; } = "";
}
