namespace ETA.Models;

/// <summary>폐수_의뢰_항목: 의뢰 1건당 분석항목 1개 = 행 1개</summary>
public class WasteRequestItem
{
    public int    Id        { get; set; }
    public int    의뢰Id    { get; set; }
    // 의뢰 JOIN 필드 (조회 편의)
    public string 의뢰번호  { get; set; } = "";
    public string 구분      { get; set; } = "";
    public string 업체명    { get; set; } = "";
    public string 채취일자  { get; set; } = "";
    // 항목별 상태
    public string 항목      { get; set; } = "";   // BOD, TOC, SS, TN, TP, ...
    public string 상태      { get; set; } = "미담"; // 미담 | 담음 | 완료
    public string 배정자    { get; set; } = "";
    public string 배정일시  { get; set; } = "";
    public string 완료일시  { get; set; } = "";
}

/// <summary>처리시설_작업: 시료명 단위 작업 배정 현황</summary>
public class FacilityWorkItem
{
    public int    Id        { get; set; }
    public int    마스터Id  { get; set; }
    public string 채취일자  { get; set; } = "";
    public string 시설명    { get; set; } = "";
    public string 시료명    { get; set; } = "";
    public string 항목목록  { get; set; } = "";   // 쉼표 구분 활성 항목
    public string 상태      { get; set; } = "미담";
    public string 배정자    { get; set; } = "";
    public string 배정일시  { get; set; } = "";
    public string 완료일시  { get; set; } = "";
    public string 비고마스터 { get; set; } = "";  // 처리시설_마스터.비고 (주기 표시)
}
