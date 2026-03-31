namespace ETA.Models;

/// <summary>
/// 자료TO측정인 Show2 패널용 행 모델
/// (DataToMeasurerWindow 의 동명 클래스와 별개 — ETA.Models 네임스페이스)
/// </summary>
public class MeasurerRow
{
    public string 항목명       { get; set; } = "";
    public string 법적기준     { get; set; } = "";
    public string 결과값       { get; set; } = "";
    public string 측정방법     { get; set; } = "";
    public string 장비명       { get; set; } = "";
    public string 담당자       { get; set; } = "";
    public string 시작일       { get; set; } = "";
    public string 시작시간     { get; set; } = "09:00";
    public string 종료일       { get; set; } = "";
    public string 종료시간     { get; set; } = "18:00";

    // DataToMeasurerWindow 로 전달할 때 필요한 원본 데이터
    public string 약칭         { get; set; } = "";
    public string 시료명       { get; set; } = "";
    public string 채취일자     { get; set; } = "";
}
