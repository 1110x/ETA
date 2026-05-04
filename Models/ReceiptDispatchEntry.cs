namespace ETA.Models;

/// <summary>측정대행 접수/발송 대장 1행 — 수질분석센터_결과 6컬럼 뷰</summary>
public class ReceiptDispatchEntry
{
    public int    Id          { get; set; }
    public string 접수번호    { get; set; } = "";   // 견적번호 재활용
    public string 접수일      { get; set; } = "";   // 채취일자
    public string 시료명      { get; set; } = "";
    public string 업체명      { get; set; } = "";   // 의뢰사업장
    public string 약칭        { get; set; } = "";
    public string 분석항목    { get; set; } = "";   // O 표시된 분석항목 콤마결합
    public string 발송일      { get; set; } = "";

    public string 의뢰인및업체명 =>
        string.IsNullOrEmpty(약칭) ? 업체명
        : string.IsNullOrEmpty(업체명) ? 약칭
        : $"{업체명} ({약칭})";
}
