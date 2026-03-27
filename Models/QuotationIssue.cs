namespace ETA.Models;

/// <summary>견적발행내역 테이블 1행 — rowid 기반</summary>
public class QuotationIssue
{
    /// <summary>rowid (PK 대용)</summary>
    public int     Id       { get; set; }

    // ── 실제 테이블 컬럼 ──────────────────────────────────────────────────
    public string  발행일   { get; set; } = "";   // 견적발행일자
    public string  업체명   { get; set; } = "";
    public string  약칭     { get; set; } = "";
    public string  시료명   { get; set; } = "";   // 트리 자식 노드 표시용
    public string  견적번호 { get; set; } = "";
    public string  견적구분 { get; set; } = "";   // 적용구분
    public string  담당자   { get; set; } = "";   // 견적요청 담당자
    public decimal 총금액   { get; set; }         // 합계 금액

    /// <summary>트리 노드 표시 레이블</summary>
    public string DisplayLabel =>
        string.IsNullOrEmpty(시료명)
            ? $"{업체명}  [{견적번호}]"
            : 시료명;
}
