namespace ETA.Models;
using System;
using ETA.Services;

public class Contract
{
    // ── 원본 키 (UPDATE/DELETE WHERE 절용) ───────────────────────────────────
    public string OriginalCompanyName        { get; set; } = string.Empty;

    // ── 필드 ─────────────────────────────────────────────────────────────────
    public string    C_CompanyName              { get; set; } = string.Empty;
    public DateTime? C_ContractStart            { get; set; }
    public DateTime? C_ContractEnd              { get; set; }
    public int?      C_ContractDays             { get; set; }
    public decimal?  C_ContractAmountVATExcluded{ get; set; }
    public string    C_Abbreviation             { get; set; } = string.Empty;
    public string    C_ContractType             { get; set; } = string.Empty;
    public string    C_Address                  { get; set; } = string.Empty;
    public string    C_Representative           { get; set; } = string.Empty;
    public string    C_FacilityType             { get; set; } = string.Empty;
    public string    C_CategoryType             { get; set; } = string.Empty;
    public string    C_MainProduct              { get; set; } = string.Empty;
    public string    C_ContactPerson            { get; set; } = string.Empty;
    public string    C_PhoneNumber              { get; set; } = string.Empty;
    public string    C_Email                    { get; set; } = string.Empty;

    // ── 표시용 헬퍼 ──────────────────────────────────────────────────────────
    public bool   HasAbbr   => !string.IsNullOrWhiteSpace(C_Abbreviation);
    public string BadgeBg   => BadgeColorHelper.GetBadgeColor(C_Abbreviation).Bg;
    public string BadgeFg   => BadgeColorHelper.GetBadgeColor(C_Abbreviation).Fg;
    public string C_ContractStartStr =>
        C_ContractStart.HasValue ? C_ContractStart.Value.ToString("yyyy-MM-dd") : "";
    public string C_ContractEndStr =>
        C_ContractEnd.HasValue ? C_ContractEnd.Value.ToString("yyyy-MM-dd") : "";
    public string C_ContractAmountStr =>
        C_ContractAmountVATExcluded.HasValue
            ? C_ContractAmountVATExcluded.Value.ToString("N0") + " 원"
            : "";
}