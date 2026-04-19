namespace ETA.Models;
using System;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

public class Contract
{
    // ── 원본 키 (UPDATE/DELETE WHERE 절용) ───────────────────────────────────
    public string OriginalCompanyName        { get; set; } = string.Empty;

    // ── 소프트 삭제 ──────────────────────────────────────────────────────────
    public bool   C_IsDeleted                { get; set; } = false;
    public DateTime? C_DeletedAt              { get; set; } = null;

    // ── 필드 ─────────────────────────────────────────────────────────────────
    public string    C_CompanyName              { get; set; } = string.Empty;
    public DateTime? C_ContractStart            { get; set; }
    public DateTime? C_ContractEnd              { get; set; }
    public int?      C_ContractDays             { get; set; }
    public decimal?  C_ContractAmountVATExcluded{ get; set; }
    public string    C_Abbreviation             { get; set; } = string.Empty;
    public string    C_ContractType             { get; set; } = string.Empty;  // 계약근거 (계약번호)
    public string    C_PlaceName                { get; set; } = string.Empty;  // 처리시설명 (측정인 cmb_emis_cmpy_plc_no 매칭용)
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

    // ── 잔여계약일 배지 ───────────────────────────────────────────────────────
    public int? DaysLeft =>
        C_ContractEnd.HasValue
            ? (int)(C_ContractEnd.Value.Date - DateTime.Today).TotalDays
            : (int?)null;

    public bool   HasDaysLeft   => DaysLeft.HasValue;
    public string DaysLeftText  =>
        DaysLeft switch
        {
            null    => "",
            >= 0    => $"D-{DaysLeft.Value}",
            _       => "만료",
        };
    public string DaysLeftBg =>
        DaysLeft switch
        {
            null    => "#333344",
            > 90    => "#1a3a1a",
            > 30    => "#3a3000",
            > 0     => "#3a1a00",
            _       => "#3a0000",
        };
    public string DaysLeftFg =>
        DaysLeft switch
        {
            null    => "#888888",
            > 90    => "#66cc66",
            > 30    => "#ccaa44",
            > 0     => "#cc8844",
            _       => "#cc4444",
        };
}