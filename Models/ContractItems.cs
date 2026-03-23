namespace ETA.Models;
using System;

public class Contract2
{
    public string C_CompanyName              { get; set; } = string.Empty;
    public DateTime? C_ContractStart         { get; set; }
    public DateTime? C_ContractEnd           { get; set; }
    public int? C_ContractDays               { get; set; }
    public decimal? C_ContractAmountVATExcluded { get; set; }
    public string C_Abbreviation             { get; set; } = string.Empty;
    public string C_ContractType             { get; set; } = string.Empty;
    public string C_Address                  { get; set; } = string.Empty;
    public string C_Representative           { get; set; } = string.Empty;
    public string C_FacilityType             { get; set; } = string.Empty;
    public string C_CategoryType             { get; set; } = string.Empty;
    public string C_MainProduct              { get; set; } = string.Empty;
    public string C_ContactPerson            { get; set; } = string.Empty;
    public string C_PhoneNumber              { get; set; } = string.Empty;
    public string C_Email                    { get; set; } = string.Empty;
}