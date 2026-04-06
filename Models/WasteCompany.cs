using System;
using System.Collections.Generic;

namespace ETA.Models;

public class WasteCompany
{
    public string 프로젝트 { get; set; } = string.Empty;
    public string 프로젝트명 { get; set; } = string.Empty;
    public string 관리번호 { get; set; } = string.Empty;
    public string 업체명 { get; set; } = string.Empty;
    public string Original업체명 { get; set; } = string.Empty;
    public string 사업자번호 { get; set; } = string.Empty;
    public string 약칭               { get; set; } = string.Empty;
    public string 비용부담금_업체명  { get; set; } = string.Empty;
    // 허용기준
    public string BOD      { get; set; } = string.Empty;
    public string TOC      { get; set; } = string.Empty;
    public string SS       { get; set; } = string.Empty;
    public string TN       { get; set; } = string.Empty;
    public string TP       { get; set; } = string.Empty;
    public string Phenols  { get; set; } = string.Empty;
    public string NHexan   { get; set; } = string.Empty;
    public string 승인유량         { get; set; } = string.Empty;
    public string 기타특이사항     { get; set; } = string.Empty;
}