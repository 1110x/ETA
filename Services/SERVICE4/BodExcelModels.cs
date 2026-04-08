using ETA.Models;
using ETA.Views.Pages.PAGE1;
using System;
using System.Collections.Generic;

namespace ETA.Services.SERVICE4;

/// <summary>엑셀 한 행의 파싱 결과 + 이후 매칭/소스 상태</summary>
public class ExcelRow
{
    public string 시료명 { get; set; } = "";
    public string 원본시료명 { get; set; } = ""; // 수동매칭 전 원래 이름
    public string SN { get; set; } = "";
    public string Result { get; set; } = "";
    public string 시료량 { get; set; } = "";
    public string D1 { get; set; } = "";
    public string D2 { get; set; } = "";
    public string Fxy { get; set; } = "";  // f(x/y) 식종액 함유율
    public string P { get; set; } = "";    // 희석배수
    // TOC TCIC 전용 (TOC_TCIC_DATA 컬럼 매핑)
    public string TCAU  { get; set; } = "";
    public string TCcon { get; set; } = "";
    public string ICAU  { get; set; } = "";
    public string ICcon { get; set; } = "";
    public WasteSample? Matched { get; set; }                   // 폐수배출업소 매칭
    public AnalysisRequestRecord? MatchedAnalysis { get; set; } // 수질분석센터 매칭
    public FacilityResultRow? MatchedFacility { get; set; }     // 처리시설 매칭
    public string? MatchedFacilityName { get; set; }            // 처리시설명
    public MatchStatus Status { get; set; }
    public SourceType Source { get; set; }
    public bool Enabled { get; set; } = true;
    /// <summary>정도관리(BK/CCV/FBK/MBK/DW 등) 시료 여부 — 기록부에는 남기되 의뢰및결과엔 반영 안 함</summary>
    public bool IsControl { get; set; }
}

public enum MatchStatus { 입력가능, 덮어쓰기, 미매칭, 대기 }
public enum SourceType { 미분류, 폐수배출업소, 수질분석센터, 처리시설 }

/// <summary>엑셀 문서 헤더 정보 (행1~7) + 검정곡선 정보</summary>
public class ExcelDocInfo
{
    public string 문서번호 { get; set; } = "";
    public string 분석방법 { get; set; } = "";
    public string 결과표시 { get; set; } = "";
    public string 관련근거 { get; set; } = "";
    // 식종수의 BOD (행6)
    public string 식종수_시료량 { get; set; } = "";
    public string 식종수_D1 { get; set; } = "";
    public string 식종수_D2 { get; set; } = "";
    public string 식종수_P { get; set; } = "";
    public string 식종수_Result { get; set; } = "";
    public string 식종수_Remark { get; set; } = "";  // 식종수(%) 1.5
    // SCF (행7)
    public string SCF_시료량 { get; set; } = "";
    public string SCF_D1 { get; set; } = "";
    public string SCF_D2 { get; set; } = "";
    public string SCF_Result { get; set; } = "";
    // N-Hexan 바탕시료 (행7)
    public string 바탕시료_시료량 { get; set; } = "";
    public string 바탕시료_건조전 { get; set; } = "";
    public string 바탕시료_건조후 { get; set; } = "";
    public string 바탕시료_무게차 { get; set; } = "";
    public string 바탕시료_희석배수 { get; set; } = "";
    public string 바탕시료_결과 { get; set; } = "";
    // 검량곡선 (UV VIS 등)
    public bool IsUVVIS { get; set; }
    public bool IsSS { get; set; }
    public bool IsNHEX { get; set; }
    public string[] Standard_Points { get; set; } = Array.Empty<string>(); // 표준용액 농도
    public string Standard_Slope { get; set; } = "";   // 기울기 a
    public string Standard_Intercept { get; set; } = ""; // 절편 b
    public string[] Abs_Values { get; set; } = Array.Empty<string>(); // 흡광도 측정값
    // TOC 전용
    public bool IsTocNPOC { get; set; }
    public bool IsTocTCIC { get; set; }
    public string TocSlope_TC { get; set; } = "";
    public string TocSlope_IC { get; set; } = "";
    public string TocIntercept_TC { get; set; } = "";
    public string TocIntercept_IC { get; set; } = "";
    public string TocR2_TC { get; set; } = "";
    public string TocR2_IC { get; set; } = "";
    public string[] TocStdConcs    { get; set; } = Array.Empty<string>(); // TC ST-1~5 공칭농도
    public string[] TocStdAreas    { get; set; } = Array.Empty<string>(); // TC ST-1~5 면적(AU)
    public string[] TocStdConcs_IC { get; set; } = Array.Empty<string>(); // IC ST-1~5 공칭농도 (TCIC 전용)
    public string[] TocStdAreas_IC { get; set; } = Array.Empty<string>(); // IC ST-1~5 면적(AU) (TCIC 전용)
    public string Abs_R2 { get; set; } = "";   // R²
    // GC 전용 (Agilent ChemStation/MassHunter)
    public bool IsGcMode { get; set; }
    public string GcFormat { get; set; } = "";           // VocMulti/VocSingle/SingleNoIstd/SingleExpConc
    public List<GcCompoundCalInfo> GcCompoundCals { get; set; } = new();
}

public class GcCompoundCalInfo
{
    public string   Name      { get; set; } = "";
    public bool     HasIstd   { get; set; }
    public string   Slope     { get; set; } = "";
    public string   Intercept { get; set; } = "";
    public string   R         { get; set; } = "";
    public string[] StdConcs  { get; set; } = Array.Empty<string>(); // ST-1~N 공칭농도
    public string[] StdResps  { get; set; } = Array.Empty<string>(); // ST-1~N 응답
}
