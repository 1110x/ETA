using System;
using System.Collections.Generic;

namespace ETA.Models;

// ── 의뢰 정보 (시료 1건) ─────────────────────────────────────────────────────
public class SampleRequest
{
    public int    Id              { get; set; }

    // ── 고정 컬럼 ────────────────────────────────────────────────────────────
    public string 채취일자     { get; set; } = "";
    public string 채취시간     { get; set; } = "";
    public string 의뢰사업장   { get; set; } = "";
    public string 약칭         { get; set; } = "";   // 트리 그룹 기준
    public string 시료명       { get; set; } = "";
    public string 견적번호     { get; set; } = "";
    public string 입회자       { get; set; } = "";
    public string 시료채취자1  { get; set; } = "";   // DB: "시료채취자-1"
    public string 시료채취자2  { get; set; } = "";   // DB: "시료채취자-2"
    public string 방류허용기준 { get; set; } = "";   // DB: "방류허용기준 적용유무"
    public string 정도보증     { get; set; } = "";   // DB: "정도보증유무"
    public string 분석종료일 { get; set; } = "";
    public string 견적구분     { get; set; } = "";

    // ── 분석결과: 항목명 → 결과값 (NULL 컬럼은 포함 안 됨) ─────────────────
    public Dictionary<string, string> 분석결과 { get; set; } = new();

    // ── 트리 표시용 ──────────────────────────────────────────────────────────
    public string TreeLabel => $"【{채취일자}】 {시료명}";
}

// ── 첨부파일 ─────────────────────────────────────────────────────────────────
public class SampleAttachment
{
    public int    Id        { get; set; }
    public int    SampleId  { get; set; }
    public string 원본파일명 { get; set; } = "";
    public string 저장경로  { get; set; } = "";
    public string 등록일시  { get; set; } = "";
}

// ── 분석 결과 행 (리스트 1행) ────────────────────────────────────────────────
public class AnalysisResultRow
{
    public string 항목명         { get; set; } = "";
    public string 결과값         { get; set; } = "";
    public string 단위           { get; set; } = "";
    public string 분석방법       { get; set; } = "";
    public string 분석장비       { get; set; } = "";
    public string ES             { get; set; } = "";
    public string Category       { get; set; } = "";
    public string Original결과값 { get; set; } = "";
    public string DB컬럼명       { get; set; } = "";
}