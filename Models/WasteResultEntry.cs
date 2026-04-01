using System.Collections.Generic;

namespace ETA.Models;

/// <summary>폐수_결과 테이블 헤더 레코드 (업체 x 채취일자)</summary>
public class WasteResultEntry
{
    public int    Id       { get; set; }
    public int?   의뢰Id   { get; set; }   // 폐수_의뢰.id (nullable)
    public string 관리번호 { get; set; } = "";
    public string 업체명   { get; set; } = "";
    public string 채취일자 { get; set; } = "";
    public string 입력자   { get; set; } = "";
    public string 입력일시 { get; set; } = "";

    /// <summary>항목별 결과 목록 (조회 시 채워짐)</summary>
    public List<WasteResultItem> 항목들 { get; set; } = new();
}

/// <summary>폐수_결과_항목 테이블 — 항목(BOD/TOC 등) x 결과값</summary>
public class WasteResultItem
{
    public int    Id     { get; set; }
    public int    결과Id { get; set; }   // FK → WasteResultEntry.Id
    public string 항목   { get; set; } = "";   // 'BOD', 'TOC', 'SS', ...
    public string 결과값 { get; set; } = "";
    public string 단위   { get; set; } = "";
    public string 비고   { get; set; } = "";
}

/// <summary>결과_제출이력 테이블</summary>
public class ResultSubmitLog
{
    public int    Id       { get; set; }
    public string 제출유형 { get; set; } = "";   // '측정인' | 'ERP' | 'Zero4'
    public string 대상유형 { get; set; } = "";   // '처리시설' | '폐수업소'
    public string 대상명   { get; set; } = "";
    public string 기간시작 { get; set; } = "";
    public string 기간종료 { get; set; } = "";
    public int    제출건수 { get; set; }
    public string 제출자   { get; set; } = "";
    public string 제출일시 { get; set; } = "";
    public string 파일경로 { get; set; } = "";
    public string 상태     { get; set; } = "대기";   // '대기' | '완료' | '오류'
    public string 비고     { get; set; } = "";
}
