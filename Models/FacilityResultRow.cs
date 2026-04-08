using System.Collections.Generic;

namespace ETA.Models;

public class FacilityResultRow
{
    public int    Id        { get; set; }   // 0 = 미저장
    public int    마스터Id   { get; set; }
    public string 시료명     { get; set; } = "";
    public string 비고마스터  { get; set; } = "";  // 처리시설_마스터.비고 (주1회 목요일 등)
    public string 비고       { get; set; } = "";  // 측정결과 비고

    /// <summary>동적 항목값 (컬럼명 → 값). BOD/TOC/SS/T-N 등 컬럼명 그대로 사용</summary>
    public Dictionary<string, string> Values { get; set; } = new();

    /// <summary>동적 항목 활성 여부 (컬럼명 → bool). 처리시설_마스터의 'O' 값 기준</summary>
    public Dictionary<string, bool> Active { get; set; } = new();

    /// <summary>편의 인덱서: row["BOD"] 등으로 직접 접근</summary>
    public string this[string key]
    {
        get => Values.TryGetValue(key, out var v) ? v : "";
        set => Values[key] = value ?? "";
    }

    public bool IsActive(string key) => Active.TryGetValue(key, out var v) && v;

    // ── 하위 호환 프로퍼티 (기존 코드 점진 마이그레이션용) ─────────────────
    public string BOD       { get => this["BOD"];       set => this["BOD"]       = value; }
    public string TOC       { get => this["TOC"];       set => this["TOC"]       = value; }
    public string SS        { get => this["SS"];        set => this["SS"]        = value; }
    public string TN        { get => this["T-N"];       set => this["T-N"]       = value; }
    public string TP        { get => this["T-P"];       set => this["T-P"]       = value; }
    public string 총대장균군 { get => this["총대장균군"]; set => this["총대장균군"] = value; }
    public string COD       { get => this["COD"];       set => this["COD"]       = value; }
    public string 염소이온   { get => this["염소이온"];   set => this["염소이온"]   = value; }
    public string 영양염류   { get => this["영양염류"];   set => this["영양염류"]   = value; }
    public string 함수율    { get => this["함수율"];    set => this["함수율"]    = value; }
    public string 중금속    { get => this["중금속"];    set => this["중금속"]    = value; }

    public bool BOD활성       { get => IsActive("BOD");       set => Active["BOD"]       = value; }
    public bool TOC활성       { get => IsActive("TOC");       set => Active["TOC"]       = value; }
    public bool SS활성        { get => IsActive("SS");        set => Active["SS"]        = value; }
    public bool TN활성        { get => IsActive("T-N");       set => Active["T-N"]       = value; }
    public bool TP활성        { get => IsActive("T-P");       set => Active["T-P"]       = value; }
    public bool 총대장균군활성 { get => IsActive("총대장균군"); set => Active["총대장균군"] = value; }
    public bool COD활성       { get => IsActive("COD");       set => Active["COD"]       = value; }
    public bool 염소이온활성   { get => IsActive("염소이온");   set => Active["염소이온"]   = value; }
    public bool 영양염류활성   { get => IsActive("영양염류");   set => Active["영양염류"]   = value; }
    public bool 함수율활성    { get => IsActive("함수율");    set => Active["함수율"]    = value; }
    public bool 중금속활성    { get => IsActive("중금속");    set => Active["중금속"]    = value; }
}
