namespace ETA.Models;

/// <summary>시약 — 실험실 시약 재고 및 위험도 관리</summary>
public class Reagent
{
    public int    Id             { get; set; }
    public string ITEM_NO        { get; set; } = "";
    public string 품목명         { get; set; } = "";   // 국문명
    public string 영문명         { get; set; } = "";
    public string CAS번호        { get; set; } = "";
    public string 화학식         { get; set; } = "";
    public string 규격           { get; set; } = "";   // 예) 500mL, 1kg
    public string 단위           { get; set; } = "";   // 예) mL, g, L
    public string 제조사         { get; set; } = "";
    public string 위험등급       { get; set; } = "일반"; // 일반/주의/위험
    public string GHS            { get; set; } = "";   // 예) GHS02,GHS06
    public string 보관조건       { get; set; } = "";
    public int    재고량         { get; set; } = 0;
    public int    당월사용량      { get; set; } = 0;
    public int    전월사용량      { get; set; } = 0;
    public int    적정사용량      { get; set; } = 0;
    public int    최대적정보유량   { get; set; } = 0;
    public string 만료일         { get; set; } = "";   // yyyy-MM-dd
    public string 비고           { get; set; } = "";
    public string 등록일         { get; set; } = "";   // yyyy-MM-dd
    public string 상태           { get; set; } = "정상"; // 정상/주의/폐기

    // ── 화학물질관리법 분류 (다중 선택) ──────────────────────────────────────
    public bool   유독물질       { get; set; } = false;
    public bool   허가물질       { get; set; } = false;
    public bool   제한물질       { get; set; } = false;
    public bool   금지물질       { get; set; } = false;
    public bool   사고대비물질   { get; set; } = false;
}
