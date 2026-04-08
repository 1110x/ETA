using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// BOD 전용 xlsx 파서 — 행6 식종수, 행7 SCF, 행8~ 데이터 (좌우 2-page).
/// SS/NHEX/UVVIS/TOC는 각각 전용 파서(SsXlsxParser, NHexXlsxParser, UvvisXlsxParser, TocXlsxParser).
/// </summary>
public static class BodExcelParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo  DocInfo,
        string?       DocDate,
        string        DetectedFormat);

    /// <param name="path">.xlsx 파일 경로</param>
    /// <param name="activeItems">카테고리의 분석항목 약칭 리스트 (첫 항목을 itemAbbr로 사용)</param>
    /// <param name="resultFormatter">결과값 포맷터 (value, itemAbbr) → 포맷된 값. null이면 원본 사용</param>
    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo();

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheets.First();

        var docDate = XlsxParserHelpers.ReadDocDate(ws);
        XlsxParserHelpers.ReadDocHeader(ws, docInfo);

        // 형식 교차 검증: UVVIS/SS 양식이면 BOD 파서로 부르면 안 됨
        var row5A = ws.Cell(5, 1).GetString().Trim();
        var row7A = ws.Cell(7, 1).GetString().Trim();
        if (row5A.Equals("STANDARD", StringComparison.OrdinalIgnoreCase))
            throw new XlsxParseException("선택한 카테고리(BOD)와 엑셀 형식(UVVIS)이 일치하지 않습니다.");
        if (row7A.Contains("시료명"))
            throw new XlsxParseException("선택한 카테고리(BOD)와 엑셀 형식(SS)이 일치하지 않습니다.");

        string itemAbbr = activeItems.FirstOrDefault() ?? "BOD";

        // 행6 식종수, 행7 SCF
        XlsxParserHelpers.ReadSeedScfRows(ws, docInfo);

        // 좌측 페이지 (col 1~8) + 우측 페이지 (col 9~16)
        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);
        XlsxParserHelpers.ParsePage(ws, rows, colName: 9, colResult: 15, colSN: 16, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);

        return new ParseResult(rows, docInfo, docDate, "BOD");
    }
}
