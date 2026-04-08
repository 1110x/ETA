using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// SS(부유물질) xlsx 파서 — 단일 페이지, A=시료명, B=시료량, C=전무게, D=후무게,
/// E=무게차, F=희석배수, G=결과, H=SN. 행7=컬럼헤더, 데이터는 행8~.
/// </summary>
public static class SsXlsxParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo  DocInfo,
        string?       DocDate,
        string        DetectedFormat);

    public static ParseResult Parse(
        string path,
        string[] activeItems,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo { IsSS = true };

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheets.First();

        var docDate = XlsxParserHelpers.ReadDocDate(ws);
        XlsxParserHelpers.ReadDocHeader(ws, docInfo);

        // 형식 검증: 행7 A열에 "시료명" 텍스트 확인
        var row7A = ws.Cell(7, 1).GetString().Trim();
        if (!row7A.Contains("시료명"))
            throw new XlsxParseException($"선택한 카테고리(SS)와 엑셀 형식이 일치하지 않습니다.");

        string itemAbbr = activeItems.FirstOrDefault() ?? "SS";

        // SS: A=시료명, B=시료량, C=전무게, D=후무게, E=전후무게차, F=희석배수, G=결과, H=SN
        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);

        return new ParseResult(rows, docInfo, docDate, "SS");
    }
}
