using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// N-Hexan xlsx 파서 — 좌우 2-page 레이아웃.
/// A=구분, B=시료량, C=건조전무게, D=건조후무게, E=무게차, F=희석배수, G=농도, H=SN.
/// 행6=컬럼헤더, 행7=바탕시료, 데이터는 행8~.
/// </summary>
public static class NHexXlsxParser
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
        var docInfo = new ExcelDocInfo { IsNHEX = true };

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheets.First();

        var docDate = XlsxParserHelpers.ReadDocDate(ws);
        XlsxParserHelpers.ReadDocHeader(ws, docInfo);

        string itemAbbr = activeItems.FirstOrDefault() ?? "N-Hexan";

        // 바탕시료 (행7, 왼쪽 블록 기준)
        docInfo.바탕시료_시료량   = ws.Cell(7, 2).GetString().Trim();
        docInfo.바탕시료_건조전   = ws.Cell(7, 3).GetString().Trim();
        docInfo.바탕시료_건조후   = ws.Cell(7, 4).GetString().Trim();
        docInfo.바탕시료_무게차   = ws.Cell(7, 5).GetString().Trim();
        docInfo.바탕시료_희석배수 = ws.Cell(7, 6).GetString().Trim();
        var r7g = ws.Cell(7, 7).GetString().Trim();
        if (double.TryParse(r7g, out var r7v)) r7g = r7v.ToString("F1");
        docInfo.바탕시료_결과 = r7g;

        // 좌측 페이지 (col 1~8) + 우측 페이지 (col 9~16)
        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);
        XlsxParserHelpers.ParsePage(ws, rows, colName: 9, colResult: 15, colSN: 16, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);

        return new ParseResult(rows, docInfo, docDate, "NHEX");
    }
}
