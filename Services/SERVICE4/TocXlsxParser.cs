using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// TOC(NPOC/TCIC) xlsx 파서 — BOD 양식과 유사하지만 검정곡선(기울기·절편) 셀이 추가됨.
/// NPOC: 행5 = Standard, G=기울기 H=절편
/// TCIC: 행5 = Standard(TC), 행7 = Standard(IC), 각각 G=기울기 H=절편
/// </summary>
public static class TocXlsxParser
{
    public sealed record ParseResult(
        List<ExcelRow> Rows,
        ExcelDocInfo  DocInfo,
        string?       DocDate,
        string        Method); // "NPOC" / "TCIC"

    public static ParseResult Parse(
        string path,
        Func<string, string, string>? resultFormatter = null)
    {
        var rows    = new List<ExcelRow>();
        var docInfo = new ExcelDocInfo();

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheets.First();

        var docDate = XlsxParserHelpers.ReadDocDate(ws);
        XlsxParserHelpers.ReadDocHeader(ws, docInfo);
        XlsxParserHelpers.ReadSeedScfRows(ws, docInfo);

        // 행5 A열 내용으로 NPOC vs TCIC 판별
        var row5A = ws.Cell(5, 1).GetString().Trim();
        bool isNpoc = row5A.Equals("STANDARD", StringComparison.OrdinalIgnoreCase);

        string method;
        if (isNpoc)
        {
            // NPOC: 행5 G=기울기(a), H=절편(b)
            docInfo.IsTocNPOC       = true;
            docInfo.TocSlope_TC     = ws.Cell(5, 7).GetString().Trim();
            docInfo.TocIntercept_TC = ws.Cell(5, 8).GetString().Trim();
            method = "NPOC";
        }
        else
        {
            // TC-IC: 행5=Standard(TC) G=기울기, H=절편 / 행7=Standard(IC) G=기울기, H=절편
            docInfo.IsTocTCIC       = true;
            docInfo.TocSlope_TC     = ws.Cell(5, 7).GetString().Trim();
            docInfo.TocIntercept_TC = ws.Cell(5, 8).GetString().Trim();
            docInfo.TocSlope_IC     = ws.Cell(7, 7).GetString().Trim();
            docInfo.TocIntercept_IC = ws.Cell(7, 8).GetString().Trim();
            method = "TCIC";
        }

        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8,
            itemAbbr: "TOC", resultFormatter: resultFormatter);
        XlsxParserHelpers.ParsePage(ws, rows, colName: 9, colResult: 15, colSN: 16, startRow: 8,
            itemAbbr: "TOC", resultFormatter: resultFormatter);

        return new ParseResult(rows, docInfo, docDate, method);
    }
}
