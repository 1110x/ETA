using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// UV-VIS 분광광도법 xlsx 파서 — T-N / T-P / Phenols 공용.
/// 행5 = STANDARD 표준용액 농도(B~F) + 기울기 a(G) + 절편 b(H).
/// 행6 = 흡광도(B~F) + R²(G).
/// 좌우 2-page 레이아웃. BOD와 달리 colResult=6 (농도컬럼).
/// </summary>
public static class UvvisXlsxParser
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
        var docInfo = new ExcelDocInfo { IsUVVIS = true };

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wb = new XLWorkbook(fs);
        var ws = wb.Worksheets.First();

        var docDate = XlsxParserHelpers.ReadDocDate(ws);
        XlsxParserHelpers.ReadDocHeader(ws, docInfo);

        // 형식 검증: 행5 A열에 "STANDARD" 확인
        var row5A = ws.Cell(5, 1).GetString().Trim();
        if (!row5A.Equals("STANDARD", StringComparison.OrdinalIgnoreCase))
            throw new XlsxParseException("선택한 카테고리(UVVIS)와 엑셀 형식이 일치하지 않습니다.");

        string itemAbbr = activeItems.FirstOrDefault() ?? "UVVIS";

        // 행5: STANDARD 표준용액 농도 + 기울기(a) + 절편(b)
        var stdPts = new List<string>();
        for (int c = 2; c <= 6; c++)
        {
            var v = ws.Cell(5, c).GetString().Trim();
            if (double.TryParse(v, out var dv)) v = dv.ToString("G");
            stdPts.Add(v);
        }
        docInfo.Standard_Points = stdPts.ToArray();
        var slope = ws.Cell(5, 7).GetString().Trim();
        if (double.TryParse(slope, out var sv)) slope = sv.ToString("G6");
        docInfo.Standard_Slope = slope;
        var intercept = ws.Cell(5, 8).GetString().Trim();
        if (double.TryParse(intercept, out var iv)) intercept = iv.ToString("G6");
        docInfo.Standard_Intercept = intercept;

        // 행6: abs 흡광도 + R²
        var absVals = new List<string>();
        for (int c = 2; c <= 6; c++)
        {
            var v = ws.Cell(6, c).GetString().Trim();
            if (double.TryParse(v, out var av)) v = av.ToString("G6");
            absVals.Add(v);
        }
        docInfo.Abs_Values = absVals.ToArray();
        var r2 = ws.Cell(6, 7).GetString().Trim();
        if (double.TryParse(r2, out var r2v)) r2 = r2v.ToString("F5");
        docInfo.Abs_R2 = r2;

        // UV VIS: colResult=6 (BOD는 7), 페이지2는 colResult=14
        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 6, colSN: 8, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);
        XlsxParserHelpers.ParsePage(ws, rows, colName: 9, colResult: 14, colSN: 16, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);

        return new ParseResult(rows, docInfo, docDate, "UVVIS");
    }
}
