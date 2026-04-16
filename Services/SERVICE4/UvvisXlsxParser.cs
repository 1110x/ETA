using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// UV-VIS 분광광도법 xlsx 파서 — T-N / T-P / Phenols 공용.
///
/// [총질소/총인 형식] 행5=STANDARD(농도+기울기+절편), 행6=abs(흡광도+R²), 행7=헤더, 행8~=데이터
/// [페놀류 형식] 행5=직접법 ST/abs(농도+기울기+절편), 행6=추출법 ST/abs, 행7=헤더, 행8~=데이터
///
/// 좌우 2-page 레이아웃. colResult=6(좌), 14(우).
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

        string itemAbbr = activeItems.FirstOrDefault() ?? "UVVIS";

        // 형식 판별: 행5 A열 텍스트로 총질소/총인 vs 페놀류 구분
        var row5A = ws.Cell(5, 1).GetString().Trim();
        bool isPhenols = row5A.Contains("직접법") || row5A.Contains("추출법");
        bool isStandard = row5A.Equals("STANDARD", StringComparison.OrdinalIgnoreCase);

        if (!isPhenols && !isStandard)
            throw new XlsxParseException("선택한 카테고리(UVVIS)와 엑셀 형식이 일치하지 않습니다.");

        if (isPhenols)
        {
            // 페놀류 형식: 행5=직접법(농도+기울기+절편), 행6=추출법(농도+기울기+절편)
            // 검량선은 추출법 기준 사용 (기본값). 직접법 비고에 따라 일부 시료가 사용
            var stdPts = new List<string>();
            for (int c = 2; c <= 6; c++)
            {
                var v = ws.Cell(5, c).GetString().Trim();
                if (double.TryParse(v, out var dv)) v = dv.ToString("G");
                stdPts.Add(v);
            }
            docInfo.Standard_Points = stdPts.ToArray();
            var slopeDirect = ws.Cell(5, 7).GetString().Trim();
            if (double.TryParse(slopeDirect, out var sd)) slopeDirect = sd.ToString("G6");
            // 추출법 기울기/절편/R² (행6)
            var slopeExtract = ws.Cell(6, 7).GetString().Trim();
            if (double.TryParse(slopeExtract, out var se)) slopeExtract = se.ToString("G6");
            docInfo.Standard_Slope = slopeExtract;
            var interceptExtract = ws.Cell(6, 8).GetString().Trim();
            if (double.TryParse(interceptExtract, out var ie)) interceptExtract = ie.ToString("G6");
            docInfo.Standard_Intercept = interceptExtract;
            // R²는 없지만 표시용으로 빈값
            docInfo.Abs_R2 = "";
            // 흡광도 값: 행6 B~F (추출법 abs)
            var absVals = new List<string>();
            for (int c = 2; c <= 6; c++)
            {
                var v = ws.Cell(6, c).GetString().Trim();
                if (double.TryParse(v, out var av)) v = av.ToString("G6");
                absVals.Add(v);
            }
            docInfo.Abs_Values = absVals.ToArray();
        }
        else
        {
            // 총질소/총인 형식: 행5=STANDARD(농도+기울기+절편), 행6=abs(흡광도+R²)
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
        }

        // 데이터 행: 페놀류는 R8~(헤더 R7), 총질소/총인도 동일
        // colResult=6(좌), 14(우)
        XlsxParserHelpers.ParsePage(ws, rows, colName: 1, colResult: 6, colSN: 8, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);
        XlsxParserHelpers.ParsePage(ws, rows, colName: 9, colResult: 14, colSN: 16, startRow: 8,
            itemAbbr: itemAbbr, resultFormatter: resultFormatter);

        string fmt = isPhenols ? "UVVIS_Phenols" : "UVVIS";
        return new ParseResult(rows, docInfo, docDate, fmt);
    }
}
