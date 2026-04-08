using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ETA.Services.SERVICE4;

/// <summary>BodExcelParser / TocXlsxParser 공통 xlsx 파싱 헬퍼</summary>
internal static class XlsxParserHelpers
{
    /// <summary>Row1 B열에서 문서 날짜 추출 (텍스트/Excel 서식 모두 처리)</summary>
    public static string? ReadDocDate(IXLWorksheet ws)
    {
        var b1Cell = ws.Cell(1, 2);
        if (b1Cell.TryGetValue<DateTime>(out var parsedDocDate))
            return parsedDocDate.ToString("yyyy-MM-dd");
        var dateCellVal = b1Cell.GetString().Trim();
        if (DateTime.TryParse(dateCellVal, out parsedDocDate))
            return parsedDocDate.ToString("yyyy-MM-dd");
        return null;
    }

    /// <summary>행1~4 문서 헤더 정보 (문서번호/분석방법/결과표시/관련근거)</summary>
    public static void ReadDocHeader(IXLWorksheet ws, ExcelDocInfo docInfo)
    {
        docInfo.문서번호 = ws.Cell(1, 2).GetString().Trim();
        docInfo.분석방법 = ws.Cell(2, 2).GetString().Trim();
        docInfo.결과표시 = ws.Cell(3, 2).GetString().Trim();
        docInfo.관련근거 = ws.Cell(4, 2).GetString().Trim();
    }

    /// <summary>행6 식종수 / 행7 SCF (BOD / TOC 공통 레이아웃)</summary>
    public static void ReadSeedScfRows(IXLWorksheet ws, ExcelDocInfo docInfo)
    {
        docInfo.식종수_시료량 = ws.Cell(6, 2).GetString().Trim();
        docInfo.식종수_D1     = ws.Cell(6, 3).GetString().Trim();
        docInfo.식종수_D2     = ws.Cell(6, 4).GetString().Trim();
        docInfo.식종수_P      = ws.Cell(6, 6).GetString().Trim();
        var r6r = ws.Cell(6, 7).GetString().Trim();
        if (double.TryParse(r6r, out var r6v)) r6r = r6v.ToString("F1");
        docInfo.식종수_Result = r6r;
        docInfo.식종수_Remark = ws.Cell(6, 8).GetString().Trim();
        docInfo.SCF_시료량 = ws.Cell(7, 2).GetString().Trim();
        docInfo.SCF_D1     = ws.Cell(7, 3).GetString().Trim();
        docInfo.SCF_D2     = ws.Cell(7, 4).GetString().Trim();
        var r7r = ws.Cell(7, 7).GetString().Trim();
        if (double.TryParse(r7r, out var r7v)) r7r = r7v.ToString("F4");
        docInfo.SCF_Result = r7r;
    }

    /// <summary>
    /// 데이터 행(좌/우 페이지 단위) 파싱 — 시료명/시료량/D1/D2/f(x/y)/P/결과값/SN.
    /// BOD / TOC / SS / NHEX / UVVIS 공통.
    /// </summary>
    public static void ParsePage(
        IXLWorksheet ws,
        List<ExcelRow> rows,
        int colName, int colResult, int colSN, int startRow,
        string itemAbbr,
        Func<string, string, string>? resultFormatter)
    {
        // 기초정보 컬럼: 시료량=colName+1, D1=colName+2, D2=colName+3, f(x/y)=colName+4, P=colName+5
        int colVol = colName + 1, colD1 = colName + 2, colD2 = colName + 3;
        int colFxy = colName + 4, colP  = colName + 5;

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        for (int r = startRow; r <= lastRow; r++)
        {
            var nameCell   = ws.Cell(r, colName).GetString().Trim();
            var resultCell = ws.Cell(r, colResult).GetString().Trim();
            var snCell     = ws.Cell(r, colSN).GetString().Trim();

            if (string.IsNullOrEmpty(nameCell) && string.IsNullOrEmpty(snCell)) break;
            if (string.IsNullOrEmpty(nameCell)) continue;
            if (nameCell.Contains("식종") || nameCell.Contains("SCF") || nameCell.Contains("분석담당")) continue;

            if (resultFormatter != null && double.TryParse(resultCell, out var val))
                resultCell = resultFormatter(val.ToString(), itemAbbr);

            var exRow = new ExcelRow { 시료명 = nameCell, SN = snCell, Result = resultCell };
            exRow.시료량 = ws.Cell(r, colVol).GetString().Trim();
            exRow.D1   = ws.Cell(r, colD1).GetString().Trim();
            exRow.D2   = ws.Cell(r, colD2).GetString().Trim();
            exRow.Fxy  = ws.Cell(r, colFxy).GetString().Trim();
            exRow.P    = colP < colResult ? ws.Cell(r, colP).GetString().Trim() : "";
            rows.Add(exRow);
        }
    }
}

/// <summary>xlsx 파서 공통 예외</summary>
public sealed class XlsxParseException : Exception
{
    public XlsxParseException(string message) : base(message) { }
}
