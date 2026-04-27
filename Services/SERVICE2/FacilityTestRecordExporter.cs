using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ETA.Services.Common;
using ETA.Services.SERVICE1;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 처리시설 시료를 기본항목 시험기록부 엑셀 템플릿에 채워 내보내는 유틸.
/// - 템플릿 위치: Data/Templates/기본항목 시험기록부/
/// - 출력 위치:  WritableDataRoot/Data/Exports/
/// - 좌측(A~H) 30행 먼저 채우고 넘치면 우측(I~P) 30행으로 확장 (최대 60건).
/// - 시료명 열(A/I)에 시료명, Remark/시료번호 열(H/P)에 시설명을 넣는다 (처리시설 SN이 없으므로).
/// - 측정값(시료량·흡광도·D1/D2·희석배수 등)은 사용자가 엑셀에서 직접 입력.
/// - 검량선/대조군 값(row 5~7)은 매일 다르므로 사용자가 덮어써서 사용.
/// </summary>
public static class FacilityTestRecordExporter
{
    // Show4 카테고리 키 → 템플릿 파일명
    private static readonly Dictionary<string, string> _templateFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["BOD"]     = "BOD 시험기록부.xlsx",
        ["TN"]      = "TN 시험기록부.xlsx",
        ["TP"]      = "TP 시험기록부.xlsx",
        ["PHENOLS"] = "Phenols 시험기록부.xlsx",
        ["SS"]      = "SS 시험기록부.xlsx",
    };

    // 카테고리 → 시험기록부 DB 테이블 조회용 Analyte 키 (한글 이름으로 변환)
    private static readonly Dictionary<string, string> _categoryToAnalyte = new(StringComparer.OrdinalIgnoreCase)
    {
        ["BOD"]     = "BOD",
        ["TN"]      = "T-N",
        ["TP"]      = "T-P",
        ["PHENOLS"] = "페놀류",
        ["SS"]      = "SS",
    };

    // 각 템플릿의 '사용자 입력/서명 컬럼' (공식 제외). 빈 시료행 지우기에 사용.
    //   left(1~8=A~H), right(9~16=I~P)
    private static readonly Dictionary<string, (int[] Left, int[] Right)> _inputColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        // BOD (row 5 headers): A=시료명 B=V C=D1 D=D2 E=f F=P G=Result(공식) H=Remark
        ["BOD"]     = (new[] { 1, 2, 3, 4, 5, 6, 8 }, new[] { 9, 10, 11, 12, 13, 14, 16 }),
        // TN/TP/Phenols (row 7 headers): A B C D | E,F 공식 | G=비고 H=시료번호
        ["TN"]      = (new[] { 1, 2, 3, 4, 7, 8 },    new[] { 9, 10, 11, 12, 15, 16 }),
        ["TP"]      = (new[] { 1, 2, 3, 4, 7, 8 },    new[] { 9, 10, 11, 12, 15, 16 }),
        ["PHENOLS"] = (new[] { 1, 2, 3, 4, 7, 8 },    new[] { 9, 10, 11, 12, 15, 16 }),
        // SS (row 7 headers): A B C D | E=공식 | F=희석배수 | G=결과(공식) H=Remark
        ["SS"]      = (new[] { 1, 2, 3, 4, 6, 8 },    new[] { 9, 10, 11, 12, 14, 16 }),
    };

    private const int FirstSampleRow = 8;
    private const int LastSampleRow  = 37;   // 좌/우 각 30행
    private const int LeftNameCol    = 1;    // A
    private const int LeftRemarkCol  = 8;    // H
    private const int RightNameCol   = 9;    // I
    private const int RightRemarkCol = 16;   // P
    private const int PerPage = LastSampleRow - FirstSampleRow + 1;  // 30

    public static bool IsSupported(string categoryKey)
        => !string.IsNullOrEmpty(categoryKey) && _templateFiles.ContainsKey(categoryKey);

    public static IEnumerable<string> SupportedCategories => _templateFiles.Keys;

    /// <summary>카테고리+날짜 → 시험기록부 엑셀 생성. 성공 시 생성된 파일 경로 반환.</summary>
    public static async Task<string?> ExportAsync(string categoryKey, string date)
    {
        if (!_templateFiles.TryGetValue(categoryKey, out var templateName)) return null;
        if (!DateTime.TryParse(date, out var d)) return null;

        var templatePath = Path.Combine(AppPaths.RootPath, "Data", "Templates", "기본항목 시험기록부", templateName);
        if (!File.Exists(templatePath))
        {
            Debug.WriteLine($"[FacilityTestRecordExporter] 템플릿 없음: {templatePath}");
            return null;
        }

        var samples = await Task.Run(() => LoadFacilitySamples(categoryKey, date));
        if (samples.Count == 0)
        {
            Debug.WriteLine($"[FacilityTestRecordExporter] 날짜 {date} · {categoryKey} — 처리시설 시료 없음");
            return null;
        }

        var exportsDir = Path.Combine(AppPaths.WritableDataRoot, "Data", "Exports");
        try { Directory.CreateDirectory(exportsDir); } catch { }
        var baseName = Path.GetFileNameWithoutExtension(templateName);
        var outPath  = Path.Combine(exportsDir, $"{baseName}_{d:yyyyMMdd}_{DateTime.Now:HHmmss}.xlsx");

        return await Task.Run(() =>
        {
            try
            {
                File.Copy(templatePath, outPath, overwrite: true);
                using var wb = new XLWorkbook(outPath);
                var ws = wb.Worksheets.First();

                // 헤더의 분석일자 (B1) — 우측(J1)은 =B1 공식이라 자동 갱신
                ws.Cell(1, 2).Value = d;
                ws.Cell(1, 2).Style.DateFormat.Format = "yyyy-mm-dd";

                // 시료행(row 8~37) 사용자입력 컬럼 비우기
                if (_inputColumns.TryGetValue(categoryKey, out var cols))
                {
                    for (int r = FirstSampleRow; r <= LastSampleRow; r++)
                    {
                        foreach (var c in cols.Left)  ws.Cell(r, c).Value  = Blank.Value;
                        foreach (var c in cols.Right) ws.Cell(r, c).Value = Blank.Value;
                    }
                }

                // 시료명 / 시설명 채우기 (좌측 30건 → 우측 30건 순)
                int writeCount = Math.Min(samples.Count, PerPage * 2);
                for (int i = 0; i < writeCount; i++)
                {
                    var s = samples[i];
                    if (i < PerPage)
                    {
                        int r = FirstSampleRow + i;
                        ws.Cell(r, LeftNameCol).Value   = s.SampleName;
                        ws.Cell(r, LeftRemarkCol).Value = s.Facility;
                    }
                    else
                    {
                        int r = FirstSampleRow + (i - PerPage);
                        ws.Cell(r, RightNameCol).Value   = s.SampleName;
                        ws.Cell(r, RightRemarkCol).Value = s.Facility;
                    }
                }

                wb.Save();
                Debug.WriteLine($"[FacilityTestRecordExporter] 출력 완료: {outPath} (시료 {writeCount}/{samples.Count})");
                return outPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FacilityTestRecordExporter] 오류: {ex.Message}");
                try { if (File.Exists(outPath)) File.Delete(outPath); } catch { }
                return null;
            }
        });
    }

    private readonly record struct SampleRow(string Facility, string SampleName, string SN);

    private static List<SampleRow> LoadFacilitySamples(string categoryKey, string date)
    {
        var list = new List<SampleRow>();

        var analyteKey = _categoryToAnalyte.TryGetValue(categoryKey, out var a) ? a : categoryKey;
        var tableName  = AnalysisService.GetRecordTableName(analyteKey);
        if (string.IsNullOrEmpty(tableName)) return list;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, tableName)) return list;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT m.시설명, t.시료명, t.SN
                FROM `{tableName}` t
                LEFT JOIN `처리시설_마스터` m ON m.시료명 = t.시료명
                WHERE LEFT(t.분석일, 10) = @d AND t.소스구분 = '처리시설'
                ORDER BY m.시설명, t.시료명, t.SN";
            cmd.Parameters.AddWithValue("@d", date);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var facility = r.IsDBNull(0) ? "" : r.GetValue(0)?.ToString() ?? "";
                var raw      = r.IsDBNull(1) ? "" : r.GetValue(1)?.ToString() ?? "";
                var sn       = r.IsDBNull(2) ? "" : r.GetValue(2)?.ToString() ?? "";

                // 시료명에 "원본\n↳ 매칭명" 형태로 저장된 경우 매칭명만 사용
                var sampleName = raw;
                if (raw.Contains('\n'))
                {
                    var parts = raw.Split('\n', 2);
                    sampleName = parts[1].TrimStart('↳').Trim();
                }
                list.Add(new SampleRow(facility, sampleName, sn));
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[FacilityTestRecordExporter] 시료 조회 오류: {ex.Message}");
        }
        return list;
    }
}
