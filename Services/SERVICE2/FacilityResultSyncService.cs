using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 회사 서버(rewater.wayble.eco)에서 처리시설 분석결과 동기화
/// Phase 2: 처리시설_결과 테이블에 데이터 INSERT
/// </summary>
public static class FacilityResultSyncService
{
    private const string BaseUrl = "https://rewater.wayble.eco/stp/api/subnote";
    private const string SessionCookie = "stpsession=OTczYzQ0NTMtM2ZhNC00YWVmLWI2NDYtM2QzODBiNWQ4YWM5; _ga_889WWPX1W1=GS2.1.s1776344282$o3$g1$t1776344289$j53$l0$h0; _ga=GA1.1.452346581.1775044184";

    // siteCd → 시설명 매핑
    private static readonly Dictionary<string, string> SiteCdMap = new()
    {
        { "PJT1020", "중흥" },
        { "PJT1021", "월내" },
        { "PJT1114", "4단계" },
        { "PJT1022", "율촌" },
        { "PJT1298", "세풍" }
    };

    // 처리시설별 officeCd
    private static readonly List<string> FacilityOffices = new()
    {
        "SITE053", "SITE065", "SITE066"
    };

    private static void Log(string msg)
    {
        Debug.WriteLine($"[FacilityResultSync] {msg}");
    }

    /// <summary>
    /// siteCd로 시설명 조회
    /// </summary>
    private static string GetFacilityName(string siteCd)
    {
        return SiteCdMap.TryGetValue(siteCd, out var name) ? name : "";
    }

    /// <summary>
    /// 모든 처리시설의 특정 날짜 데이터 동기화
    /// </summary>
    public static async Task SyncFacilityResultsAsync(string inputDate)
    {
        Log($"처리시설 동기화 시작: {inputDate}");

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            foreach (var officeCd in FacilityOffices)
            {
                Log($"동기화 중: {officeCd}");
                await SyncOfficeCdAsync(conn, officeCd, inputDate);
            }

            Log("처리시설 동기화 완료");
        }
        catch (Exception ex)
        {
            Log($"동기화 오류: {ex.Message}\n{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 특정 처리시설의 데이터 동기화
    /// </summary>
    private static async Task SyncOfficeCdAsync(System.Data.DbConnection conn, string officeCd, string inputDate)
    {
        try
        {
            // 1. OCR 데이터 조회
            var ocrList = await GetOcrListAsync(officeCd, inputDate);
            if (ocrList == null || ocrList.Count == 0)
            {
                Log($"{officeCd}: OCR 데이터 없음");
                return;
            }

            Log($"{officeCd}: {ocrList.Count}건 데이터");

            // 2. 각 시료별 처리시설_결과 INSERT
            foreach (var ocr in ocrList)
            {
                // siteCd로 시설명 결정
                var facilityName = GetFacilityName(ocr.siteCd);
                if (!string.IsNullOrEmpty(facilityName))
                {
                    InsertFacilityResult(conn, facilityName, ocr);
                }
            }
        }
        catch (Exception ex)
        {
            Log($"{officeCd} 동기화 실패: {ex.Message}");
        }
    }

    /// <summary>
    /// API에서 OCR 목록 조회
    /// </summary>
    private static async Task<List<SubNoteOcr>> GetOcrListAsync(string officeCd, string inputDate)
    {
        try
        {
            using var handler = new HttpClientHandler();
            using var client = new HttpClient(handler);

            var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/selectSubNoteOcrList");
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15");
            request.Headers.Add("Accept", "application/json, text/javascript, */*; q=0.01");
            request.Headers.Add("X-Requested-With", "XMLHttpRequest");
            request.Headers.Add("Cookie", SessionCookie);

            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("officeCd", officeCd),
                new KeyValuePair<string, string>("inputDate", inputDate)
            });
            request.Content = content;

            var response = await client.SendAsync(request);
            var json = await response.Content.ReadAsStringAsync();

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("body", out var body) &&
                body.TryGetProperty("subNoteList", out var list))
            {
                var result = new List<SubNoteOcr>();
                foreach (var item in list.EnumerateArray())
                {
                    result.Add(JsonSerializer.Deserialize<SubNoteOcr>(item.GetRawText())!);
                }
                return result;
            }

            return new List<SubNoteOcr>();
        }
        catch (Exception ex)
        {
            Log($"OCR 조회 오류 ({officeCd}): {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 처리시설_결과에 데이터 INSERT
    /// </summary>
    private static void InsertFacilityResult(System.Data.DbConnection conn, string facilityName, SubNoteOcr ocr)
    {
        try
        {
            // 시료명: siteNm (없으면 sampleCategory)
            var sampleName = string.IsNullOrEmpty(ocr.siteNm) || ocr.siteNm == "-"
                ? ocr.sampleCategoryNm
                : ocr.siteNm;

            // 채취일자: inputDate 형식 변환 (20260415 → 2026-04-15)
            var collectionDate = ConvertExcelDate(ocr.inputDate);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                INSERT INTO `처리시설_결과`
                    (시설명, 시료명, 채취일자, BOD, TOC, SS, `T-N`, `T-P`,
                     총대장균군, 입력자, 입력일시)
                VALUES
                    (@시설명, @시료명, @채취일자, @BOD, @TOC, @SS, @TN, @TP,
                     @총대장균군, @입력자, @입력일시)";

            cmd.Parameters.AddWithValue("@시설명", facilityName);
            cmd.Parameters.AddWithValue("@시료명", sampleName);
            cmd.Parameters.AddWithValue("@채취일자", collectionDate);

            // 분석값: null이면 빈값
            cmd.Parameters.AddWithValue("@BOD", FormatAnalysisValue(ocr.bodD1, ocr.bodD2, ocr.bodP));
            cmd.Parameters.AddWithValue("@TOC", FormatAnalysisValue(ocr.tocMl, ocr.tocP));
            cmd.Parameters.AddWithValue("@SS", FormatAnalysisValue(ocr.ssMl, ocr.ssP));
            cmd.Parameters.AddWithValue("@TN", FormatAnalysisValue(ocr.tnMl, ocr.tnP));
            cmd.Parameters.AddWithValue("@TP", FormatAnalysisValue(ocr.tpMl, ocr.tpP));
            cmd.Parameters.AddWithValue("@총대장균군", FormatAnalysisValue(ocr.coliA, ocr.coliB));
            cmd.Parameters.AddWithValue("@입력자", ocr.createUser ?? "");
            cmd.Parameters.AddWithValue("@입력일시", ocr.createTime ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            Log($"INSERT 오류: {ex.Message}");
        }
    }

    private static string ConvertExcelDate(string dateStr)
    {
        if (string.IsNullOrEmpty(dateStr) || dateStr.Length < 8) return "";

        // 20260415 → 2026-04-15
        var year = dateStr.Substring(0, 4);
        var month = dateStr.Substring(4, 2);
        var day = dateStr.Substring(6, 2);

        return $"{year}-{month}-{day}";
    }

    private static string FormatAnalysisValue(params string[] values)
    {
        var nonNull = values
            .Where(v => !string.IsNullOrEmpty(v) && v != "-")
            .ToList();

        return nonNull.Count > 0 ? string.Join(", ", nonNull) : "";
    }
}

/// <summary>
/// API 응답: SubNote OCR 데이터
/// </summary>
public class SubNoteOcr
{
    public string officeCd { get; set; }
    public string siteNm { get; set; }
    public string siteCd { get; set; }
    public string inputDate { get; set; }
    public string sampleCategory { get; set; }
    public string sampleCategoryNm { get; set; }
    public string bodNo { get; set; }
    public string bodMl { get; set; }
    public string bodD1 { get; set; }
    public string bodD2 { get; set; }
    public string bodP { get; set; }
    public string tocMl { get; set; }
    public string tocP { get; set; }
    public string ssMl { get; set; }
    public string ssP { get; set; }
    public string tnMl { get; set; }
    public string tnP { get; set; }
    public string tpMl { get; set; }
    public string tpP { get; set; }
    public string coliA { get; set; }
    public string coliB { get; set; }
    public string createTime { get; set; }
    public string createUser { get; set; }
}
