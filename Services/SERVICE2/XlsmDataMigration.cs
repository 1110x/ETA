using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ETA.Services.Common;

namespace ETA.Services.SERVICE2;

/// <summary>
/// Phase 2 xlsm 데이터 마이그레이션
/// Docs/CHUNGHA-김가린.xlsm의 DATA 시트들 → 비용부담금_결과
///
/// xlsm 구조:
/// - 열A: 의뢰날짜 (Excel 날짜 시리얼)
/// - 열B~: 바탕시험 데이터 및 참고값
/// - 각 시트별 시작 컬럼: 실제 시료분석 데이터 (반복 블록 구조)
/// </summary>
public static class XlsmDataMigration
{
    private static string _logPath = "";

    private static void Log(string msg)
    {
        Debug.WriteLine($"[XlsmDataMigration] {msg}");

        // BOD.log 파일에도 기록
        try
        {
            if (string.IsNullOrEmpty(_logPath))
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "ETA", "Data");
                Directory.CreateDirectory(logDir);
                _logPath = Path.Combine(logDir, "BOD.log");
            }

            File.AppendAllText(_logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}\n");
        }
        catch { }
    }

    public static void ExecutePhase2()
    {
        var xlsmPath = "Docs/CHUNGHA-김가린.xlsm";

        try
        {
            Console.WriteLine("\n[Phase 2] xlsm 데이터 마이그레이션");
            Console.WriteLine("".PadRight(70, '='));

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 생물화학적_산소요구량_시험기록부 테이블 생성 (없으면)
            try
            {
                Console.WriteLine("\n[Step 4] 생물화학적_산소요구량_시험기록부 테이블 확인");
                if (!DbConnectionFactory.TableExists(conn, "생물화학적_산소요구량_시험기록부"))
                {
                    Console.WriteLine("  테이블 없음, CREATE 중...");
                    using var createCmd = conn.CreateCommand();
                    createCmd.CommandText = $@"
                        CREATE TABLE `생물화학적_산소요구량_시험기록부` (
                            Id          INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                            분석일      TEXT NOT NULL,
                            SN          TEXT NOT NULL,
                            의뢰명      TEXT DEFAULT '',
                            시료명      TEXT DEFAULT '',
                            구분        TEXT DEFAULT '여수',
                            소스구분    TEXT DEFAULT '',
                            비고        TEXT DEFAULT '',
                            등록일시    TEXT DEFAULT '',
                            시료분석    TEXT DEFAULT '',
                            D1          TEXT DEFAULT '',
                            D2          TEXT DEFAULT '',
                            희석배수    TEXT DEFAULT '',
                            결과        TEXT DEFAULT '',
                            식종시료량  TEXT DEFAULT '',
                            식종D1      TEXT DEFAULT '',
                            식종D2      TEXT DEFAULT '',
                            식종BOD     TEXT DEFAULT '',
                            식종함유량  TEXT DEFAULT '',
                            업체명      TEXT DEFAULT ''
                        )";
                    createCmd.ExecuteNonQuery();
                    Console.WriteLine("✓ 생물화학적_산소요구량_시험기록부 CREATE 완료");
                    Log("생물화학적_산소요구량_시험기록부 CREATE 완료");
                }
                else
                {
                    Console.WriteLine("✓ 생물화학적_산소요구량_시험기록부 테이블 존재");
                }
            }
            catch (Exception e)
            {
                Log($"테이블 생성 오류: {e.Message}");
                Console.WriteLine($"✗ 테이블 생성 오류: {e.Message}");
            }

            if (!File.Exists(xlsmPath))
            {
                Console.WriteLine($"\n⊘ 파일 없음: {xlsmPath}");
                Console.WriteLine("✓ Phase 2 건너뜀");
                return;
            }

            using var fs = new FileStream(xlsmPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);

            int totalLoaded = 0;
            Console.WriteLine("\n[데이터 로드 중...]");

            // [5] BOD-DATA 마이그레이션 완료 — 재실행 불필요
            // totalLoaded += LoadBodData(wb);

            // [Phase 2-처리시설] 회사 서버에서 처리시설 분석결과 동기화
            // ⏳ 보류 중: 데이터 검증 후 활성화 예정
            /*
            Console.WriteLine("\n[Step 처리시설] 처리시설_결과 API 동기화");
            try
            {
                var today = DateTime.Now.ToString("yyyyMMdd");
                FacilityResultSyncService.SyncFacilityResultsAsync(today).GetAwaiter().GetResult();
                Console.WriteLine("✓ 처리시설_결과 동기화 완료");
            }
            catch (Exception ex)
            {
                Log($"처리시설 동기화 오류: {ex.Message}");
                Console.WriteLine($"⚠ 처리시설 동기화 실패: {ex.Message}");
            }
            */

            Console.WriteLine($"\n✓ Phase 2 완료: {totalLoaded}개 로드");
        }
        catch (Exception e)
        {
            Log($"마이그레이션 오류: {e.Message}\n{e.StackTrace}");
            Console.WriteLine($"✗ Phase 2 오류: {e.Message}");
        }
    }

    private static int LoadRequestNumberData(XLWorkbook wb)
    {
        const string sheetName = "의뢰번호";

        if (!wb.TryGetWorksheet(sheetName, out var ws))
        {
            Log($"시트 없음: {sheetName}");
            Console.WriteLine($"  ⊘ {sheetName}: 없음");
            return 0;
        }

        int loaded = 0;
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

        Log($"{sheetName}: 총행={lastRow}");

        // 의뢰번호 시트 구조:
        // A: 제수일
        // B,D,F,H...: 시료명 (홀수 컬럼)
        // C,E,G,I...: SN (짝수 컬럼)
        // 각 행에서 B~I까지 4쌍의 (시료명, SN) 튜플이 있음

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        for (int r = 2; r <= lastRow; r++)
        {
            string 채수일raw = GetCellValue(ws.Cell(r, 1)) ?? "";

            if (string.IsNullOrEmpty(채수일raw))
                break;

            if (!TryParseExcelDate(채수일raw, out var dateStr))
                continue;

            // 각 행에서 (업체명, SN) 쌍 추출: B=업체명,C=SN / D=업체명,E=SN / ...
            int[] 업체명Cols = { 2, 4, 6, 8, 10, 12, 14, 16, 18, 20 };  // B, D, F, H ...
            int[] snCols    = { 3, 5, 7, 9, 11, 13, 15, 17, 19, 21 };   // C, E, G, I ...

            for (int pairIdx = 0; pairIdx < 업체명Cols.Length; pairIdx++)
            {
                string 업체명 = GetCellValue(ws.Cell(r, 업체명Cols[pairIdx])) ?? "";
                string SN    = GetCellValue(ws.Cell(r, snCols[pairIdx])) ?? "";

                if (string.IsNullOrEmpty(SN))
                    break;

                try
                {
                    // SN 접두사로 구분 판별 및 접두사 제거
                    string 구분 = "여수";
                    string cleanSN = SN;

                    if (SN.StartsWith("[율촌]"))
                    {
                        구분 = "율촌";
                        cleanSN = SN.Substring(5);  // [율촌] 제거
                    }
                    else if (SN.StartsWith("[세풍]"))
                    {
                        구분 = "세풍";
                        cleanSN = SN.Substring(5);  // [세풍] 제거
                    }

                    // SN 형식 정규화: 5-26-7 → 05-26-07
                    cleanSN = NormalizeSN(cleanSN);

                    // 해당 날짜+구분의 다음 순서 번호
                    using var seqCmd = conn.CreateCommand();
                    seqCmd.CommandText = "SELECT COALESCE(MAX(순서),0)+1 FROM `비용부담금_결과` WHERE 채수일=@date AND 구분=@div";
                    seqCmd.Parameters.AddWithValue("@date", dateStr);
                    seqCmd.Parameters.AddWithValue("@div", 구분);
                    var nextSeq = Convert.ToInt32(seqCmd.ExecuteScalar());

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = "INSERT INTO `비용부담금_결과` (채수일, 구분, 순서, SN, 업체명) VALUES (@date, @div, @seq, @sn, @nm)";
                    cmd.Parameters.AddWithValue("@date", dateStr);
                    cmd.Parameters.AddWithValue("@div",  구분);
                    cmd.Parameters.AddWithValue("@seq",  nextSeq);
                    cmd.Parameters.AddWithValue("@sn",   cleanSN);
                    cmd.Parameters.AddWithValue("@nm",   업체명);
                    cmd.ExecuteNonQuery();

                    loaded++;
                    Log($"  행{r} 쌍{pairIdx}: 채수일={dateStr}, 구분={구분}, SN={SN}, 업체명={업체명}");
                }
                catch (Exception ex)
                {
                    Log($"{sheetName} 행{r} 쌍{pairIdx}: {ex.Message}");
                }
            }
        }

        Console.WriteLine($"  ✓ {sheetName}: {loaded}행");
        return loaded;
    }


    private static int LoadBodData(XLWorkbook wb)
    {
        const string sheetName = "BOD-DATA";
        const string tableName = "생물화학적_산소요구량_시험기록부";

        if (!wb.TryGetWorksheet(sheetName, out var ws))
        {
            Log($"시트 없음: {sheetName}");
            Console.WriteLine($"  ⊘ {sheetName}: 없음");
            return 0;
        }

        int loaded = 0;
        int lastRow = 2252; // 실제 데이터 마지막 행

        Log($"{sheetName}: 총행={lastRow}");

        // 데이터 행 순회 (행 2부터, 헤더 스킵)
        for (int r = 2; r <= lastRow; r++)
        {
            // 분석일자 (열A)
            string dateStr = GetCellValue(ws.Cell(r, 1));
            if (string.IsNullOrEmpty(dateStr)) continue;
            if (!TryParseExcelDate(dateStr, out var date)) continue;

            // 고정영역 (행당 공통)
            string 분석자1   = GetCellValue(ws.Cell(r, 2)) ?? "";
            string 분석자2   = GetCellValue(ws.Cell(r, 3)) ?? "";
            string 식종시료량 = GetCellValue(ws.Cell(r, 4)) ?? "";
            string min15DO  = GetCellValue(ws.Cell(r, 5)) ?? "";
            string day5DO   = GetCellValue(ws.Cell(r, 6)) ?? "";
            // 2022-10-05부터 식종액함유량 컬럼이 빠져 한 칸씩 당겨짐
            // 이전: G=식종BOD, H=식종함유량, I=희석수시료량, J=식종D1, K=식종D2
            // 이후: G=식종함유량, H=희석수시료량, I=식종D1, J=식종D2, 식종BOD=없음
            bool 컬럼이동 = DateTime.TryParse(date, out var 분석dt행) && 분석dt행 >= new DateTime(2022, 10, 5);
            string 식종BOD   = 컬럼이동 ? ""                                    : (GetCellValue(ws.Cell(r, 7))  ?? "");
            string 식종함유량  = 컬럼이동 ? (GetCellValue(ws.Cell(r, 7))  ?? "") : (GetCellValue(ws.Cell(r, 8))  ?? "");
            string 희석수시료량 = 컬럼이동 ? (GetCellValue(ws.Cell(r, 8))  ?? "") : (GetCellValue(ws.Cell(r, 9))  ?? "");
            string 식종D1    = 컬럼이동 ? (GetCellValue(ws.Cell(r, 9))  ?? "") : (GetCellValue(ws.Cell(r, 10)) ?? "");
            string 식종D2    = 컬럼이동 ? (GetCellValue(ws.Cell(r, 10)) ?? "") : (GetCellValue(ws.Cell(r, 11)) ?? "");

            // 샘플 블록 (col 30부터, 8컬럼 간격)
            for (int blockNum = 0; blockNum < 100; blockNum++)
            {
                int sampleNameCol = 30 + (blockNum * 8);
                int sampleQtyCol  = 31 + (blockNum * 8);
                int d1Col         = 32 + (blockNum * 8);
                int d2Col         = 33 + (blockNum * 8);
                int pCol          = 35 + (blockNum * 8);
                int resultCol     = 36 + (blockNum * 8);
                int snCol         = 37 + (blockNum * 8);

                if (sampleNameCol > 500) break;

                string sampleName = GetCellValue(ws.Cell(r, sampleNameCol));
                if (string.IsNullOrEmpty(sampleName)) break;

                string resultStr = GetCellValue(ws.Cell(r, resultCol));
                if (string.IsNullOrEmpty(resultStr) || resultStr == "0") continue;

                string snStr = GetCellValue(ws.Cell(r, snCol)) ?? "";

                // SN이 숫자만인 경우 skip
                if (string.IsNullOrEmpty(snStr) || double.TryParse(snStr, out _)) continue;

                try
                {
                    string sampleQty = GetCellValue(ws.Cell(r, sampleQtyCol)) ?? "";
                    string d1        = GetCellValue(ws.Cell(r, d1Col)) ?? "";
                    string d2        = GetCellValue(ws.Cell(r, d2Col)) ?? "";
                    string p         = GetCellValue(ws.Cell(r, pCol))  ?? "";

                    string 구분 = "여수";
                    string cleanSN = snStr;
                    if (snStr.StartsWith("[율촌]")) { 구분 = "율촌"; cleanSN = snStr.Substring(5); }
                    else if (snStr.StartsWith("[세풍]")) { 구분 = "세풍"; cleanSN = snStr.Substring(5); }
                    cleanSN = NormalizeSN(cleanSN);

                    Log($"  행{r} 블록{blockNum}: SN={cleanSN}, 구분={구분}, 시료명={sampleName}, 결과={resultStr}");

                    WasteSampleService.UpsertBodData(
                        tableName:    tableName,
                        채수일:       date,
                        sn:           cleanSN,
                        업체명:       sampleName,
                        구분:         구분,
                        시료량:       sampleQty,
                        d1:           d1,
                        d2:           d2,
                        희석배수:     p,
                        결과:         resultStr,
                        소스구분:     "폐수배출업소",
                        시료명:       sampleName,
                        식종시료량:   식종시료량,
                        식종D1:       식종D1,
                        식종D2:       식종D2,
                        식종BOD:      식종BOD,
                        식종함유량:   식종함유량,
                        분석자1:      분석자1,
                        분석자2:      분석자2,
                        minDO:        min15DO,
                        dayDO:        day5DO,
                        희석수시료량: 희석수시료량
                    );

                    // 비용부담금_결과 BOD UPDATE: 분석일 + SN 기준 채수일 역추적
                    WasteSampleService.UpdateWasteResult(date, cleanSN, sampleName, BOD: resultStr);

                    loaded++;
                }
                catch (Exception ex)
                {
                    Log($"{sheetName} 행{r} 블록{blockNum}: {ex.Message}");
                }
            }
        }
        Console.WriteLine($"  ✓ {sheetName}: {loaded}행");
        return loaded;
    }

    /// <summary>SN 형식 정규화: 5-26-7 → 05-26-07</summary>
    private static string NormalizeSN(string sn)
    {
        if (string.IsNullOrEmpty(sn))
            return sn;

        var parts = sn.Split('-');
        if (parts.Length != 3)
            return sn;

        try
        {
            int month = int.Parse(parts[0]);
            int day = int.Parse(parts[1]);
            int seq = int.Parse(parts[2]);

            return $"{month:D2}-{day:D2}-{seq:D2}";
        }
        catch
        {
            return sn;
        }
    }

    private static bool TryParseExcelDate(string dateStr, out string isoDate)
    {
        isoDate = "";

        if (string.IsNullOrEmpty(dateStr))
            return false;

        // 엑셀 날짜 시리얼 (1은 1900-01-01)
        if (double.TryParse(dateStr, out var serial))
        {
            try
            {
                var days = (int)serial;
                if (days > 59)
                    days -= 1; // 1900-02-29 버그 보정

                var baseDate = new DateTime(1899, 12, 30);
                var resultDate = baseDate.AddDays(days);
                isoDate = resultDate.ToString("yyyy-MM-dd");
                return true;
            }
            catch
            {
                return false;
            }
        }

        // 이미 yyyy-MM-dd 형식이면 그대로
        if (dateStr.Length == 10 && dateStr[4] == '-' && dateStr[7] == '-')
        {
            isoDate = dateStr;
            return true;
        }

        // "2020. 1. 1. 오전 12:00:00" 형식 처리
        // "yyyy. M. d. 오전/오후 HH:mm:ss"
        try
        {
            // ". 오" 패턴으로 시간 부분 제거
            int ampmPos = dateStr.IndexOf(". 오");
            if (ampmPos > 0)
            {
                string datePart = dateStr.Substring(0, ampmPos); // "2020. 1. 1"

                // "2020. 1. 1" → "2020-01-01"
                var parts = datePart.Split(new[] { ". " }, StringSplitOptions.None);
                if (parts.Length == 3 &&
                    int.TryParse(parts[0], out var year) &&
                    int.TryParse(parts[1], out var month) &&
                    int.TryParse(parts[2], out var day))
                {
                    isoDate = new DateTime(year, month, day).ToString("yyyy-MM-dd");
                    return true;
                }
            }
        }
        catch { }

        return false;
    }

    private static string? GetCellValue(IXLCell cell)
    {
        try
        {
            if (cell == null)
                return null;

            // 문자열 타입
            if (cell.DataType == XLDataType.Text)
            {
                var val = cell.GetString();
                return string.IsNullOrEmpty(val) ? null : val.Trim();
            }

            // 숫자 타입
            if (cell.DataType == XLDataType.Number)
            {
                try
                {
                    var numVal = cell.GetDouble();
                    if (double.IsNaN(numVal) || double.IsInfinity(numVal))
                        return null;

                    // 정수인 경우
                    if (numVal == (int)numVal)
                        return ((int)numVal).ToString();
                    else
                        return numVal.ToString("G");
                }
                catch
                {
                    return null;
                }
            }

            // 기타 타입 또는 빈 셀
            var strVal = cell.GetString();
            return string.IsNullOrEmpty(strVal) ? null : strVal.Trim();
        }
        catch
        {
            return null;
        }
    }
}
