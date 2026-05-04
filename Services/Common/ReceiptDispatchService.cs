using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ETA.Models;

namespace ETA.Services.Common;

/// <summary>측정대행 접수/발송 대장 — 독립 테이블 (엑셀 `접수발송대장` 시트 1:1 매핑)</summary>
public static class ReceiptDispatchService
{
    private const string TableName = "접수발송대장";

    public static void EnsureTable(System.Data.Common.DbConnection conn)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                CREATE TABLE IF NOT EXISTS `{TableName}` (
                    Id              INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                    접수번호        TEXT DEFAULT '',
                    접수일          TEXT DEFAULT '',
                    시료명          TEXT DEFAULT '',
                    의뢰인및업체명  TEXT DEFAULT '',
                    분석항목        TEXT DEFAULT '',
                    발송일          TEXT DEFAULT '',
                    등록일시        TEXT DEFAULT ''
                )";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Debug.WriteLine($"EnsureTable 오류: {ex.Message}"); }
    }

    public static List<ReceiptDispatchEntry> GetAll()
    {
        var list = new List<ReceiptDispatchEntry>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            SELECT Id, COALESCE(접수번호,''), COALESCE(접수일,''), COALESCE(시료명,''),
                   COALESCE(의뢰인및업체명,''), COALESCE(분석항목,''), COALESCE(발송일,'')
            FROM `{TableName}`
            ORDER BY 접수일 DESC, Id DESC";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ReceiptDispatchEntry
            {
                Id = Convert.ToInt32(r.GetValue(0)),
                접수번호 = S(r, 1),
                접수일   = S(r, 2),
                시료명   = S(r, 3),
                업체명   = S(r, 4),  // ReceiptDispatchEntry 의 업체명 필드를 의뢰인및업체명 로 사용
                분석항목 = S(r, 5),
                발송일   = S(r, 6),
            });
        }
        return list;
    }

    public static int Insert(ReceiptDispatchEntry e)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO `{TableName}` (접수번호, 접수일, 시료명, 의뢰인및업체명, 분석항목, 발송일, 등록일시)
            VALUES (@n, @d, @s, @c, @i, @sd, @t)";
        cmd.Parameters.AddWithValue("@n",  e.접수번호 ?? "");
        cmd.Parameters.AddWithValue("@d",  e.접수일   ?? "");
        cmd.Parameters.AddWithValue("@s",  e.시료명   ?? "");
        cmd.Parameters.AddWithValue("@c",  e.업체명   ?? "");
        cmd.Parameters.AddWithValue("@i",  e.분석항목 ?? "");
        cmd.Parameters.AddWithValue("@sd", e.발송일   ?? "");
        cmd.Parameters.AddWithValue("@t",  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();

        using var idCmd = conn.CreateCommand();
        idCmd.CommandText = $"SELECT {DbConnectionFactory.LastInsertId}";
        return Convert.ToInt32(idCmd.ExecuteScalar());
    }

    public static bool Update(ReceiptDispatchEntry e)
    {
        if (e.Id <= 0) return false;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            UPDATE `{TableName}` SET
                접수번호=@n, 접수일=@d, 시료명=@s, 의뢰인및업체명=@c, 분석항목=@i, 발송일=@sd
            WHERE Id=@id";
        cmd.Parameters.AddWithValue("@n",  e.접수번호 ?? "");
        cmd.Parameters.AddWithValue("@d",  e.접수일   ?? "");
        cmd.Parameters.AddWithValue("@s",  e.시료명   ?? "");
        cmd.Parameters.AddWithValue("@c",  e.업체명   ?? "");
        cmd.Parameters.AddWithValue("@i",  e.분석항목 ?? "");
        cmd.Parameters.AddWithValue("@sd", e.발송일   ?? "");
        cmd.Parameters.AddWithValue("@id", e.Id);
        return cmd.ExecuteNonQuery() > 0;
    }

    public static bool Delete(int id)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"DELETE FROM `{TableName}` WHERE Id=@id";
        cmd.Parameters.AddWithValue("@id", id);
        return cmd.ExecuteNonQuery() > 0;
    }

    /// <summary>전체 행 삭제 — 처음부터 다시 작성하고 싶을 때</summary>
    public static int DeleteAll()
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureTable(conn);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"DELETE FROM `{TableName}`";
        return cmd.ExecuteNonQuery();
    }

    /// <summary>견적번호 → 견적요청담당 매핑 (견적발행내역 테이블 1회 스캔)</summary>
    private static Dictionary<string, string> LoadQuoteManagerMap(System.Data.Common.DbConnection conn)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            if (!DbConnectionFactory.TableExists(conn, "견적발행내역")) return map;
            var cols = DbConnectionFactory.GetColumnNames(conn, "견적발행내역");
            // 견적요청담당/담당자 컬럼 이름 후보
            string? managerCol = new[] { "견적요청담당", "견적요청 담당자", "담당자", "요청담당자" }
                .FirstOrDefault(c => cols.Any(x => string.Equals(x, c, StringComparison.OrdinalIgnoreCase)));
            if (managerCol == null) return map;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT COALESCE(`견적번호`,''), COALESCE(`{managerCol}`,'') FROM `견적발행내역`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var k = r.GetValue(0)?.ToString() ?? "";
                var v = r.GetValue(1)?.ToString() ?? "";
                if (!string.IsNullOrEmpty(k) && !string.IsNullOrEmpty(v))
                    map[k] = v;
            }
        }
        catch { }
        return map;
    }

    /// <summary>의뢰사업장(회사 전체이름) + 견적요청담당 → "회사 / 담당자" 포맷</summary>
    public static string ComputeRequesterLabel(int rowId)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT COALESCE(`의뢰사업장`,''), COALESCE(`약칭`,''), COALESCE(`견적번호`,'')
                                 FROM `수질분석센터_결과` WHERE {DbConnectionFactory.RowId} = @id LIMIT 1";
            cmd.Parameters.AddWithValue("@id", rowId);
            string company = "", abbr = "", quote = "";
            using (var r = cmd.ExecuteReader())
            {
                if (r.Read()) { company = S(r,0); abbr = S(r,1); quote = S(r,2); }
            }
            var label = string.IsNullOrEmpty(company) ? abbr : company;
            var managerMap = LoadQuoteManagerMap(conn);
            if (!string.IsNullOrEmpty(quote) && managerMap.TryGetValue(quote, out var mgr) && !string.IsNullOrWhiteSpace(mgr))
                label = $"{label} / {mgr}";
            return label;
        }
        catch (Exception ex) { Debug.WriteLine($"ComputeRequesterLabel 오류: {ex.Message}"); return ""; }
    }

    /// <summary>특정 의뢰(rowId) 의 "O" 표시된 분석항목 → "XX 외 N건" 포맷 (없으면 빈 문자열)</summary>
    public static string ComputeAnalyteSummary(int rowId)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "수질분석센터_결과")) return "";
            var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id","Id","채취일자","채취시간","의뢰사업장","약칭","시료명",
                "견적번호","입회자","시료채취자-1","시료채취자-2",
                "방류허용기준 적용유무","정도보증유무","정도보증",
                "분석완료일자","분석종료일","견적구분",
                "시료유형","접수일자","접수담당자","업체담당자",
                "현장_온도","현장_pH","현장_용존산소","현장_전기전도도","현장_잔류염소",
                "생태_염분","생태_암모니아","생태_경도","발송일",
            };
            var allCols = DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과");
            var analyteCols = allCols.Where(c => !fixedCols.Contains(c)).ToList();
            if (analyteCols.Count == 0) return "";

            var rid = DbConnectionFactory.RowId;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT {string.Join(",", analyteCols.Select(c => $"`{c}`"))} " +
                              $"FROM `수질분석센터_결과` WHERE {rid} = @id LIMIT 1";
            cmd.Parameters.AddWithValue("@id", rowId);
            using var rdr = cmd.ExecuteReader();
            if (!rdr.Read()) return "";

            var marked = new List<string>();
            for (int i = 0; i < analyteCols.Count; i++)
            {
                if (rdr.IsDBNull(i)) continue;
                var v = rdr.GetValue(i)?.ToString()?.Trim() ?? "";
                // "O"(원래 의뢰 표시) 또는 결과값(분석완료 후 채워진 숫자/문자) 모두 → 의뢰된 항목
                if (!string.IsNullOrEmpty(v))
                    marked.Add(analyteCols[i]);
            }
            return marked.Count switch
            {
                0 => "",
                1 => marked[0],
                _ => $"{marked[0]} 외 {marked.Count - 1}건",
            };
        }
        catch (Exception ex) { Debug.WriteLine($"ComputeAnalyteSummary 오류: {ex.Message}"); return ""; }
    }

    /// <summary>엑셀 `접수발송정리` 메크로 재현 — 선택된 채취일자의 모든 의뢰를 대장에 일괄 추가.
    /// 접수번호: renewus{YYMMDD}-{seq:000}-A (해당 날짜 시퀀스)
    /// 분석항목: 수질분석센터_결과 의 "O" 마킹 컬럼 → 1개면 항목명, 다수면 "{첫항목}외 N-1건"
    /// 의뢰사업장 → 의뢰인및업체명, 시료명 그대로.</summary>
    public static int GenerateForDate(string 채취일자)
    {
        if (string.IsNullOrWhiteSpace(채취일자)) return 0;
        var iso = 채취일자.Length >= 10 ? 채취일자[..10] : 채취일자;
        int count = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTable(conn);

            // 분석항목 컬럼 후보 — 고정 컬럼 제외한 나머지를 분석항목 컬럼으로 처리
            var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id","Id","채취일자","채취시간","의뢰사업장","약칭","시료명",
                "견적번호","입회자","시료채취자-1","시료채취자-2",
                "방류허용기준 적용유무","정도보증유무","정도보증",
                "분석완료일자","분석종료일","견적구분",
                "시료유형","접수일자","접수담당자","업체담당자",
                "현장_온도","현장_pH","현장_용존산소","현장_전기전도도","현장_잔류염소",
                "생태_염분","생태_암모니아","생태_경도","발송일",
            };
            var allCols = DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과");
            var analyteCols = allCols.Where(c => !fixedCols.Contains(c)).ToList();

            var rid = DbConnectionFactory.RowId;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT {rid},
                COALESCE(`채취일자`,''), COALESCE(`시료명`,''),
                COALESCE(`의뢰사업장`,''), COALESCE(`약칭`,''),
                COALESCE(`견적번호`,'')
                {(analyteCols.Count > 0 ? "," + string.Join(",", analyteCols.Select(c => $"`{c}`")) : "")}
                FROM `수질분석센터_결과`
                WHERE `채취일자` LIKE @date
                ORDER BY {rid} ASC";
            cmd.Parameters.AddWithValue("@date", iso + "%");

            // YYMMDD prefix
            string yymmdd = iso.Replace("-", "").Substring(2, 6);
            int seq = 1;

            // 견적번호 → 견적요청담당 매핑 캐시
            var managerByQuote = LoadQuoteManagerMap(conn);

            using var rdr = cmd.ExecuteReader();
            var batch = new List<ReceiptDispatchEntry>();
            while (rdr.Read())
            {
                string rDate    = S(rdr, 1).Length >= 10 ? S(rdr, 1)[..10] : S(rdr, 1);
                string rSample  = S(rdr, 2);
                string rCompany = S(rdr, 3);
                string rAbbr    = S(rdr, 4);
                string rQuote   = S(rdr, 5);

                // 분석항목 — 비어있지 않은 모든 분석항목 컬럼 (offset = 6)
                // "O" (원래 의뢰 표시) 또는 결과값 (분석완료 후 숫자/문자) 모두 의뢰된 것으로 간주
                var marked = new List<string>();
                for (int i = 0; i < analyteCols.Count; i++)
                {
                    int idx = 6 + i;
                    if (rdr.IsDBNull(idx)) continue;
                    var v = rdr.GetValue(idx)?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(v))
                        marked.Add(analyteCols[i]);
                }
                string 분석항목 = marked.Count switch
                {
                    0 => "",
                    1 => marked[0],
                    _ => $"{marked[0]} 외 {marked.Count - 1}건",
                };

                // 의뢰인및업체명 = "회사 전체이름 / 견적요청담당"
                var company = string.IsNullOrEmpty(rCompany) ? rAbbr : rCompany;
                managerByQuote.TryGetValue(rQuote, out var manager);
                var 의뢰인및업체명 = string.IsNullOrWhiteSpace(manager)
                    ? company
                    : $"{company} / {manager}";

                batch.Add(new ReceiptDispatchEntry
                {
                    접수번호 = $"renewus{yymmdd}-{seq:D3}-A",
                    접수일   = rDate,
                    시료명   = rSample,
                    업체명   = 의뢰인및업체명,
                    분석항목 = 분석항목,
                    발송일   = "",
                });
                seq++;
            }

            foreach (var e in batch)
            {
                Insert(e);
                count++;
            }
        }
        catch (Exception ex) { Debug.WriteLine($"GenerateForDate 오류: {ex.Message}"); }
        return count;
    }

    /// <summary>분석완료일자 → 발송일 일괄 갱신 — 메크로 `발송일확인` 재현.
    /// 같은 채취일자+시료명 의뢰의 `분석완료일자` 컬럼에서 MAX 일자를 읽어 "YYYY-MM-DD(요일)" 포맷으로 저장.</summary>
    public static int RecomputeDispatchDates(string 접수일)
    {
        if (string.IsNullOrWhiteSpace(접수일)) return 0;
        int updated = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTable(conn);

            // 수질분석센터_결과 의 실제 컬럼명 — 분석완료일자 또는 분석종료일 둘 중 존재하는 것
            var sourceCols = DbConnectionFactory.GetColumnNames(conn, "수질분석센터_결과");
            string? completionCol = new[] { "분석완료일자", "분석종료일" }
                .FirstOrDefault(c => sourceCols.Any(x => string.Equals(x, c, StringComparison.OrdinalIgnoreCase)));
            if (completionCol == null) return 0;

            using var qcmd = conn.CreateCommand();
            qcmd.CommandText = $"SELECT Id, COALESCE(시료명,'') FROM `{TableName}` WHERE 접수일 = @d";
            qcmd.Parameters.AddWithValue("@d", 접수일);
            var rows = new List<(int Id, string Sample)>();
            using (var r = qcmd.ExecuteReader())
                while (r.Read()) rows.Add((Convert.ToInt32(r.GetValue(0)), S(r, 1)));

            foreach (var (id, sample) in rows)
            {
                using var dcmd = conn.CreateCommand();
                dcmd.CommandText = $@"SELECT MAX(COALESCE(`{completionCol}`, ''))
                                     FROM `수질분석센터_결과`
                                     WHERE `채취일자` LIKE @d AND `시료명` = @s";
                dcmd.Parameters.AddWithValue("@d", 접수일 + "%");
                dcmd.Parameters.AddWithValue("@s", sample);
                var v = dcmd.ExecuteScalar();
                var raw = v?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(raw)) continue;

                string formatted = FormatDispatchDate(raw);
                if (string.IsNullOrEmpty(formatted)) continue;

                using var ucmd = conn.CreateCommand();
                ucmd.CommandText = $"UPDATE `{TableName}` SET 발송일 = @v WHERE Id = @id";
                ucmd.Parameters.AddWithValue("@v", formatted);
                ucmd.Parameters.AddWithValue("@id", id);
                if (ucmd.ExecuteNonQuery() > 0) updated++;
            }
        }
        catch (Exception ex) { Debug.WriteLine($"RecomputeDispatchDates 오류: {ex.Message}"); }
        return updated;
    }

    /// <summary>"YYYY-MM-DD(요일)" 포맷 — 엑셀 메크로 Format(d, "YYYY-MM-DD(AAA)") 와 동일</summary>
    private static string FormatDispatchDate(string raw)
    {
        var s = raw.Length >= 10 ? raw[..10] : raw;
        if (DateTime.TryParse(s, out var dt))
        {
            string[] dows = { "일", "월", "화", "수", "목", "금", "토" };
            return $"{dt:yyyy-MM-dd}({dows[(int)dt.DayOfWeek]})";
        }
        return raw;
    }

    /// <summary>엑셀 `접수발송대장` 시트의 6컬럼 행을 일괄 import — 기존 데이터는 유지(추가만).</summary>
    public static int ImportFromExcel(string xlsxPath)
    {
        if (!System.IO.File.Exists(xlsxPath)) return 0;
        int count = 0;
        try
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(xlsxPath);
            var ws = wb.Worksheet("접수발송대장");
            if (ws == null) return 0;

            // 헤더 행 찾기 — "접수번호" 가 들어있는 행
            int headerRow = 1;
            for (int r = 1; r <= 5; r++)
            {
                var v = ws.Cell(r, 1).GetString().Trim();
                if (v == "접수번호") { headerRow = r; break; }
            }

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTable(conn);

            int last = ws.LastRowUsed()?.RowNumber() ?? headerRow;
            for (int r = headerRow + 1; r <= last; r++)
            {
                var no = ws.Cell(r, 1).GetString().Trim();
                if (string.IsNullOrEmpty(no)) continue;
                // 푸터/서명 행 제외 — "시료접수", "리뉴어스", "담당자", "기술책임자", "품질책임자" 등
                if (IsFooterText(no)) continue;
                var e = new ReceiptDispatchEntry
                {
                    접수번호 = no,
                    접수일   = FormatDate(ws.Cell(r, 2)),
                    시료명   = ws.Cell(r, 3).GetString().Trim(),
                    업체명   = ws.Cell(r, 4).GetString().Trim().Replace("\n", " "),
                    분석항목 = ws.Cell(r, 5).GetString().Trim(),
                    발송일   = FormatDate(ws.Cell(r, 6)),
                };
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"
                    INSERT INTO `{TableName}` (접수번호, 접수일, 시료명, 의뢰인및업체명, 분석항목, 발송일, 등록일시)
                    VALUES (@n, @d, @s, @c, @i, @sd, @t)";
                cmd.Parameters.AddWithValue("@n",  e.접수번호);
                cmd.Parameters.AddWithValue("@d",  e.접수일);
                cmd.Parameters.AddWithValue("@s",  e.시료명);
                cmd.Parameters.AddWithValue("@c",  e.업체명);
                cmd.Parameters.AddWithValue("@i",  e.분석항목);
                cmd.Parameters.AddWithValue("@sd", e.발송일);
                cmd.Parameters.AddWithValue("@t",  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.ExecuteNonQuery();
                count++;
            }
        }
        catch (Exception ex) { Debug.WriteLine($"ImportFromExcel 오류: {ex.Message}"); }
        return count;
    }

    /// <summary>접수번호 셀 내용이 푸터/서명 텍스트인지 — 데이터로 가져오면 안 되는 항목 차단.</summary>
    private static bool IsFooterText(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return true;
        var t = text.Trim();
        string[] patterns =
        {
            "시료접수", "시험성적서 발송", "담  당  자", "담 당 자", "담당자",
            "기술책임자", "품질책임자", "(서명)", "리뉴어스", "수질분석센터",
        };
        return patterns.Any(p => t.Contains(p));
    }

    /// <summary>이미 DB 에 들어있는 푸터/서명 잔존 행을 일괄 정리</summary>
    public static int PurgeFooterRows()
    {
        int n = 0;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            EnsureTable(conn);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT Id, COALESCE(접수번호,''), COALESCE(시료명,''), COALESCE(의뢰인및업체명,'') FROM `{TableName}`";
            var idsToDelete = new List<int>();
            using (var r = cmd.ExecuteReader())
                while (r.Read())
                {
                    int id = Convert.ToInt32(r.GetValue(0));
                    string a = S(r, 1), b = S(r, 2), c = S(r, 3);
                    if (IsFooterText(a) || IsFooterText(b) || IsFooterText(c))
                        idsToDelete.Add(id);
                }
            foreach (var id in idsToDelete)
            {
                using var d = conn.CreateCommand();
                d.CommandText = $"DELETE FROM `{TableName}` WHERE Id = @id";
                d.Parameters.AddWithValue("@id", id);
                if (d.ExecuteNonQuery() > 0) n++;
            }
        }
        catch (Exception ex) { Debug.WriteLine($"PurgeFooterRows 오류: {ex.Message}"); }
        return n;
    }

    private static string FormatDate(ClosedXML.Excel.IXLCell cell)
    {
        try
        {
            if (cell.DataType == ClosedXML.Excel.XLDataType.DateTime)
                return cell.GetDateTime().ToString("yyyy-MM-dd");
            var s = cell.GetString().Trim();
            if (DateTime.TryParse(s, out var dt)) return dt.ToString("yyyy-MM-dd");
            return s;
        }
        catch { return ""; }
    }

    private static string S(System.Data.Common.DbDataReader r, int i) =>
        r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
}
