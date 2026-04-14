using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services.Common;
using ClosedXML.Excel;

namespace ETA.Services.SERVICE1;

public static class ContractService
{
    // ── DB 경로 (AgentService 와 동일 방식) ──────────────────────────────────

    // ── 전체 조회 ─────────────────────────────────────────────────────────────
    public static List<Contract> GetAllContracts()
    {
        var list   = new List<Contract>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                   C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                   C_Address, C_Representative, C_FacilityType, C_CategoryType,
                   C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email
            FROM `계약 DB`
            ORDER BY C_CompanyName ASC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var c = new Contract
            {
                C_CompanyName    = S(r, "C_CompanyName"),
                OriginalCompanyName = S(r, "C_CompanyName"),
                C_Abbreviation   = S(r, "C_Abbreviation"),
                C_ContractType   = S(r, "C_ContractType"),
                C_Address        = S(r, "C_Address"),
                C_Representative = S(r, "C_Representative"),
                C_FacilityType   = S(r, "C_FacilityType"),
                C_CategoryType   = S(r, "C_CategoryType"),
                C_MainProduct    = S(r, "C_MainProduct"),
                C_ContactPerson  = S(r, "C_ContactPerson"),
                C_PhoneNumber    = S(r, "C_PhoneNumber"),
                C_Email          = S(r, "C_Email"),
                C_ContractDays   = NullInt(r, "C_ContractDays"),
                C_ContractAmountVATExcluded = NullDecimal(r, "C_ContractAmountVATExcluded"),
            };

            var startStr = S(r, "C_ContractStart");
            if (DateTime.TryParse(startStr, out var start)) c.C_ContractStart = start;

            var endStr = S(r, "C_ContractEnd");
            if (DateTime.TryParse(endStr, out var end)) c.C_ContractEnd = end;

            list.Add(c);
        }

        return list;
    }

    // ── 수정 ─────────────────────────────────────────────────────────────────
    public static bool Update(Contract contract)
    {
        if (string.IsNullOrEmpty(contract.OriginalCompanyName))
        {
        }

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            UPDATE `계약 DB` SET
                C_CompanyName=@name, C_ContractStart=@start, C_ContractEnd=@end,
                C_ContractDays=@days, C_ContractAmountVATExcluded=@amount,
                C_Abbreviation=@abbr, C_ContractType=@type, C_Address=@addr,
                C_Representative=@rep, C_FacilityType=@ftype, C_CategoryType=@ctype,
                C_MainProduct=@prod, C_ContactPerson=@contact,
                C_PhoneNumber=@phone, C_Email=@email
            WHERE C_CompanyName=@original";

        SetParams(cmd, contract);
        cmd.Parameters.AddWithValue("@original", contract.OriginalCompanyName);

        int rows = cmd.ExecuteNonQuery();
        if (rows > 0) { contract.OriginalCompanyName = contract.C_CompanyName; return true; }
        return false;
    }

    // ── 추가 ─────────────────────────────────────────────────────────────────
    public static bool Insert(Contract contract)
    {
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO `계약 DB`
                (C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                 C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                 C_Address, C_Representative, C_FacilityType, C_CategoryType,
                 C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email)
            VALUES
                (@name, @start, @end, @days, @amount, @abbr, @type, @addr,
                 @rep, @ftype, @ctype, @prod, @contact, @phone, @email)";

        SetParams(cmd, contract);
        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
    }

    // ── 업체명으로 대표자 조회 ────────────────────────────────────────────────
    public static string GetRepresentativeByCompany(string companyName)
    {
        if (string.IsNullOrWhiteSpace(companyName)) return "";
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT C_Representative FROM `계약 DB` WHERE C_CompanyName=@name LIMIT 1";
        cmd.Parameters.AddWithValue("@name", companyName);
        return cmd.ExecuteScalar()?.ToString() ?? "";
    }

    // ── 삭제 ─────────────────────────────────────────────────────────────────
    public static bool Delete(Contract contract)
    {
        if (string.IsNullOrEmpty(contract.C_CompanyName)) return false;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"DELETE FROM `계약 DB` WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", contract.C_CompanyName);

        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
    }

    // ── 분석정보 항목 목록 조회 ───────────────────────────────────────────────
    /// <summary>분석정보 테이블 Analyte 컬럼(=C컬럼) 항목명 목록을 순서대로 반환합니다.</summary>

    // ── 분석정보 Analyte → 약칭 맵 조회 ───────────────────────────────
    /// <summary>분석정보 테이블에서 Analyte → 약칭 Dictionary를 반환합니다.</summary>
    public static Dictionary<string, string> GetAnalyteAliasMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            // 약칭 컬럼 없으면 자동 추가
            if (!DbConnectionFactory.ColumnExists(conn, "분석정보", "약칭"))
            {
                using var alt = conn.CreateCommand();
                try { alt.CommandText = "ALTER TABLE `분석정보` ADD COLUMN `약칭` TEXT DEFAULT ''"; alt.ExecuteNonQuery(); } catch { }
            }
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `Analyte`, COALESCE(`약칭`, '') FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte` <> '' ORDER BY ES ASC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var analyte = r.IsDBNull(0) ? "" : r.GetString(0);
                var alias   = r.IsDBNull(1) ? "" : r.GetString(1) ?? "";
                if (!string.IsNullOrWhiteSpace(analyte))
                    map[analyte] = alias;
            }
        }
        catch { }
        return map;
    }

    public static List<string> GetAnalysisItems()
    {
        var result = new List<string>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT `Analyte` FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte` <> '' ORDER BY ES ASC";
        try
        {
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var v = r.IsDBNull(0) ? "" : r.GetString(0);
                if (!string.IsNullOrWhiteSpace(v)) result.Add(v);
            }
        }
        catch (Exception ex) { }

        return result;
    }

    // ── 계약 DB에 분석단가 컬럼 보장 ─────────────────────────────────────────
    /// <summary>분석정보 DB 항목을 기반으로 `계약 DB`에 단가 컬럼을 보장합니다.</summary>
    public static void EnsureContractPriceColumns() =>
        EnsureContractPriceColumns(GetAnalysisItems());

    /// <summary>지정된 컬럼명 목록을 `계약 DB`에 없으면 ALTER TABLE로 추가합니다.</summary>
    public static void EnsureContractPriceColumns(IEnumerable<string> priceColumns)
    {
        var items = priceColumns.ToList();
        if (items.Count == 0) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var existing = DbConnectionFactory.GetColumnNames(conn, "계약 DB");

        foreach (var item in items)
        {
            if (existing.Contains(item)) continue;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"ALTER TABLE `계약 DB` ADD COLUMN IF NOT EXISTS `{item}` DECIMAL(15,2) NULL DEFAULT NULL";

            try { cmd.ExecuteNonQuery(); }
            catch (Exception ex) { }
        }
    }

    // ── Excel → DB 임포트 ─────────────────────────────────────────────────────
    /// <summary>
    /// ETA DB.xlsm 의 '계약 DB' 시트를 읽어 MariaDB에 UPSERT합니다.<br/>
    /// 컬럼 매핑: 번호(skip), 업체명(키), 계약시작, 계약종료, 계약일수,
    /// 계약금액(VAT별도), 약칭, 계약구분, 주소, 대표자, 시설별, 종류별,
    /// 주생산품, 담당자, 연락처, 이메일, [17열~] 분석단가 항목들
    /// </summary>
    /// <returns>(성공 건수, 오류 건수)</returns>
    public static (int ok, int err) ImportFromExcel(string filePath)
    {
        int ok = 0, err = 0;

        if (!File.Exists(filePath))
        { return (0, 0); }

        List<string>                  headers;
        List<List<string>>            rows;

        // ── Excel 읽기 ──────────────────────────────────────────────────────
        try
        {
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheets.FirstOrDefault(s => s.Name == "계약 DB")
                     ?? throw new Exception("'계약 DB' 시트 없음");

            var usedRange = ws.RangeUsed();
            if (usedRange == null) return (0, 0);

            int lastRow = usedRange.LastRow().RowNumber();
            int lastCol = usedRange.LastColumn().ColumnNumber();

            // 헤더 행
            headers = Enumerable.Range(1, lastCol)
                                 .Select(c => ws.Cell(1, c).GetString().Trim())
                                 .ToList();

            // 데이터 행
            rows = new List<List<string>>();
            for (int r = 2; r <= lastRow; r++)
            {
                var row = Enumerable.Range(1, lastCol)
                                    .Select(c =>
                                    {
                                        var cell = ws.Cell(r, c);
                                        // 날짜 셀은 DateTime 으로 변환
                                        if (cell.DataType == XLDataType.DateTime)
                                            return cell.GetDateTime().ToString("yyyy-MM-dd");
                                        // 숫자(날짜 시리얼 포함) 방어 처리
                                        if (cell.DataType == XLDataType.Number)
                                        {
                                            // Excel 날짜 시리얼 판정 (포맷에 날짜 형식 포함)
                                            var fmt = cell.Style.NumberFormat.Format;
                                            if (fmt.Contains("yy") || fmt.Contains("mm") || fmt.Contains("dd"))
                                                return DateTime.FromOADate(cell.GetDouble()).ToString("yyyy-MM-dd");
                                        }
                                        return cell.GetString().Trim();
                                    })
                                    .ToList();
                rows.Add(row);
            }
        }
        catch (Exception ex)
        {
            return (0, -1);
        }

        // 헤더에서 단가 컬럼 추출 (17열 이후)
        var priceColNames = headers.Skip(16).Where(h => !string.IsNullOrWhiteSpace(h)).ToList();
        if (priceColNames.Count > 0)
            EnsureContractPriceColumns(priceColNames);

        // ── DB UPSERT ───────────────────────────────────────────────────────
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var existingDbCols = DbConnectionFactory.GetColumnNames(conn, "계약 DB").ToHashSet();

        foreach (var row in rows)
        {
            if (row.Count < 2) continue;
            string companyName = row[1];          // 열2 = 업체명
            if (string.IsNullOrWhiteSpace(companyName)) continue;

            try
            {
                // 기본정보 (열2~16)
                string start   = row.Count > 2  ? row[2]  : "";
                string end     = row.Count > 3  ? row[3]  : "";
                string days    = row.Count > 4  ? row[4]  : "";
                string amount  = row.Count > 5  ? row[5]  : "";
                string abbr    = row.Count > 6  ? row[6]  : "";
                string ctype   = row.Count > 7  ? row[7]  : "";
                string addr    = row.Count > 8  ? row[8]  : "";
                string rep     = row.Count > 9  ? row[9]  : "";
                string ftype   = row.Count > 10 ? row[10] : "";
                string cat     = row.Count > 11 ? row[11] : "";
                string prod    = row.Count > 12 ? row[12] : "";
                string contact = row.Count > 13 ? row[13] : "";
                string phone   = row.Count > 14 ? row[14] : "";
                string email   = row.Count > 15 ? row[15] : "";

                // 단가 컬럼 (열17+)
                var pricePairs = new List<(string col, string val)>();
                for (int i = 16; i < row.Count && i - 16 < priceColNames.Count; i++)
                {
                    var col = priceColNames[i - 16];
                    if (existingDbCols.Contains(col))
                        pricePairs.Add((col, row[i]));
                }

                // 존재 여부 확인
                bool exists;
                using (var chk = conn.CreateCommand())
                {
                    chk.CommandText = "SELECT COUNT(*) FROM `계약 DB` WHERE C_CompanyName=@n";
                    chk.Parameters.AddWithValue("@n", companyName);
                    exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;
                }

                if (exists)
                {
                    // UPDATE 기본정보 + 단가
                    var setParts = new List<string>
                    {
                        "C_ContractStart=@start","C_ContractEnd=@end","C_ContractDays=@days",
                        "C_ContractAmountVATExcluded=@amount","C_Abbreviation=@abbr",
                        "C_ContractType=@ctype","C_Address=@addr","C_Representative=@rep",
                        "C_FacilityType=@ftype","C_CategoryType=@cat","C_MainProduct=@prod",
                        "C_ContactPerson=@contact","C_PhoneNumber=@phone","C_Email=@email"
                    };
                    for (int p = 0; p < pricePairs.Count; p++) setParts.Add($"`{pricePairs[p].col}`=@pp{p}");

                    using var upd = conn.CreateCommand();
                    upd.CommandText = $"UPDATE `계약 DB` SET {string.Join(",", setParts)} WHERE C_CompanyName=@n";
                    AddBasicParams(upd, companyName, start, end, days, amount, abbr, ctype, addr, rep, ftype, cat, prod, contact, phone, email);
                    for (int p = 0; p < pricePairs.Count; p++)
                        AddPriceParam(upd, $"@pp{p}", pricePairs[p].val);
                    upd.ExecuteNonQuery();
                }
                else
                {
                    // INSERT 기본정보
                    using var ins = conn.CreateCommand();
                    ins.CommandText = @"INSERT INTO `계약 DB`
                        (C_CompanyName,C_ContractStart,C_ContractEnd,C_ContractDays,
                         C_ContractAmountVATExcluded,C_Abbreviation,C_ContractType,C_Address,
                         C_Representative,C_FacilityType,C_CategoryType,C_MainProduct,
                         C_ContactPerson,C_PhoneNumber,C_Email)
                        VALUES(@n,@start,@end,@days,@amount,@abbr,@ctype,@addr,@rep,@ftype,@cat,@prod,@contact,@phone,@email)";
                    AddBasicParams(ins, companyName, start, end, days, amount, abbr, ctype, addr, rep, ftype, cat, prod, contact, phone, email);
                    ins.ExecuteNonQuery();

                    // UPDATE 단가 (INSERT 후 별도)
                    if (pricePairs.Count > 0)
                    {
                        var setParts = pricePairs.Select((p, i) => $"`{p.col}`=@pp{i}").ToList();
                        using var upd2 = conn.CreateCommand();
                        upd2.CommandText = $"UPDATE `계약 DB` SET {string.Join(",", setParts)} WHERE C_CompanyName=@n";
                        upd2.Parameters.AddWithValue("@n", companyName);
                        for (int p = 0; p < pricePairs.Count; p++)
                            AddPriceParam(upd2, $"@pp{p}", pricePairs[p].val);
                        upd2.ExecuteNonQuery();
                    }
                }
                ok++;
            }
            catch (Exception ex)
            {
                err++;
            }
        }

        return (ok, err);
    }

    private static void AddBasicParams(DbCommand cmd, string name, string start, string end,
        string days, string amount, string abbr, string ctype, string addr, string rep,
        string ftype, string cat, string prod, string contact, string phone, string email)
    {
        cmd.Parameters.AddWithValue("@n",       name);
        cmd.Parameters.AddWithValue("@start",   ParseDateOrNull(start)  ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@end",     ParseDateOrNull(end)    ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@days",    int.TryParse(days, out int d) ? d : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@amount",  decimal.TryParse(amount.Replace(",",""), out decimal a) ? a : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@abbr",    abbr);
        cmd.Parameters.AddWithValue("@ctype",   ctype);
        cmd.Parameters.AddWithValue("@addr",    addr);
        cmd.Parameters.AddWithValue("@rep",     rep);
        cmd.Parameters.AddWithValue("@ftype",   ftype);
        cmd.Parameters.AddWithValue("@cat",     cat);
        cmd.Parameters.AddWithValue("@prod",    prod);
        cmd.Parameters.AddWithValue("@contact", contact);
        cmd.Parameters.AddWithValue("@phone",   phone);
        cmd.Parameters.AddWithValue("@email",   email);
    }

    private static void AddPriceParam(DbCommand cmd, string paramName, string val)
    {
        var clean = val.Replace(",", "").Trim();
        if (decimal.TryParse(clean, out decimal d))
            cmd.Parameters.AddWithValue(paramName, d);
        else
            cmd.Parameters.AddWithValue(paramName, DBNull.Value);
    }

    private static string? ParseDateOrNull(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        if (DateTime.TryParse(s, out var dt)) return dt.ToString("yyyy-MM-dd");
        // Excel 날짜 시리얼 방어
        if (double.TryParse(s, out double serial) && serial > 1000)
            return DateTime.FromOADate(serial).ToString("yyyy-MM-dd");
        return null;
    }

    // 계약 DB의 기본 컬럼 목록 (단가/수량 컬럼 제외용)
    private static readonly HashSet<string> _basicContractCols = new(StringComparer.OrdinalIgnoreCase)
    {
        "id", "C_CompanyName", "C_ContractStart", "C_ContractEnd", "C_ContractDays",
        "C_ContractAmountVATExcluded", "C_Abbreviation", "C_ContractType", "C_Address",
        "C_Representative", "C_FacilityType", "C_CategoryType", "C_MainProduct",
        "C_ContactPerson", "C_PhoneNumber", "C_Email", "OriginalCompanyName"
    };

    private const string QuantityPrefix = "수량_";

    // ── 업체별 분석단가 조회 ──────────────────────────────────────────────────
    /// <summary>
    /// 계약 DB에서 기본 컬럼을 제외한 단가 컬럼을 모두 읽어
    /// (Analyte, 단가문자열) 리스트로 반환합니다.
    /// </summary>
    public static List<(string Analyte, string Price)> GetContractPrices(string companyName)
    {
        var result = new List<(string Analyte, string Price)>();

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        // 계약 DB의 전체 컬럼 중 기본 컬럼·수량 컬럼 제외 = 단가 컬럼
        var allCols  = DbConnectionFactory.GetColumnNames(conn, "계약 DB");
        var priceCols = allCols
            .Where(c => !_basicContractCols.Contains(c) && !c.StartsWith('_')
                        && !c.StartsWith(QuantityPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (priceCols.Count == 0) return result;

        var colList = string.Join(", ", priceCols.Select(c => $"`{c}`"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT {colList} FROM `계약 DB` WHERE C_CompanyName=@name LIMIT 1";
        cmd.Parameters.AddWithValue("@name", companyName);

        try
        {
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                for (int i = 0; i < priceCols.Count; i++)
                {
                    var v = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
                    result.Add((priceCols[i], v));
                }
            }
            else
            {
                // 행이 없어도 컬럼 목록은 반환 (표시용)
                foreach (var col in priceCols)
                    result.Add((col, ""));
            }
        }
        catch (Exception ex) { }

        return result;
    }

    // ── 업체별 분석단가 저장 ──────────────────────────────────────────────────
    /// <summary>특정 업체의 분석항목별 단가를 일괄 UPDATE합니다.</summary>
    public static bool UpdateContractPrices(string companyName, IEnumerable<(string Analyte, string Price)> prices)
    {
        var priceList = prices.ToList();
        if (priceList.Count == 0) return false;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var existing = DbConnectionFactory.GetColumnNames(conn, "계약 DB");
        var valid = priceList.Where(p => existing.Contains(p.Analyte)).ToList();
        if (valid.Count == 0) return false;

        var sets = string.Join(", ", valid.Select((p, i) => $"`{p.Analyte}`=@p{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"UPDATE `계약 DB` SET {sets} WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", companyName);
        for (int i = 0; i < valid.Count; i++)
        {
            var v = valid[i].Price?.Trim() ?? "";
            if (decimal.TryParse(v.Replace(",", ""), out var d))
                cmd.Parameters.AddWithValue($"@p{i}", d);
            else
                cmd.Parameters.AddWithValue($"@p{i}", DBNull.Value);
        }

        try
        {
            int rows = cmd.ExecuteNonQuery();
            return rows > 0;
        }
        catch (Exception ex) { return false; }
    }

    // ── 업체별 계약수량 컬럼 보장 ─────────────────────────────────────────────
    /// <summary>분석항목 목록에 대해 `수량_{항목명}` INT 컬럼을 보장합니다.</summary>
    public static void EnsureContractQuantityColumns()
    {
        var items = GetAnalysisItems();
        if (items.Count == 0) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        var existing = DbConnectionFactory.GetColumnNames(conn, "계약 DB");

        foreach (var item in items)
        {
            var col = QuantityPrefix + item;
            if (existing.Contains(col)) continue;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"ALTER TABLE `계약 DB` ADD COLUMN IF NOT EXISTS `{col}` INT NULL DEFAULT NULL";
            try { cmd.ExecuteNonQuery(); }
            catch { }
        }
    }

    // ── 업체별 계약수량 조회 ──────────────────────────────────────────────────
    /// <summary>업체의 항목별 계약수량을 반환합니다. 키: 항목명(수량_ 제거), 값: 수량 문자열</summary>
    public static Dictionary<string, string> GetContractQuantities(string companyName)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var allCols = DbConnectionFactory.GetColumnNames(conn, "계약 DB");
        var qtyCols = allCols
            .Where(c => c.StartsWith(QuantityPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (qtyCols.Count == 0) return result;

        var colList = string.Join(", ", qtyCols.Select(c => $"`{c}`"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT {colList} FROM `계약 DB` WHERE C_CompanyName=@name LIMIT 1";
        cmd.Parameters.AddWithValue("@name", companyName);
        try
        {
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                for (int i = 0; i < qtyCols.Count; i++)
                {
                    var analyte = qtyCols[i].Substring(QuantityPrefix.Length);
                    result[analyte] = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";
                }
            }
        }
        catch { }
        return result;
    }

    // ── 업체별 계약수량 저장 ──────────────────────────────────────────────────
    /// <summary>항목별 계약수량을 일괄 UPDATE합니다. 키: 항목명, 값: 수량 문자열</summary>
    public static bool UpdateContractQuantities(string companyName, IEnumerable<(string Analyte, string Qty)> quantities)
    {
        var list = quantities.ToList();
        if (list.Count == 0) return false;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        var existing = DbConnectionFactory.GetColumnNames(conn, "계약 DB");

        var valid = list
            .Where(q => existing.Contains(QuantityPrefix + q.Analyte))
            .ToList();
        if (valid.Count == 0) return false;

        var sets = string.Join(", ", valid.Select((q, i) => $"`{QuantityPrefix}{q.Analyte}`=@q{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"UPDATE `계약 DB` SET {sets} WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", companyName);
        for (int i = 0; i < valid.Count; i++)
        {
            var v = valid[i].Qty?.Trim() ?? "";
            if (int.TryParse(v.Replace(",", ""), out var n))
                cmd.Parameters.AddWithValue($"@q{i}", n);
            else
                cmd.Parameters.AddWithValue($"@q{i}", DBNull.Value);
        }
        try { return cmd.ExecuteNonQuery() > 0; }
        catch { return false; }
    }

    // ── 업체별 처리수량 조회 (계약기간 내) ──────────────────────────────────
    /// <summary>분석의뢰및결과 테이블에서 해당 업체의 항목별 처리수량(O 마크)을 조회합니다. 계약기간 내만 집계.</summary>
    public static Dictionary<string, int> GetProcessedQuantities(string companyName, DateTime? contractStart, DateTime? contractEnd)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(companyName)) return result;

        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 분석정보의 항목 목록 (Analyte)
            var analytes = GetAnalysisItems();
            if (analytes.Count == 0) return result;

            // 각 항목에 대해 의뢰사업장 컬럼에서 "O" 마크 개수 세기
            foreach (var analyte in analytes)
            {
                result[analyte] = 0;
                try
                {
                    // 분석의뢰및결과 테이블에서 해당 업체와 항목의 "O" 마크 개수 조회 (계약기간 필터링)
                    using var cmd = conn.CreateCommand();
                    var whereClause = "의뢰사업장=@company AND `" + analyte + "`='O'";

                    // 계약기간이 지정된 경우 날짜 필터 추가
                    if (contractStart.HasValue)
                        whereClause += " AND 채수일 >= @start";
                    if (contractEnd.HasValue)
                        whereClause += " AND 채수일 <= @end";

                    cmd.CommandText = $"SELECT COUNT(*) FROM `분석의뢰및결과` WHERE {whereClause}";
                    cmd.Parameters.AddWithValue("@company", companyName);
                    if (contractStart.HasValue)
                        cmd.Parameters.AddWithValue("@start", contractStart.Value.ToString("yyyy-MM-dd"));
                    if (contractEnd.HasValue)
                        cmd.Parameters.AddWithValue("@end", contractEnd.Value.ToString("yyyy-MM-dd"));

                    var count = cmd.ExecuteScalar();
                    if (count != null && int.TryParse(count.ToString(), out int c))
                        result[analyte] = c;
                }
                catch { }
            }
        }
        catch { }

        return result;
    }

    // ── 공통 파라미터 ─────────────────────────────────────────────────────────
    private static void SetParams(DbCommand cmd, Contract c)
    {
        cmd.Parameters.AddWithValue("@name",    c.C_CompanyName    ?? "");
        cmd.Parameters.AddWithValue("@start",   c.C_ContractStart.HasValue
                                                    ? c.C_ContractStart.Value.ToString("yyyy-MM-dd")
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@end",     c.C_ContractEnd.HasValue
                                                    ? c.C_ContractEnd.Value.ToString("yyyy-MM-dd")
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@days",    c.C_ContractDays.HasValue
                                                    ? c.C_ContractDays.Value
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@amount",  c.C_ContractAmountVATExcluded.HasValue
                                                    ? c.C_ContractAmountVATExcluded.Value
                                                    : (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@abbr",    c.C_Abbreviation   ?? "");
        cmd.Parameters.AddWithValue("@type",    c.C_ContractType   ?? "");
        cmd.Parameters.AddWithValue("@addr",    c.C_Address        ?? "");
        cmd.Parameters.AddWithValue("@rep",     c.C_Representative ?? "");
        cmd.Parameters.AddWithValue("@ftype",   c.C_FacilityType   ?? "");
        cmd.Parameters.AddWithValue("@ctype",   c.C_CategoryType   ?? "");
        cmd.Parameters.AddWithValue("@prod",    c.C_MainProduct    ?? "");
        cmd.Parameters.AddWithValue("@contact", c.C_ContactPerson  ?? "");
        cmd.Parameters.AddWithValue("@phone",   c.C_PhoneNumber    ?? "");
        cmd.Parameters.AddWithValue("@email",   c.C_Email          ?? "");
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────
    private static string S(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return !r.IsDBNull(i) ? r.GetString(i) ?? "" : ""; }
        catch { return ""; }
    }
    private static int? NullInt(DbDataReader r, string col)
    {
        try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? null : r.GetInt32(i); }
        catch { return null; }
    }
    private static decimal? NullDecimal(DbDataReader r, string col)
    {
        try
        {
            int i = r.GetOrdinal(col);
            if (r.IsDBNull(i)) return null;
            return Convert.ToDecimal(r.GetValue(i));   // MariaDB가 double로 줄 때도 처리
        }
        catch { return null; }
    }
}
