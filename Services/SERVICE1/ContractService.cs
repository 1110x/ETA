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
        EnsureSoftDeleteColumns(conn);
        EnsureKukdoChemical(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                   C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                   C_Address, C_Representative, C_FacilityType, C_CategoryType,
                   C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email,
                   IFNULL(C_ContractType, ''),
                   IFNULL(C_PlaceName, '') AS C_PlaceName,
                   IFNULL(C_IsDeleted, 0) AS C_IsDeleted, C_DeletedAt
            FROM `계약 DB`
            WHERE IFNULL(C_IsDeleted, 0) = 0
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
                C_PlaceName      = S(r, "C_PlaceName"),
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
                C_IsDeleted      = Convert.ToBoolean(r.GetInt32(r.GetOrdinal("C_IsDeleted"))),
            };

            var startStr = S(r, "C_ContractStart");
            if (DateTime.TryParse(startStr, out var start)) c.C_ContractStart = start;

            var endStr = S(r, "C_ContractEnd");
            if (DateTime.TryParse(endStr, out var end)) c.C_ContractEnd = end;

            var deletedAtStr = S(r, "C_DeletedAt");
            if (DateTime.TryParse(deletedAtStr, out var deletedAt)) c.C_DeletedAt = deletedAt;

            list.Add(c);
        }

        return list;
    }

    // ── 휴지통: 삭제된 계약 조회 ──────────────────────────────────────────────
    public static List<Contract> GetDeletedContracts()
    {
        var list = new List<Contract>();
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        EnsureSoftDeleteColumns(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT C_CompanyName, C_ContractStart, C_ContractEnd, C_ContractDays,
                   C_ContractAmountVATExcluded, C_Abbreviation, C_ContractType,
                   C_Address, C_Representative, C_FacilityType, C_CategoryType,
                   C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email,
                   IFNULL(C_ContractType, ''),
                   IFNULL(C_PlaceName, '') AS C_PlaceName,
                   IFNULL(C_IsDeleted, 0) AS C_IsDeleted, C_DeletedAt
            FROM `계약 DB`
            WHERE IFNULL(C_IsDeleted, 0) = 1
            ORDER BY C_DeletedAt DESC";

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var c = new Contract
            {
                C_CompanyName    = S(r, "C_CompanyName"),
                OriginalCompanyName = S(r, "C_CompanyName"),
                C_Abbreviation   = S(r, "C_Abbreviation"),
                C_ContractType   = S(r, "C_ContractType"),
                C_PlaceName      = S(r, "C_PlaceName"),
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
                C_IsDeleted      = Convert.ToBoolean(r.GetInt32(r.GetOrdinal("C_IsDeleted"))),
            };

            var startStr = S(r, "C_ContractStart");
            if (DateTime.TryParse(startStr, out var start)) c.C_ContractStart = start;

            var endStr = S(r, "C_ContractEnd");
            if (DateTime.TryParse(endStr, out var end)) c.C_ContractEnd = end;

            var deletedAtStr = S(r, "C_DeletedAt");
            if (DateTime.TryParse(deletedAtStr, out var deletedAt)) c.C_DeletedAt = deletedAt;

            list.Add(c);
        }
        return list;
    }

    // ── 휴지통: 계약 복구 ──────────────────────────────────────────────────────
    public static bool Restore(string companyName)
    {
        if (string.IsNullOrEmpty(companyName)) return false;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE `계약 DB` SET C_IsDeleted=0, C_DeletedAt=NULL WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", companyName);
        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
    }

    // ── 휴지통: 계약 영구 삭제 ─────────────────────────────────────────────────
    public static bool PermanentDelete(string companyName)
    {
        if (string.IsNullOrEmpty(companyName)) return false;
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM `계약 DB` WHERE C_CompanyName=@name";
        cmd.Parameters.AddWithValue("@name", companyName);
        int rows = cmd.ExecuteNonQuery();
        return rows > 0;
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
                C_PhoneNumber=@phone, C_Email=@email, C_ContractType=@basisContractNo,
                C_PlaceName=@place
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
                 C_MainProduct, C_ContactPerson, C_PhoneNumber, C_Email, C_ContractType, C_PlaceName)
            VALUES
                (@name, @start, @end, @days, @amount, @abbr, @type, @addr,
                 @rep, @ftype, @ctype, @prod, @contact, @phone, @email, @basisContractNo, @place)";

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
        EnsureSoftDeleteColumns(conn);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE `계약 DB` SET C_IsDeleted=1, C_DeletedAt=NOW() WHERE C_CompanyName=@name";
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

        var existing = GetContractColumnNames(conn);

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

        var existingDbCols = GetContractColumnNames(conn).ToHashSet();

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
        "C_ContractAmountVATExcluded", "C_Abbreviation", "C_ContractType", "C_PlaceName", "C_Address",
        "C_Representative", "C_FacilityType", "C_CategoryType", "C_MainProduct",
        "C_ContactPerson", "C_PhoneNumber", "C_Email", "OriginalCompanyName"
    };

    // 계약 DB 컬럼 목록 캐시 (SHOW COLUMNS 중복 호출 방지)
    private static List<string>? _cachedContractCols = null;
    private static readonly object _colCacheLock = new();

    private static List<string> GetContractColumnNames(DbConnection conn)
    {
        if (_cachedContractCols != null) return _cachedContractCols;
        lock (_colCacheLock)
        {
            _cachedContractCols ??= DbConnectionFactory.GetColumnNames(conn, "계약 DB");
            return _cachedContractCols;
        }
    }

    /// <summary>컬럼 캐시 무효화 — ALTER TABLE 후 호출</summary>
    public static void InvalidateColumnCache() => _cachedContractCols = null;

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
        var allCols  = GetContractColumnNames(conn);
        var priceCols = allCols
            .Where(c => !_basicContractCols.Contains(c) && !c.StartsWith('_')
                        && !c.StartsWith(QuantityPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        Log($"[GetContractPrices] 업체='{companyName}', 단가컬럼={priceCols.Count}개, 첫3개={string.Join(",", priceCols.Take(3))}");
        if (priceCols.Count == 0)
        {
            Log($"  → 단가 컬럼 없음");
            return result;
        }

        var colList = string.Join(", ", priceCols.Select(c => $"`{c}`"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT C_CompanyName, {colList} FROM `계약 DB` WHERE C_CompanyName=@name LIMIT 1";
        cmd.Parameters.AddWithValue("@name", companyName);

        Log($"  SQL: {cmd.CommandText}");

        try
        {
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                var dbCompany = r.GetString(0);
                Log($"  → 조회 성공: DB업체='{dbCompany}' (요청='{companyName}')");
                for (int i = 0; i < priceCols.Count; i++)
                {
                    var v = r.IsDBNull(i + 1) ? "" : r.GetValue(i + 1)?.ToString() ?? "";
                    result.Add((priceCols[i], v));
                }
                Log($"  → {result.Count}개 단가 로드됨");
            }
            else
            {
                Log($"  → 행 없음: WHERE C_CompanyName='{companyName}' 매칭 없음");
                // 행이 없어도 컬럼 목록은 반환 (표시용)
                foreach (var col in priceCols)
                    result.Add((col, ""));
            }
        }
        catch (Exception ex)
        {
            Log($"  → 조회 오류: {ex.Message}\n{ex.StackTrace}");
        }

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

        var existing = GetContractColumnNames(conn);
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

    // ── 단가 복사 (한 계약에서 다른 계약으로) ───────────────────────────────
    /// <summary>fromCompany의 분석단가를 toCompany에 복사합니다.</summary>
    public static bool CopyContractPrices(string fromCompany, string toCompany)
    {
        if (string.IsNullOrWhiteSpace(fromCompany) || string.IsNullOrWhiteSpace(toCompany))
        {
            Log($"[CopyContractPrices] 입력값 오류: from='{fromCompany}', to='{toCompany}'");
            return false;
        }

        try
        {
            Log($"[CopyContractPrices] 시작: {fromCompany} → {toCompany}");

            // 원본 계약의 단가 조회
            var srcPrices = GetContractPrices(fromCompany);
            if (srcPrices.Count == 0)
            {
                Log($"  → 복사할 단가 없음 (source count=0)");
                return false;
            }

            Log($"  → 원본에서 {srcPrices.Count}개 단가 로드됨");

            // 대상 계약에 단가 업데이트
            bool updated = UpdateContractPrices(toCompany, srcPrices);
            if (updated)
            {
                Log($"  ✅ 완료: {srcPrices.Count}개 단가 복사 성공");
            }
            else
            {
                Log($"  ⚠️ 부분 성공: UpdateContractPrices가 0개 행 업데이트");
            }
            return updated;
        }
        catch (Exception ex)
        {
            Log($"  ❌ 오류: {ex.Message}\n{ex.StackTrace}");
            return false;
        }
    }

    // ── 업체별 계약수량 컬럼 보장 ─────────────────────────────────────────────
    /// <summary>분석항목 목록에 대해 `수량_{항목명}` INT 컬럼을 보장합니다.</summary>
    public static void EnsureContractQuantityColumns()
    {
        var items = GetAnalysisItems();
        if (items.Count == 0) return;

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();
        var existing = GetContractColumnNames(conn);

        bool altered = false;
        foreach (var item in items)
        {
            var col = QuantityPrefix + item;
            if (existing.Contains(col)) continue;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"ALTER TABLE `계약 DB` ADD COLUMN IF NOT EXISTS `{col}` INT NULL DEFAULT NULL";
            try { cmd.ExecuteNonQuery(); altered = true; }
            catch { }
        }
        if (altered) InvalidateColumnCache();
    }

    // ── 업체별 계약수량 조회 ──────────────────────────────────────────────────
    /// <summary>업체의 항목별 계약수량을 반환합니다. 키: 항목명(수량_ 제거), 값: 수량 문자열</summary>
    public static Dictionary<string, string> GetContractQuantities(string companyName)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        var allCols = GetContractColumnNames(conn);
        var qtyCols = allCols
            .Where(c => c.StartsWith(QuantityPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        Log($"[GetContractQuantities] 업체={companyName}, 수량컬럼={qtyCols.Count}개");
        if (qtyCols.Count == 0)
        {
            Log($"  → 수량 컬럼 없음");
            return result;
        }

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
                Log($"  → DB에서 {result.Count}개 수량 로드됨");
            }
            else
            {
                Log($"  → 행 없음");
            }
        }
        catch (Exception ex)
        {
            Log($"  → 조회 오류: {ex.Message}");
        }
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
        var existing = GetContractColumnNames(conn);

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
    /// <summary>수질분석센터_결과 테이블에서 해당 업체의 항목별 처리수량(O 마크)을 조회합니다. 계약기간 내만 집계.</summary>
    public static Dictionary<string, int> GetProcessedQuantities(string companyName, DateTime? contractStart, DateTime? contractEnd)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(companyName)) return result;
        try
        {
            var analytes = GetAnalysisItems();
            if (analytes.Count == 0) return result;

            foreach (var a in analytes) result[a] = 0;

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            // 67번 개별 쿼리 → 단일 쿼리로 통합
            var sums  = string.Join(", ", analytes.Select(a => $"SUM(`{a}`='O')"));
            var where = "의뢰사업장=@company";
            if (contractStart.HasValue) where += " AND 채수일 >= @start";
            if (contractEnd.HasValue)   where += " AND 채수일 <= @end";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT {sums} FROM `수질분석센터_결과` WHERE {where}";
            cmd.Parameters.AddWithValue("@company", companyName);
            if (contractStart.HasValue)
                cmd.Parameters.AddWithValue("@start", contractStart.Value.ToString("yyyy-MM-dd"));
            if (contractEnd.HasValue)
                cmd.Parameters.AddWithValue("@end", contractEnd.Value.ToString("yyyy-MM-dd"));

            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                for (int i = 0; i < analytes.Count; i++)
                {
                    if (!r.IsDBNull(i) && int.TryParse(r.GetValue(i)?.ToString(), out int c))
                        result[analytes[i]] = c;
                }
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
        cmd.Parameters.AddWithValue("@basisContractNo", c.C_ContractType ?? "");
        cmd.Parameters.AddWithValue("@place",   c.C_PlaceName      ?? "");
    }

    // ── 소프트 삭제 컬럼 마이그레이션 ──────────────────────────────────────────
    private static void EnsureSoftDeleteColumns(DbConnection conn)
    {
        try
        {
            if (!DbConnectionFactory.ColumnExists(conn, "계약 DB", "C_IsDeleted"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `계약 DB` ADD COLUMN `C_IsDeleted` TINYINT DEFAULT 0";
                cmd.ExecuteNonQuery();
            }
            if (!DbConnectionFactory.ColumnExists(conn, "계약 DB", "C_DeletedAt"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `계약 DB` ADD COLUMN `C_DeletedAt` DATETIME NULL";
                cmd.ExecuteNonQuery();
            }
            if (!DbConnectionFactory.ColumnExists(conn, "계약 DB", "C_ContractType"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `계약 DB` ADD COLUMN `C_ContractType` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
            if (!DbConnectionFactory.ColumnExists(conn, "계약 DB", "C_PlaceName"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "ALTER TABLE `계약 DB` ADD COLUMN `C_PlaceName` TEXT DEFAULT ''";
                cmd.ExecuteNonQuery();
            }
        }
        catch { }
    }

    // ── 국도화학 자동 추가 (초기화) ───────────────────────────────────────────
    private static void EnsureKukdoChemical(DbConnection conn)
    {
        try
        {
            using var checkCmd = conn.CreateCommand();
            checkCmd.CommandText = "SELECT COUNT(*) FROM `계약 DB` WHERE C_CompanyName='국도화학㈜익산지점'";
            var count = (long)checkCmd.ExecuteScalar();
            if (count == 0)
            {
                using var insertCmd = conn.CreateCommand();
                insertCmd.CommandText = @"
                    INSERT INTO `계약 DB`
                        (C_CompanyName, C_Abbreviation, C_Address, C_Representative,
                         C_FacilityType, C_CategoryType, C_MainProduct, C_ContractType)
                    VALUES
                        ('국도화학㈜익산지점', '국도화학', '전북특별자치도 익산시 석암로13길 32(석암동)',
                         '대표자', '제조업', '제조업', '에폭시수지, 골판지 제조업', '1000479061')";
                insertCmd.ExecuteNonQuery();
            }
        }
        catch { }
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

    private static void Log(string msg)
    {
        var logPath = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Contract.log"));
        try
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
            if (App.EnableLogging)
                File.AppendAllText(logPath, line + Environment.NewLine);
        }
        catch { }
    }
}
