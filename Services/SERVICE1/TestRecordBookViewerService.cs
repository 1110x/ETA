using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

/// <summary>
/// *_시험기록부 테이블 뷰어 전용 조회 서비스.
/// 테이블 목록 / 분석일 목록 / 특정 날짜 행 조회.
/// </summary>
public static class TestRecordBookViewerService
{
    /// <summary>DB 에 존재하는 모든 *_시험기록부 테이블 이름 (오름차순).</summary>
    public static List<string> GetAllTables()
    {
        var list = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME LIKE '%\_시험기록부' ESCAPE '\\'
                ORDER BY TABLE_NAME";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                if (!r.IsDBNull(0)) list.Add(r.GetString(0));
        }
        catch { }
        return list;
    }

    /// <summary>특정 테이블의 DISTINCT 분석일 목록 (최근순). 분석일 컬럼 없거나 실패 시 빈 목록.</summary>
    public static List<string> GetDates(string tableName)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(tableName)) return list;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            var cols = DbConnectionFactory.GetColumnNames(conn, tableName);
            string dateCol = cols.FirstOrDefault(c => c.Equals("분석일", StringComparison.OrdinalIgnoreCase))
                          ?? cols.FirstOrDefault(c => c.Equals("채수일", StringComparison.OrdinalIgnoreCase))
                          ?? cols.FirstOrDefault(c => c.Equals("채취일자", StringComparison.OrdinalIgnoreCase))
                          ?? "";
            if (string.IsNullOrEmpty(dateCol)) return list;

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT DISTINCT `{dateCol}` FROM `{tableName}` WHERE `{dateCol}` IS NOT NULL AND `{dateCol}` <> '' ORDER BY `{dateCol}` DESC";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                if (!r.IsDBNull(0)) list.Add(r.GetString(0));
        }
        catch { }
        return list;
    }

    /// <summary>해당 테이블/날짜의 모든 행. 반환: (컬럼목록, 행 리스트(컬럼명→값))</summary>
    public static (List<string> Columns, List<Dictionary<string, string>> Rows) GetRowsByDate(
        string tableName, string date)
    {
        var cols = new List<string>();
        var rows = new List<Dictionary<string, string>>();
        if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(date))
            return (cols, rows);
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            var allCols = DbConnectionFactory.GetColumnNames(conn, tableName);
            string dateCol = allCols.FirstOrDefault(c => c.Equals("분석일", StringComparison.OrdinalIgnoreCase))
                          ?? allCols.FirstOrDefault(c => c.Equals("채수일", StringComparison.OrdinalIgnoreCase))
                          ?? allCols.FirstOrDefault(c => c.Equals("채취일자", StringComparison.OrdinalIgnoreCase))
                          ?? "";
            if (string.IsNullOrEmpty(dateCol)) return (cols, rows);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT * FROM `{tableName}` WHERE `{dateCol}` = @d";
            cmd.Parameters.AddWithValue("@d", date);

            using var r = cmd.ExecuteReader();
            for (int i = 0; i < r.FieldCount; i++)
                cols.Add(r.GetName(i));
            while (r.Read())
            {
                var dict = new Dictionary<string, string>();
                for (int i = 0; i < r.FieldCount; i++)
                {
                    dict[cols[i]] = r.IsDBNull(i) ? "" : (r.GetValue(i)?.ToString() ?? "");
                }
                rows.Add(dict);
            }
        }
        catch { }
        return (cols, rows);
    }

    /// <summary>테이블 표시명(언더스코어 → 공백, "_시험기록부" 접미 제거).</summary>
    public static string PrettyName(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) return tableName;
        var name = tableName.EndsWith("_시험기록부") ? tableName[..^"_시험기록부".Length] : tableName;
        return name.Replace('_', ' ');
    }

    /// <summary>
    /// 테이블명을 정규화해 분석정보.Analyte 로 매핑.
    /// 1) Analyte 정규화 == 테이블 정규화 → exact
    /// 2) AliasX 정규화 == 테이블 정규화
    /// 3) 가장 긴 Analyte 정규화가 테이블 정규화의 prefix → 부모 Analyte 반환
    ///    (예: 총_유기탄소_NPOC → "총 유기탄소", 페놀류_직접법 → "페놀류")
    /// </summary>
    public static string GetAnalyteName(string tableName)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "분석정보")) return "";
            bool hasAlias = DbConnectionFactory.ColumnExists(conn, "분석정보", "AliasX");
            string aliasExpr = hasAlias ? "COALESCE(`AliasX`,'')" : "''";
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `Analyte`, {aliasExpr} FROM `분석정보`";
            var key = NormalizeForTable(PrettyName(tableName));

            var pairs = new List<(string analyte, string normA, string normAlias)>();
            using (var r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    var a = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                    if (string.IsNullOrEmpty(a)) continue;
                    var alias = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                    pairs.Add((a, NormalizeForTable(a), NormalizeForTable(alias)));
                }
            }
            // 1) exact Analyte
            foreach (var p in pairs)
                if (p.normA == key) return p.analyte;
            // 2) exact AliasX
            foreach (var p in pairs)
                if (!string.IsNullOrEmpty(p.normAlias) && p.normAlias == key) return p.analyte;
            // 3) longest-prefix Analyte
            string best = ""; int bestLen = 0;
            foreach (var p in pairs)
            {
                if (string.IsNullOrEmpty(p.normA)) continue;
                if (key.StartsWith(p.normA, StringComparison.OrdinalIgnoreCase)
                    && p.normA.Length > bestLen)
                { best = p.analyte; bestLen = p.normA.Length; }
            }
            return best;
        }
        catch { }
        return "";
    }

    /// <summary>분석정보.DecimalPlaces (없으면 1).</summary>
    public static int GetDecimalPlaces(string analyte)
    {
        if (string.IsNullOrWhiteSpace(analyte)) return 1;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COALESCE(DecimalPlaces, 1) FROM `분석정보` WHERE `Analyte`=@a LIMIT 1";
            cmd.Parameters.AddWithValue("@a", analyte);
            var v = cmd.ExecuteScalar();
            if (v != null && int.TryParse(v.ToString(), out var n)) return n;
        }
        catch { }
        return 1;
    }

    /// <summary>숫자 문자열을 분석정보 DecimalPlaces 로 포맷.</summary>
    public static string FormatWithDecimals(string value, int decimalPlaces)
    {
        if (string.IsNullOrWhiteSpace(value)) return value;
        if (!double.TryParse(value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return value;
        return v.ToString($"F{Math.Max(0, decimalPlaces)}",
            System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>결과를 LoQ 와 비교하여 미만이면 "ND" 반환, 아니면 DecimalPlaces 포맷 적용.</summary>
    public static string FormatResultWithLoQ(string value, int decimalPlaces, double? loq)
    {
        if (string.IsNullOrWhiteSpace(value)) return value;
        if (!double.TryParse(value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return value;
        if (loq.HasValue && v < loq.Value) return "ND";
        return v.ToString($"F{Math.Max(0, decimalPlaces)}",
            System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 분석정보의 농도공식 토큰을 행 값으로 치환해 사람이 읽을 수 있는 문자열 생성.
    /// 예: "( ( 흡광도 - 절편 ) / 기울기 ) × 시료량" → "( ( 0.2001 - 0.0024 ) / 0.9187 ) × 25"
    /// </summary>
    public static string SubstituteFormula(string formula, Dictionary<string, string> row)
    {
        if (string.IsNullOrWhiteSpace(formula)) return "";
        var stripTrailing = new[] { "=", "농도", "✓", "✗" };
        var tokens = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var sb = new System.Text.StringBuilder();
        bool first = true;
        foreach (var t in tokens)
        {
            if (System.Array.IndexOf(stripTrailing, t) >= 0) continue;
            string outTok;
            // 변수 치환: 행에서 같은 이름의 컬럼 값
            var v = "";
            foreach (var kv in row)
                if (string.Equals(kv.Key, t, System.StringComparison.OrdinalIgnoreCase))
                { v = kv.Value ?? ""; break; }
            outTok = !string.IsNullOrWhiteSpace(v) ? v : t;
            if (!first) sb.Append(' ');
            sb.Append(outTok);
            first = false;
        }
        return sb.ToString().Trim();
    }

    /// <summary>
    /// 각 *_시험기록부 테이블을 분석정보 테이블의 Analyte 와 매칭해
    /// (Category, ES, 약칭) 메타 정보를 구한다. 매칭 실패는 ("기타", "", "") 로.
    /// </summary>
    public static Dictionary<string, (string Category, string ES, string 약칭)> GetAnalyteMetaForTables(
        IEnumerable<string> tableNames)
    {
        var result = new Dictionary<string, (string, string, string)>(StringComparer.OrdinalIgnoreCase);
        var tables = tableNames.ToList();
        foreach (var t in tables) result[t] = ("기타", "", "");

        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "분석정보")) return result;
            bool has약칭 = DbConnectionFactory.ColumnExists(conn, "분석정보", "약칭");
            bool hasAlias = DbConnectionFactory.ColumnExists(conn, "분석정보", "AliasX");
            string abbrExpr = has약칭 ? "COALESCE(`약칭`,'')" : "''";
            string aliasExpr = hasAlias ? "COALESCE(`AliasX`,'')" : "''";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT `Analyte`, COALESCE(`Category`,''), COALESCE(`ES`,''), {abbrExpr}, {aliasExpr} FROM `분석정보`";

            // 한 Analyte 가 여러 행으로 나올 수 있으므로 리스트로 보존
            var pairs = new List<(string normA, string normAlias, string Category, string ES, string 약칭)>();
            using (var r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    var a     = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                    if (string.IsNullOrWhiteSpace(a)) continue;
                    var c     = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                    var es    = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                    var abbr  = r.IsDBNull(3) ? "" : r.GetString(3).Trim();
                    var alias = r.IsDBNull(4) ? "" : r.GetString(4).Trim();
                    pairs.Add((NormalizeForTable(a), NormalizeForTable(alias), c, es, abbr));
                }
            }

            (string C, string E, string A)? FindMeta(string key)
            {
                // 1) Analyte exact
                foreach (var p in pairs)
                    if (p.normA == key) return (p.Category, p.ES, p.약칭);
                // 2) AliasX exact
                foreach (var p in pairs)
                    if (!string.IsNullOrEmpty(p.normAlias) && p.normAlias == key) return (p.Category, p.ES, p.약칭);
                // 3) 가장 긴 Analyte 가 prefix
                (string C, string E, string A)? best = null; int bestLen = 0;
                foreach (var p in pairs)
                {
                    if (string.IsNullOrEmpty(p.normA)) continue;
                    if (key.StartsWith(p.normA, StringComparison.OrdinalIgnoreCase)
                        && p.normA.Length > bestLen)
                    { best = (p.Category, p.ES, p.약칭); bestLen = p.normA.Length; }
                }
                return best;
            }

            foreach (var t in tables)
            {
                var key = NormalizeForTable(PrettyName(t));
                var meta = FindMeta(key);
                if (meta.HasValue)
                    result[t] = (string.IsNullOrWhiteSpace(meta.Value.C) ? "기타" : meta.Value.C,
                                 meta.Value.E, meta.Value.A);
            }
        }
        catch { }
        return result;
    }

    /// <summary>
    /// 저장된 *_시험기록부 행을 분석결과입력 파서뷰와 동일 구조의 Model 로 변환.
    /// </summary>
    public static ETA.Views.Pages.PAGE1.TestRecordBookParsedView.Model BuildParsedModel(
        string table, string date, string categoryKey, string? companyFilter = null)
    {
        var m = new ETA.Views.Pages.PAGE1.TestRecordBookParsedView.Model
        {
            AnalysisDate = date,
            CategoryKey  = categoryKey ?? "",
            FileLabel    = "(저장됨)",
            TargetTable  = table,
        };
        var (cols, rows) = GetRowsByDate(table, date);
        if (!string.IsNullOrWhiteSpace(companyFilter))
        {
            rows = rows.Where(r => r.TryGetValue("업체명", out var v)
                && string.Equals((v ?? "").Trim(), companyFilter, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
        if (rows.Count == 0) return m;
        var first = rows[0];

        // 분석방법 / 비고
        m.AnalysisMethod = FirstNonEmpty(first, "분석방법");
        m.Memo           = FirstNonEmpty(first, "비고", "분석조건");

        // 분석정보 메타: ES / Method / instrument 조회
        try
        {
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            if (DbConnectionFactory.TableExists(conn, "분석정보"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"SELECT COALESCE(ES,''), COALESCE(Method,''), COALESCE(instrument,'')
                                    FROM `분석정보` WHERE `Analyte`=@a LIMIT 1";
                cmd.Parameters.AddWithValue("@a", string.IsNullOrEmpty(GetAnalyteName(table)) ? table : GetAnalyteName(table));
                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    m.ES         = r.IsDBNull(0) ? "" : r.GetString(0).Trim();
                    m.Method     = r.IsDBNull(1) ? "" : r.GetString(1).Trim();
                    m.Instrument = r.IsDBNull(2) ? "" : r.GetString(2).Trim();
                }
            }
        }
        catch { }

        // ── 검정곡선 통합 파싱 ──
        // 컬럼명을 두 축으로 분류:
        //   1) Set suffix: "_TC", "_IC" 같은 분리 검정곡선 식별자 (없으면 default)
        //   2) Role suffix: _농도/_mgL → 농도, _값/_AU/_abs/흡광 → 응답, _ISTD → 내부표준
        // 새 양식이 추가돼도 setSuffixes / role 룰만 확장하면 됨.
        var setSuffixes = new[] { "_TC", "_IC" };
        var re = new System.Text.RegularExpressions.Regex(
            @"^(?:ST\w*|표준)[-_ ]?(\d+)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // setKey ("" | "TC" | "IC" 등) → (Keys, Conc, Resp, Istd)
        var sets = new Dictionary<string, (List<int> Keys,
                                            Dictionary<int,string> Conc,
                                            Dictionary<int,string> Resp,
                                            Dictionary<int,string> Istd)>(StringComparer.OrdinalIgnoreCase);
        (List<int>, Dictionary<int,string>, Dictionary<int,string>, Dictionary<int,string>) Acquire(string key)
        {
            if (!sets.TryGetValue(key, out var s))
            {
                s = (new List<int>(), new Dictionary<int,string>(), new Dictionary<int,string>(), new Dictionary<int,string>());
                sets[key] = s;
            }
            return s;
        }

        foreach (var c in cols)
        {
            var match = re.Match(c);
            if (!match.Success) continue;
            if (!int.TryParse(match.Groups[1].Value, out var n)) continue;

            // 1) set suffix 추출
            string setKey = "";
            string body   = c;
            foreach (var sfx in setSuffixes)
            {
                if (body.EndsWith(sfx, System.StringComparison.OrdinalIgnoreCase))
                {
                    setKey = sfx.TrimStart('_');
                    body   = body.Substring(0, body.Length - sfx.Length);
                    break;
                }
            }

            // 2) role 분류
            bool isIstd = body.EndsWith("_ISTD", System.StringComparison.OrdinalIgnoreCase);
            bool isConc = body.EndsWith("_농도",  System.StringComparison.OrdinalIgnoreCase)
                       || body.EndsWith("_mgL",  System.StringComparison.OrdinalIgnoreCase)
                       || body.EndsWith("_conc", System.StringComparison.OrdinalIgnoreCase);
            bool isAbs  = body.EndsWith("_값",   System.StringComparison.OrdinalIgnoreCase)
                       || body.EndsWith("_AU",   System.StringComparison.OrdinalIgnoreCase)
                       || body.EndsWith("_abs",  System.StringComparison.OrdinalIgnoreCase)
                       || body.Contains("흡광",   System.StringComparison.OrdinalIgnoreCase)
                       || body.Contains("Abs",   System.StringComparison.OrdinalIgnoreCase);

            var bucket = Acquire(setKey);
            if (!bucket.Item1.Contains(n)) bucket.Item1.Add(n);
            if      (isIstd) bucket.Item4[n] = c;
            else if (isConc) bucket.Item2[n] = c;
            else if (isAbs)  bucket.Item3[n] = c;
            else if (!bucket.Item2.ContainsKey(n)) bucket.Item2[n] = c;
        }
        foreach (var k in sets.Keys.ToList()) sets[k].Keys.Sort();

        // 검정계수 (set 별 — 기울기/절편/R2[_TC|_IC])
        (string a, string b, string r2) ReadCoef(string sufx)
        {
            string ax = "", bx = "", rx = "";
            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(ax)) ax = FirstNonEmpty(row, "검량선_a"+sufx, "기울기"+sufx, "slope"+sufx, "Slope"+sufx);
                if (string.IsNullOrWhiteSpace(bx)) bx = FirstNonEmpty(row, "검량선_b"+sufx, "절편"+sufx, "intercept"+sufx, "Intercept"+sufx);
                if (string.IsNullOrWhiteSpace(rx)) rx = FirstNonEmpty(row, "R2"+sufx, "R²"+sufx, "검량선_R2"+sufx, "R_squared"+sufx, "상관계수"+sufx, "R값"+sufx);
            }
            return (ax, bx, rx);
        }

        // 모델 매핑 — 다중 set (TCIC) 또는 단일 set
        bool hasTcSet = sets.ContainsKey("TC") && sets["TC"].Keys.Count > 0;
        bool hasIcSet = sets.ContainsKey("IC") && sets["IC"].Keys.Count > 0;

        void FillModelSet(string key, List<int> outKeys, List<string> outConc, List<string> outResp,
                          out string slopeText, out string r2Text, List<string>? outIstd = null)
        {
            slopeText = ""; r2Text = "";
            if (!sets.TryGetValue(key, out var s)) return;

            bool any = s.Keys.Any(n =>
                (s.Conc.TryGetValue(n, out var cc) && !string.IsNullOrWhiteSpace(FirstValue(first, cc))) ||
                (s.Resp.TryGetValue(n, out var ac) && !string.IsNullOrWhiteSpace(FirstValue(first, ac))));
            if (!any) return;

            foreach (var n in s.Keys)
            {
                outKeys.Add(n);
                outConc.Add(s.Conc.TryGetValue(n, out var cc) ? FirstValue(first, cc) : "");
                outResp.Add(s.Resp.TryGetValue(n, out var ac) ? FirstValue(first, ac) : "");
                if (outIstd != null && s.Istd.Count > 0)
                    outIstd.Add(s.Istd.TryGetValue(n, out var ic) ? FirstValue(first, ic) : "");
            }

            string sufx = string.IsNullOrEmpty(key) ? "" : "_" + key;
            var (a, b, r2) = ReadCoef(sufx);
            if (!string.IsNullOrWhiteSpace(a) || !string.IsNullOrWhiteSpace(b))
                slopeText = $"a={ToDecimalStr(a)}  b={ToDecimalStr(b)}";
            if (!string.IsNullOrWhiteSpace(r2)) r2Text = $"R²={ToDecimalStr(r2)}";
        }

        if (hasTcSet && hasIcSet)
        {
            m.IsTcic = true;
            string tcSlope, tcR2, icSlope, icR2;
            FillModelSet("TC", m.TcStandardKeys, m.TcStandardConc, m.TcStandardAbs, out tcSlope, out tcR2);
            FillModelSet("IC", m.IcStandardKeys, m.IcStandardConc, m.IcStandardAbs, out icSlope, out icR2);
            m.TcSlopeText = tcSlope; m.TcR2Text = tcR2;
            m.IcSlopeText = icSlope; m.IcR2Text = icR2;
        }
        else
        {
            string slope, r2;
            FillModelSet("", m.StandardKeys, m.StandardConc, m.StandardAbs, out slope, out r2, m.StandardIstd);
            m.SlopeText = slope; m.R2Text = r2;
        }

        // (구 hasAnyCalVal/단일 set 처리는 위 통합 알고리즘으로 대체됨)

        // ── BOD 전용 — 식종 정보 / GGA 정도관리 (검정곡선 자리) ──
        bool isBod = table.Contains("산소요구량") || table.Contains("BOD",
            System.StringComparison.OrdinalIgnoreCase);
        if (isBod)
        {
            m.IsBod = true;
            m.SeedHeaders = new List<string> { "구분", "시료량(V)", "D1", "D2", "Result(mg/L)", "비고" };

            string seedV   = FirstValue(first, "식종시료량");
            string seedD1  = FirstValue(first, "15min_DO");
            string seedD2  = FirstValue(first, "5Day_DO");
            string seedBod = FirstValue(first, "식종BOD");
            string seedPct = FirstValue(first, "식종함유량");
            if (!string.IsNullOrWhiteSpace(seedV) || !string.IsNullOrWhiteSpace(seedD1))
                m.SeedRows.Add(new List<string> { "식종수의 BOD", seedV, seedD1, seedD2,
                    seedBod, string.IsNullOrWhiteSpace(seedPct) ? "" : $"y%={seedPct}" });

            string scfV   = FirstValue(first, "SCF_시료량");
            string scfD1  = FirstValue(first, "SCF_D1");
            string scfD2  = FirstValue(first, "SCF_D2");
            string scfRes = FirstValue(first, "SCF_Result");
            if (!string.IsNullOrWhiteSpace(scfV) || !string.IsNullOrWhiteSpace(scfD1))
                m.SeedRows.Add(new List<string> { "SCF (식종희석수)", scfV, scfD1, scfD2, scfRes, "" });

            for (int i = 1; i <= 3; i++)
            {
                string ggaV   = FirstValue(first, $"GGA{i}_V");
                string ggaD1  = FirstValue(first, $"GGA{i}_D1");
                string ggaD2  = FirstValue(first, $"GGA{i}_D2");
                string ggaBod = FirstValue(first, $"GGA{i}_BOD");
                if (!string.IsNullOrWhiteSpace(ggaV) || !string.IsNullOrWhiteSpace(ggaBod))
                    m.SeedRows.Add(new List<string> { $"GGA {i}", ggaV, ggaD1, ggaD2, ggaBod, "" });
            }
        }

        // 시료 그리드 — 우선순위 컬럼만, 순서 고정
        // TOC NPOC 등 흡광도 대신 AU 컬럼을 쓰는 테이블은 시료량을 숨기고 AU 표시
        bool hasAU  = cols.Any(c => c.Equals("AU",   System.StringComparison.OrdinalIgnoreCase));
        bool hasAbs = cols.Any(c => c.Equals("흡광도", System.StringComparison.OrdinalIgnoreCase));
        bool hasTcic = cols.Any(c => c.Equals("TCAU", System.StringComparison.OrdinalIgnoreCase))
                    && cols.Any(c => c.Equals("ICAU", System.StringComparison.OrdinalIgnoreCase));
        // 총대장균군(평판집락법): A/B 두 평판 colony 수 + 희석배수
        bool isColiform = cols.Any(c => c.Equals("A", System.StringComparison.OrdinalIgnoreCase))
                       && cols.Any(c => c.Equals("B", System.StringComparison.OrdinalIgnoreCase))
                       && cols.Any(c => c.Equals("사용희석배수", System.StringComparison.OrdinalIgnoreCase));
        // SS 중량법 (NHexan/SS): 전무게/후무게/무게차
        bool isGravimetric = cols.Any(c => c.Equals("전무게", System.StringComparison.OrdinalIgnoreCase))
                          && cols.Any(c => c.Equals("후무게", System.StringComparison.OrdinalIgnoreCase));
        // VOC/GC: 행단위 ISTD 컬럼 존재 (시료별 내부표준 응답)
        bool hasIstdRow = cols.Any(c => c.Equals("ISTD", System.StringComparison.OrdinalIgnoreCase));
        string[] priority = isBod
            ? new[] { "SN", "시료명", "시료구분", "시료량", "D1", "D2", "F_xy", "희석배수", "결과", "계산식" }
            : isColiform
                ? new[] { "SN", "시료명", "시료구분", "시료량", "A", "B", "희석배수", "결과", "계산식" }
                : isGravimetric
                    ? new[] { "SN", "시료명", "시료구분", "시료량", "전무게", "후무게", "무게차", "희석배수", "결과", "계산식" }
                    : hasTcic
                        ? new[] { "SN", "시료명", "시료구분", "TCAU", "TCcon", "ICAU", "ICcon", "농도", "희석배수", "결과", "계산식" }
                        : hasIstdRow
                            ? new[] { "SN", "시료명", "시료구분", "흡광도", "ISTD", "농도", "희석배수", "결과", "계산식" }
                            : (hasAU && !hasAbs)
                                ? new[] { "SN", "시료명", "시료구분", "AU", "계산농도", "농도", "희석배수", "결과값", "결과", "계산식" }
                                : new[] { "SN", "시료명", "시료구분", "시료량", "흡광도", "계산농도", "농도", "희석배수", "결과값", "결과", "계산식" };
        var hideExact = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase)
        {
            "id","rowid","_id","Id","분석일","등록일시","분석방법","비고","업체명",
            "구분","소스구분","분석자","분석자1","분석자2","측정자","작성자","의뢰명",
            "기울기","절편","slope","intercept","검량선_a","검량선_b",
            "R2","R²","검량선_R2","R_squared","상관계수",
            "원시료희석배수","원시료_희석배수",
            // BOD 전용 — Seed 섹션에서 별도 표시되므로 시료 그리드에서 숨김
            "식종시료량","15min_DO","5Day_DO","식종BOD","식종함유량","희석수시료량",
            "SCF_시료량","SCF_D1","SCF_D2","SCF_Result","식종D1","식종D2","식종여부",
            "GGA1_V","GGA1_D1","GGA1_D2","GGA1_BOD",
            "GGA2_V","GGA2_D1","GGA2_D2","GGA2_BOD",
            "GGA3_V","GGA3_D1","GGA3_D2","GGA3_BOD",
        };
        // AU 컬럼 사용 테이블(TOC NPOC 등)은 시료량 컬럼 숨김 — 의미 없는 컬럼
        if (hasAU && !hasAbs) hideExact.Add("시료량");
        bool IsCalCol(string c) => re.IsMatch(c);

        m.SampleHeaders = new List<string>();
        foreach (var p in priority)
            if (cols.Any(c => string.Equals(c, p, System.StringComparison.OrdinalIgnoreCase))
                && !m.SampleHeaders.Contains(p))
                m.SampleHeaders.Add(p);

        // 결과 우측에 계산식 컬럼이 항상 표시되도록 강제 추가 (DB 컬럼 없어도 동적 생성)
        if (!m.SampleHeaders.Contains("계산식"))
            m.SampleHeaders.Add("계산식");

        // 분석정보의 농도 계산공식 (항목별 1개) — 토큰 치환 + 재계산용
        string analyte = GetAnalyteName(table);
        string formula = string.IsNullOrWhiteSpace(analyte) ? "" :
            ETA.Services.SERVICE3.AnalysisNoteService.GetFormula(analyte);
        int decimalPlaces = GetDecimalPlaces(analyte);
        double? loq = AnalysisService.GetLoQ(analyte);

        // 모델에 보존 — 결과표시(출력) 등에서 시료별 식 대신 일반 수식 사용
        m.Analyte       = analyte ?? "";
        m.ResultFormula = formula ?? "";

        // 설정 → 분석조건 (오븐온도/유량 등 Key/Value) 로드 — 분석정보 아래 표시용
        try
        {
            var conds = ETA.Services.SERVICE1.AnalysisConditionService.Load(analyte);
            foreach (var c in conds)
            {
                if (!string.IsNullOrWhiteSpace(c.Key))
                    m.AnalysisConditions.Add((c.Key, c.Value ?? ""));
            }
        }
        catch { /* 분석조건 부재는 치명 X */ }

        // 처리시설_마스터.id 기반 정렬 매핑 (시료구분=시설명, 시료명)
        var orderMap = LoadFacilityOrderMap();

        var sortable = new List<(int order, List<string> values, string sampleClass)>();
        foreach (var row in rows)
        {
            var values = new List<string>();
            foreach (var h in m.SampleHeaders)
            {
                if (h == "계산식")
                {
                    var stored = FirstValue(row, "계산식");
                    if (!string.IsNullOrWhiteSpace(stored)) { values.Add(stored); continue; }

                    // BOD 전용 — ES 04305.1c 공식
                    if (isBod)
                    {
                        string d1 = FirstValue(row, "D1");
                        string d2 = FirstValue(row, "D2");
                        string fxy = FirstValue(row, "F_xy");
                        string p   = FirstValue(row, "희석배수");
                        string b1  = FirstValue(row, "15min_DO");
                        string b2  = FirstValue(row, "5Day_DO");
                        bool hasSeed = double.TryParse(fxy, out var f) && f > 0;
                        string txt;
                        double? calc = null;
                        if (double.TryParse(d1, out var vd1) && double.TryParse(d2, out var vd2)
                            && double.TryParse(p,  out var vp))
                        {
                            if (hasSeed && double.TryParse(b1, out var vb1) && double.TryParse(b2, out var vb2))
                            {
                                txt = $"( ( {d1} - {d2} ) - ( {b1} - {b2} ) × {fxy} ) × {p}";
                                calc = ((vd1 - vd2) - (vb1 - vb2) * f) * vp;
                            }
                            else
                            {
                                txt = $"( {d1} - {d2} ) × {p}";
                                calc = (vd1 - vd2) * vp;
                            }
                            values.Add(calc.HasValue
                                ? $"{txt} = {FormatWithDecimals(calc.Value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture), decimalPlaces)}"
                                : txt);
                        }
                        else { values.Add(""); }
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(formula))
                    {
                        // 분석정보의 농도 계산공식을 토큰 치환 + 평가
                        var pretty = SubstituteFormula(formula, row);
                        var vars = new Dictionary<string, double>();
                        foreach (var kv in row)
                        {
                            if (double.TryParse((kv.Value ?? "").Replace(",", ""),
                                System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out var d))
                                vars[kv.Key.Trim()] = d;
                        }
                        var v = ETA.Services.SERVICE3.AnalysisNoteService.EvaluateFormula(formula, vars);
                        values.Add(v.HasValue ? $"{pretty} = {v.Value:G6}" : pretty);
                    }
                    else
                    {
                        var conc = FirstNonEmpty(row, "계산농도", "농도");
                        var dil  = FirstValue(row, "희석배수");
                        values.Add(!string.IsNullOrWhiteSpace(conc) && !string.IsNullOrWhiteSpace(dil)
                            ? $"{conc} × {dil}" : "");
                    }
                }
                else if (h == "결과" || h == "결과값")
                {
                    var raw = FirstValue(row, h);
                    values.Add(FormatResultWithLoQ(raw, decimalPlaces, loq));
                }
                else
                {
                    values.Add(FirstValue(row, h));
                }
            }
            // 실제 시설명은 SN 컬럼에 저장됨 (시료구분 컬럼은 공백인 경우가 많음)
            var sn      = FirstValue(row, "SN");
            var sName   = FirstValue(row, "시료명");
            int ord     = orderMap.TryGetValue((sn, sName), out var id) ? id : int.MaxValue;
            // 뱃지 색상 키: SN(시설명) 우선, 없으면 소스구분
            var sCls    = !string.IsNullOrWhiteSpace(sn) ? sn : FirstValue(row, "소스구분");
            sortable.Add((ord, values, sCls));
        }
        sortable.Sort((x, y) => x.order.CompareTo(y.order));

        // QC 분리 기준: 저장 시 QAQC 소스는 SN="QC" 로 들어감 ([WasteSampleService.SaveRawData])
        // 의뢰시료 (폐수배출업소/처리시설/수질분석센터) 는 SN=업체명/시설명 → 시료분석결과 섹션
        int snColIdx = m.SampleHeaders.IndexOf("SN");
        foreach (var item in sortable)
        {
            // 모든 행을 시료분석결과(SampleRows) 에 포함 — CCV/MBK/DW/FBK 등 QC 도 함께 표시 (실험담당자 요청)
            m.SampleRows.Add(item.values);
            m.SampleClassByRow.Add(item.sampleClass);

            // QC 행은 검정곡선의 보증(QcRows) 에 별도 추가 — 검토용으로 두 군데 노출
            string sn = (snColIdx >= 0 && snColIdx < item.values.Count)
                        ? (item.values[snColIdx] ?? "").Trim()
                        : "";
            if (sn.Equals("QC", System.StringComparison.OrdinalIgnoreCase))
            {
                m.QcRows.Add(item.values);
                m.QcClassByRow.Add(item.sampleClass);
            }
        }

        return m;
    }

    /// <summary>
    /// 분석계획 메뉴와 정확히 동일한 순서를 시험기록부 viewer 에 적용.
    /// 시설 순서: GetAnalysisPlanStructure() 가 반환하는 facilities[] 의 인덱스
    /// 시료 순서: 같은 메서드의 samples[fac][] 인덱스
    /// </summary>
    private static Dictionary<(string, string), int> LoadFacilityOrderMap()
    {
        var m = new Dictionary<(string, string), int>();
        try
        {
            // 1) 분석계획 구조 — 시설/시료 순서가 페이지와 일치
            var (facilities, samples) = ETA.Services.SERVICE2.FacilityResultService
                .GetAnalysisPlanStructure();

            for (int fi = 0; fi < facilities.Length; fi++)
            {
                var fac = facilities[fi];
                if (!samples.TryGetValue(fac, out var sams)) continue;
                for (int si = 0; si < sams.Length; si++)
                {
                    int key = fi * 1_000_000 + si * 100;
                    m[(fac, sams[si])] = key;
                }
            }

            // 2) 분석계획에 없는 (시설, 시료) — 마스터에서 보조 (해당 시설은 끝으로)
            using var conn = DbConnectionFactory.CreateConnection(); conn.Open();
            if (DbConnectionFactory.TableExists(conn, "처리시설_마스터"))
            {
                // 분석계획에 등록된 시설은 fi, 미등록은 999 로
                var facIdx = new Dictionary<string, int>();
                for (int i = 0; i < facilities.Length; i++) facIdx[facilities[i]] = i;

                using var cm = conn.CreateCommand();
                cm.CommandText = "SELECT id, COALESCE(시설명,''), COALESCE(시료명,'') FROM `처리시설_마스터` ORDER BY id";
                using var rm = cm.ExecuteReader();
                int unknownCounter = 0;
                while (rm.Read())
                {
                    int id  = rm.IsDBNull(0) ? 0 : Convert.ToInt32(rm.GetValue(0));
                    var fac = rm.IsDBNull(1) ? "" : rm.GetString(1);
                    var sam = rm.IsDBNull(2) ? "" : rm.GetString(2);
                    if (m.ContainsKey((fac, sam))) continue;
                    int fi = facIdx.TryGetValue(fac, out var v) ? v : 999;
                    m[(fac, sam)] = fi * 1_000_000 + 900_000 + (unknownCounter++);
                }
            }
        }
        catch { }
        return m;
    }

    /// <summary>지수표기(4e-05) 같은 값을 0.00004 같은 소수표기로 변환.</summary>
    private static string ToDecimalStr(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        if (double.TryParse(s,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
            return d.ToString("0.##########",
                System.Globalization.CultureInfo.InvariantCulture);
        return s;
    }

    private static string FirstValue(Dictionary<string, string> row, string key)
    {
        foreach (var kv in row)
            if (string.Equals(kv.Key, key, System.StringComparison.OrdinalIgnoreCase))
                return kv.Value ?? "";
        return "";
    }

    private static string FirstNonEmpty(Dictionary<string, string> row, params string[] keys)
    {
        foreach (var k in keys)
        {
            var v = FirstValue(row, k);
            if (!string.IsNullOrWhiteSpace(v)) return v;
        }
        return "";
    }

    /// <summary>
    /// 테이블 약칭으로 시험기록부 템플릿 파일 경로를 찾는다 (없으면 null).
    /// 예: "TN" → "{templateFolder}/TN 시험기록부.xlsx"
    /// </summary>
    public static string? FindTemplateFile(string abbr)
    {
        if (string.IsNullOrWhiteSpace(abbr)) return null;
        try
        {
            var folder = ETA.Services.Common.TemplateConfiguration.Resolve("TestRecordBookFolder");
            if (!System.IO.Directory.Exists(folder)) return null;
            var candidate = System.IO.Path.Combine(folder, $"{abbr} 시험기록부.xlsx");
            if (System.IO.File.Exists(candidate)) return candidate;
            var candidate2 = System.IO.Path.Combine(folder, $"{abbr}_시험기록부.xlsx");
            if (System.IO.File.Exists(candidate2)) return candidate2;
            // 느슨한 매칭: 파일명에 약칭이 포함된 xlsx
            foreach (var f in System.IO.Directory.GetFiles(folder, "*.xlsx"))
            {
                var name = System.IO.Path.GetFileNameWithoutExtension(f);
                if (name.StartsWith(abbr, StringComparison.OrdinalIgnoreCase) && name.Contains("시험기록부"))
                    return f;
            }
        }
        catch { }
        return null;
    }

    /// <summary>시험기록부 템플릿 시트의 최소 렌더링 정보.</summary>
    public class TemplatePreview
    {
        public string FilePath  = "";
        public string SheetName = "";
        public int FirstRow, LastRow, FirstCol, LastCol;
        /// <summary>키 = (row, col), 값 = 표시 문자열 + 서식 힌트.</summary>
        public Dictionary<(int r, int c), TemplateCell> Cells = new();
        /// <summary>병합 셀: (startRow, startCol, rowSpan, colSpan).</summary>
        public List<(int r, int c, int rs, int cs)> Merged = new();
        public Dictionary<int, double> ColWidths = new();
        public Dictionary<int, double> RowHeights = new();
    }

    public class TemplateCell
    {
        public string Text = "";
        public bool Bold;
        public bool Center;
        public bool HasBorder;
        public string? BgHex;
    }

    /// <summary>템플릿 파일의 첫 시트(또는 시트명 명시)를 읽어 렌더 가능한 구조로 반환.</summary>
    public static TemplatePreview? LoadTemplatePreview(string templatePath)
    {
        try
        {
            if (!System.IO.File.Exists(templatePath)) return null;
            using var wb = new ClosedXML.Excel.XLWorkbook(templatePath);
            var ws = wb.Worksheet(1);
            var used = ws.RangeUsed();
            if (used == null) return null;

            var p = new TemplatePreview
            {
                FilePath  = templatePath,
                SheetName = ws.Name,
                FirstRow  = used.FirstRow().RowNumber(),
                LastRow   = used.LastRow().RowNumber(),
                FirstCol  = used.FirstColumn().ColumnNumber(),
                LastCol   = used.LastColumn().ColumnNumber(),
            };

            for (int r = p.FirstRow; r <= p.LastRow; r++)
            for (int c = p.FirstCol; c <= p.LastCol; c++)
            {
                var cell = ws.Cell(r, c);
                string text;
                try { text = cell.GetFormattedString() ?? ""; } catch { text = cell.Value.ToString() ?? ""; }
                var tc = new TemplateCell
                {
                    Text = text,
                    Bold = cell.Style.Font.Bold,
                    Center = cell.Style.Alignment.Horizontal == ClosedXML.Excel.XLAlignmentHorizontalValues.Center,
                    HasBorder = cell.Style.Border.TopBorder != ClosedXML.Excel.XLBorderStyleValues.None
                              || cell.Style.Border.BottomBorder != ClosedXML.Excel.XLBorderStyleValues.None
                              || cell.Style.Border.LeftBorder != ClosedXML.Excel.XLBorderStyleValues.None
                              || cell.Style.Border.RightBorder != ClosedXML.Excel.XLBorderStyleValues.None,
                };
                if (cell.Style.Fill.BackgroundColor != null
                    && cell.Style.Fill.BackgroundColor.ColorType == ClosedXML.Excel.XLColorType.Color)
                {
                    try
                    {
                        var clr = cell.Style.Fill.BackgroundColor.Color;
                        tc.BgHex = $"#{clr.R:X2}{clr.G:X2}{clr.B:X2}";
                    }
                    catch { }
                }
                p.Cells[(r, c)] = tc;
            }

            foreach (var mr in ws.MergedRanges)
            {
                int sr = mr.FirstRow().RowNumber();
                int sc = mr.FirstColumn().ColumnNumber();
                int rs = mr.LastRow().RowNumber() - sr + 1;
                int cs = mr.LastColumn().ColumnNumber() - sc + 1;
                p.Merged.Add((sr, sc, rs, cs));
            }

            // 대략적인 컬럼 폭(기본 단위 → px 변환)
            for (int c = p.FirstCol; c <= p.LastCol; c++)
            {
                try { p.ColWidths[c] = Math.Max(40, ws.Column(c).Width * 7); }
                catch { p.ColWidths[c] = 80; }
            }
            for (int r = p.FirstRow; r <= p.LastRow; r++)
            {
                try { p.RowHeights[r] = Math.Max(18, ws.Row(r).Height); }
                catch { p.RowHeights[r] = 22; }
            }

            return p;
        }
        catch { return null; }
    }

    /// <summary>
    /// 선택된 (테이블, 날짜) 를 엑셀로 출력.
    /// 약칭(abbr) 에 매칭되는 템플릿이 있으면 템플릿을 복사해서 사용, 없으면 단순 표.
    /// 반환: 저장된 파일 경로(실패시 null).
    /// </summary>
    public static string? ExportToExcel(string tableName, string date, string savePath, string abbr)
    {
        try
        {
            var (cols, rows) = GetRowsByDate(tableName, date);
            if (rows.Count == 0) return null;

            var templatePath = FindTemplateFile(abbr);
            if (templatePath != null)
            {
                // 템플릿 복사 → 후속 데이터 매핑은 항목별 작업 필요
                System.IO.File.Copy(templatePath, savePath, overwrite: true);
                return savePath;
            }

            // 템플릿 없음 → 단순 표 생성
            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add(PrettyName(tableName));

            ws.Cell(1, 1).Value = $"{PrettyName(tableName)} 시험기록부";
            ws.Cell(1, 1).Style.Font.FontSize = 16;
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Range(1, 1, 1, Math.Max(cols.Count, 6)).Merge();

            ws.Cell(2, 1).Value = $"분석일: {date}";
            ws.Cell(2, 1).Style.Font.Bold = true;

            for (int i = 0; i < cols.Count; i++)
            {
                ws.Cell(4, i + 1).Value = cols[i];
                ws.Cell(4, i + 1).Style.Font.Bold = true;
                ws.Cell(4, i + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                ws.Cell(4, i + 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            }

            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < cols.Count; c++)
                    ws.Cell(5 + r, c + 1).Value = rows[r].TryGetValue(cols[c], out var v) ? v : "";

            var tableRange = ws.Range(4, 1, 4 + rows.Count, cols.Count);
            tableRange.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder  = ClosedXML.Excel.XLBorderStyleValues.Thin;
            try { ws.Columns().AdjustToContents(); } catch { }

            int signRow = 5 + rows.Count + 2;
            ws.Cell(signRow, 1).Value = "분석자";   ws.Cell(signRow, 2).Value = "(인)";
            ws.Cell(signRow + 1, 1).Value = "기술책임자"; ws.Cell(signRow + 1, 2).Value = "(인)";

            wb.SaveAs(savePath);
            return savePath;
        }
        catch
        {
            return null;
        }
    }

    // 테이블 이름 변환 규칙: 분석정보.Analyte 의 특수문자(, - . 공백 등)는 제거/통일해 비교.
    private static string NormalizeForTable(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var sb = new System.Text.StringBuilder();
        foreach (var ch in s)
        {
            if (ch == ',' || ch == '-' || ch == '.' || ch == ' ' || ch == '_'
                || ch == '(' || ch == ')' || ch == '/' || ch == '·') continue;
            sb.Append(ch);
        }
        return sb.ToString();
    }
}
