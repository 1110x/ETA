using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace ETA.Services.SERVICE4;

// ── 파서 시그니처 모델 ────────────────────────────────────────────────────
public class ParserSignature
{
    public int          Id              { get; set; }
    public string       ParserKey       { get; set; } = "";
    public string       ParserName      { get; set; } = "";
    public List<string> StdKeywords     { get; set; } = new();  // 표준물질 키워드 (weight 2)
    public List<string> HeaderKeywords  { get; set; } = new();  // 헤더/컬럼명  (weight 3)
    public List<string> SampleKeywords  { get; set; } = new();  // 시료명 패턴  (weight 1)
    public string       Delimiter       { get; set; } = ",";
    public int          HeaderLines     { get; set; } = 1;
    public int          NameColumnIndex { get; set; } = 0;
    public string       RegisteredDate  { get; set; } = "";
}

// ── 분류 결과 ─────────────────────────────────────────────────────────────
public record SignatureHit(string ParserKey, string ParserName, float Score, string Source);
//   Source: "Signature" | "ONNX" | "Both"

/// <summary>
/// 파서 시그니처 DB 기반 분류기.
/// AI 문서분류(ONNX)와 독립적으로 동작하며, 없을 때 보완 역할.
/// </summary>
public static class SignatureClassifier
{
    // ── 공유 파서 목록 (AiDocClassificationPage, ParserGeneratorPage 공용) ──
    public static readonly List<(string Name, string Key)> ParserItems =
    [
        ("BOD",                  "BOD"),
        ("SS",                   "SS"),
        ("N-Hexan",              "NHex"),
        ("UVVIS",                "UVVIS"),
        ("TOC-시마즈 (CSV/TXT)", "TOC_Shimadzu"),
        ("TOC-시마즈 (PDF)",     "TOC_Shimadzu_PDF"),
        ("TOC-예나 (PDF)",       "TOC_NPOC"),
        ("TOC-스칼라 NPOC",      "TOC_Scalar_NPOC"),
        ("TOC-스칼라 TCIC",      "TOC_Scalar_TCIC"),
        ("GC",                   "GC"),
        ("UV-Shimadzu (PDF)",    "UV_Shimadzu_PDF"),
        ("UV-Shimadzu (ASCII)",  "UV_Shimadzu_ASCII"),
        ("UV-Cary (PDF)",        "UV_Cary_PDF"),
        ("UV-Cary (CSV)",        "UV_Cary_CSV"),
        ("ICP",                  "ICP"),
        ("LCMS",                 "LCMS"),
        ("AA-수은 (CSV)",        "AA_HG_CSV"),
        ("AA-수은 (PDF)",        "AA_HG_PDF"),
    ];

    // ── DB 초기화 ─────────────────────────────────────────────────────
    private static bool _ensured;
    public static void EnsureTable()
    {
        if (_ensured) return;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "파서_시그니처"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"
                    CREATE TABLE `파서_시그니처` (
                        `_id`           INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                        `파서키`        TEXT NOT NULL,
                        `파서명`        TEXT NOT NULL,
                        `표준물질_키워드` TEXT,
                        `헤더_키워드`   TEXT,
                        `시료명_키워드` TEXT,
                        `구분자`        TEXT DEFAULT ',',
                        `헤더행수`      INTEGER DEFAULT 1,
                        `시료명_열`     INTEGER DEFAULT 0,
                        `등록일`        TEXT
                    )";
                cmd.ExecuteNonQuery();
            }
            _ensured = true;
        }
        catch (Exception ex) { }
    }

    // ── 저장 (같은 파서키면 덮어씀) ──────────────────────────────────
    public static void SaveSignature(ParserSignature sig)
    {
        EnsureTable();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();

            using var del = conn.CreateCommand();
            del.CommandText = "DELETE FROM `파서_시그니처` WHERE `파서키`=@k";
            AddParam(del, "@k", sig.ParserKey);
            del.ExecuteNonQuery();

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "INSERT INTO `파서_시그니처` (`파서키`,`파서명`,`표준물질_키워드`,`헤더_키워드`," +
                "`시료명_키워드`,`구분자`,`헤더행수`,`시료명_열`,`등록일`) " +
                "VALUES (@k,@n,@s,@h,@sm,@d,@hl,@nc,@dt)";
            AddParam(cmd, "@k",  sig.ParserKey);
            AddParam(cmd, "@n",  sig.ParserName);
            AddParam(cmd, "@s",  JsonSerializer.Serialize(sig.StdKeywords));
            AddParam(cmd, "@h",  JsonSerializer.Serialize(sig.HeaderKeywords));
            AddParam(cmd, "@sm", JsonSerializer.Serialize(sig.SampleKeywords));
            AddParam(cmd, "@d",  sig.Delimiter);
            AddParam(cmd, "@hl", sig.HeaderLines);
            AddParam(cmd, "@nc", sig.NameColumnIndex);
            AddParam(cmd, "@dt", DateTime.Now.ToString("yyyy-MM-dd"));
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            throw;
        }
    }

    // ── 전체 로드 ─────────────────────────────────────────────────────
    public static List<ParserSignature> LoadAll()
    {
        EnsureTable();
        var list = new List<ParserSignature>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT `_id`,`파서키`,`파서명`,`표준물질_키워드`,`헤더_키워드`," +
                "`시료명_키워드`,`구분자`,`헤더행수`,`시료명_열`,`등록일` " +
                "FROM `파서_시그니처` ORDER BY `파서키`";
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new ParserSignature
                {
                    Id              = r.GetInt32(0),
                    ParserKey       = r.IsDBNull(1) ? "" : r.GetString(1),
                    ParserName      = r.IsDBNull(2) ? "" : r.GetString(2),
                    StdKeywords     = DeJson(r.IsDBNull(3) ? null : r.GetString(3)),
                    HeaderKeywords  = DeJson(r.IsDBNull(4) ? null : r.GetString(4)),
                    SampleKeywords  = DeJson(r.IsDBNull(5) ? null : r.GetString(5)),
                    Delimiter       = r.IsDBNull(6) ? "," : r.GetString(6),
                    HeaderLines     = r.IsDBNull(7) ? 1   : r.GetInt32(7),
                    NameColumnIndex = r.IsDBNull(8) ? 0   : r.GetInt32(8),
                    RegisteredDate  = r.IsDBNull(9) ? ""  : r.GetString(9),
                });
        }
        catch (Exception ex) { }
        return list;
    }

    public static int Count()
    {
        EnsureTable();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM `파서_시그니처`";
            var v = cmd.ExecuteScalar();
            return v == null || v == DBNull.Value ? 0 : Convert.ToInt32(v);
        }
        catch { return 0; }
    }

    public static bool HasSignatures() => Count() > 0;

    // ── 분류 ─────────────────────────────────────────────────────────
    /// <summary>가장 점수 높은 파서키 반환. 점수 15% 미만이면 null.</summary>
    public static string? Classify(string text)
    {
        var hits = ClassifyTopK(text, 1);
        return hits.Count > 0 && hits[0].Score >= 0.15f ? hits[0].ParserKey : null;
    }

    /// <summary>점수 상위 k개 반환 (score > 0인 것만).</summary>
    public static List<SignatureHit> ClassifyTopK(string text, int k = 3)
    {
        return LoadAll()
            .Select(sig => new SignatureHit(sig.ParserKey, sig.ParserName, ScoreText(text, sig), "Signature"))
            .Where(h => h.Score > 0)
            .OrderByDescending(h => h.Score)
            .Take(k)
            .ToList();
    }

    private static float ScoreText(string text, ParserSignature sig)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0f;

        float maxScore = sig.StdKeywords.Count * 2f
                       + sig.HeaderKeywords.Count * 3f
                       + sig.SampleKeywords.Count * 1f;
        if (maxScore == 0) return 0f;

        var lower = text.ToLower();
        float score = 0f;
        foreach (var kw in sig.StdKeywords)
            if (!string.IsNullOrWhiteSpace(kw) && lower.Contains(kw.ToLower())) score += 2f;
        foreach (var kw in sig.HeaderKeywords)
            if (!string.IsNullOrWhiteSpace(kw) && lower.Contains(kw.ToLower())) score += 3f;
        foreach (var kw in sig.SampleKeywords)
            if (!string.IsNullOrWhiteSpace(kw) && lower.Contains(kw.ToLower())) score += 1f;

        return score / maxScore;
    }

    // ── AI 문서분류 학습 데이터에서 시그니처 자동 빌드 ──────────────
    /// <returns>(빌드 성공 수, 건너뜀 수)</returns>
    public static (int Built, int Skipped) BuildFromTrainingData()
    {
        EnsureTable();
        var byParser = new Dictionary<string, List<string>>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT `파서타입`, `파일텍스트` FROM `AI_문서분류_학습데이터` " +
                "WHERE `파일텍스트` IS NOT NULL AND `파일텍스트` != ''";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var key  = r.IsDBNull(0) ? "" : r.GetString(0);
                var text = r.IsDBNull(1) ? "" : r.GetString(1);
                if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(text)) continue;
                if (!byParser.ContainsKey(key)) byParser[key] = new();
                byParser[key].Add(text);
            }
        }
        catch (Exception ex)
        {
            return (0, 0);
        }

        int built = 0, skipped = 0;
        foreach (var (key, texts) in byParser)
        {
            var name = ParserItems.FirstOrDefault(p => p.Key == key).Name ?? key;
            var sig  = BuildSignatureFromTexts(key, name, texts);
            if (sig.StdKeywords.Count + sig.HeaderKeywords.Count == 0) { skipped++; continue; }
            try { SaveSignature(sig); built++; }
            catch { skipped++; }
        }
        return (built, skipped);
    }

    private static readonly string[] StdPrefixes =
        ["STD", "Standard", "Cal", "Blank", "BLK", "QC", "Check", "CCV", "ICV", "CCB", "ICB",
         "공시험", "표준", "표준액", "검량"];

    private static ParserSignature BuildSignatureFromTexts(string key, string name, List<string> texts)
    {
        var combined = string.Join("\n", texts);
        var lines    = combined.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0).ToArray();
        char delim   = combined.Count(c => c == ',') > combined.Count(c => c == '\t') ? ',' : '\t';

        // 표준물질 키워드
        var stdFound = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in lines)
        {
            var first = line.Split(new[] { ',', '\t', ' ' }, 2)[0].Trim('"', ' ');
            foreach (var pfx in StdPrefixes)
                if (first.StartsWith(pfx, StringComparison.OrdinalIgnoreCase)) { stdFound.Add(pfx); break; }
        }

        // 헤더 키워드: 처음 20행에 반복 등장하는 단어 (숫자 제외, 3자 이상)
        var headerKw = lines.Take(20)
            .SelectMany(l => l.Split(new[] { ',', '\t' }))
            .Select(t => t.Trim('"', ' '))
            .Where(t => t.Length >= 3 && !double.TryParse(t, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _))
            .GroupBy(t => t, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() >= Math.Max(2, texts.Count))
            .OrderByDescending(g => g.Count())
            .Select(g => g.Key)
            .Take(8)
            .ToList();

        return new ParserSignature
        {
            ParserKey       = key,
            ParserName      = name,
            StdKeywords     = stdFound.Take(6).ToList(),
            HeaderKeywords  = headerKw,
            RegisteredDate  = DateTime.Now.ToString("yyyy-MM-dd"),
            Delimiter       = delim.ToString(),
        };
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────
    private static void AddParam(System.Data.Common.DbCommand cmd, string name, object val)
    {
        var p = cmd.CreateParameter();
        p.ParameterName = name;
        p.Value         = val;
        cmd.Parameters.Add(p);
    }

    private static List<string> DeJson(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return new();
        try { return JsonSerializer.Deserialize<List<string>>(json) ?? new(); }
        catch { return new(); }
    }
}
