using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ETA.Services.SERVICE4;

// ──────────────────────────────────────────────────────────────────────────
// Agilent ChemStation / MassHunter GC CSV 파서
//
// 지원 포맷 (Data/GC 자료/ 샘플 기준):
//   VocMulti      : 다성분 ISTD (VOC3.csv, VOC10.csv)
//   VocSingle     : 단일성분 ISTD (1,4 다이옥산.csv, 나프탈렌.csv, 스타이렌.csv,
//                                   페놀.csv, 아크릴아미드.csv, 에피클로로하이드린.csv,
//                                   폼알데하이드.csv, 펜타클로로페놀.csv)
//   SingleNoIstd  : 단일성분 비ISTD (PCB.csv, DEHA.csv, DEHP.csv)
//   SingleExpConc : 단일성분 비ISTD + Exp. Conc. 칼럼 (유기인.csv)
//
// 공통 섹션 구조:
//   1) 배치 헤더 (Batch Data Path / Analysis Time / ...)
//   2) 성분 섹션: 성분명 1줄 → Data File,...,Final Conc. 헤더 → Cal/Sample/QC/Blank 행
//   3) Calibration 섹션: 성분명 → Cal. Path,Type,Level,Enable,Conc.,Response
//   4) 샘플 상세 블록 반복: Acq. Date-Time / Data File / Dil. / ... 포맷으로 희석배수 추출
// ──────────────────────────────────────────────────────────────────────────

public enum GcFileFormat
{
    Unknown,
    VocMulti,       // 다성분 + ISTD
    VocSingle,      // 단일성분 + ISTD
    SingleNoIstd,   // 단일성분 비ISTD
    SingleExpConc,  // 단일성분 비ISTD + Exp.Conc
}

public class GcInstrumentRow
{
    public string RawName   { get; set; } = "";    // 예: "RE-07.D", "VOC3 ST-01.D"
    public string SN        { get; set; } = "";    // 정규화 SN (03-11-04 형식)
    public string Type      { get; set; } = "";    // Cal / Sample / Blank / QC
    public string Rt        { get; set; } = "";
    public string Resp      { get; set; } = "";
    public string IstdResp  { get; set; } = "";
    public string FinalConc { get; set; } = "";    // "ND" 가능
    public string ExpConc   { get; set; } = "";    // SingleExpConc 포맷에서만
    public string Dilution  { get; set; } = "";    // 상세블록에서 조인
    public bool   IsControl { get; set; } = false; // Cal/QC/Blank/CCV/FBK/MBK/BK 등
}

public class GcCalibrationPoint
{
    public int    Level    { get; set; }
    public string Enable   { get; set; } = "x";
    public string Conc     { get; set; } = "";
    public string Response { get; set; } = "";
    public string IstdResponse { get; set; } = "";
}

public class GcCompound
{
    public string                     Name          { get; set; } = "";
    public bool                       HasIstd       { get; set; }
    public bool                       HasExpConc    { get; set; }
    public List<GcInstrumentRow>      Rows          { get; } = new();
    public List<GcCalibrationPoint>   Calibration   { get; } = new();
    // 선형회귀 y = a*x + b, R
    public double? SlopeA     { get; set; }
    public double? Intercept  { get; set; }
    public double? R          { get; set; }
}

public class GcInstrumentFile
{
    public GcFileFormat    Format        { get; set; } = GcFileFormat.Unknown;
    public string          SourcePath    { get; set; } = "";
    public string          BatchDataPath { get; set; } = "";
    public string          AnalysisTime  { get; set; } = "";
    public string          AnalystName   { get; set; } = "";
    public List<GcCompound> Compounds    { get; } = new();
}

public static class GcInstrumentParser
{
    // CP949/EUC-KR 우선, UTF-8 BOM 감지 시 UTF-8 사용
    private static string ReadAllTextEucKr(string path)
    {
        // BOM 검사
        var head = new byte[3];
        using (var fs = File.OpenRead(path))
        {
            int n = fs.Read(head, 0, 3);
            if (n >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF)
                return File.ReadAllText(path, Encoding.UTF8);
        }
        try { return File.ReadAllText(path, Encoding.GetEncoding(949)); }
        catch { return File.ReadAllText(path, Encoding.UTF8); }
    }

    public static GcFileFormat DetectFormat(string path)
    {
        string text;
        try { text = ReadAllTextEucKr(path); }
        catch { return GcFileFormat.Unknown; }

        if (!text.Contains("Batch Data Path", StringComparison.OrdinalIgnoreCase))
            return GcFileFormat.Unknown;

        // Agilent 공통 마커는 확인됨. 헤더 행 패턴으로 포맷 구분.
        bool hasIstd     = Regex.IsMatch(text, @"Data File,,Type,RT,,Resp\.,,ISTD Resp", RegexOptions.IgnoreCase);
        bool hasNoIstd   = Regex.IsMatch(text, @"Data File,Type,RT,Resp\.,Final Conc\.", RegexOptions.IgnoreCase);
        bool hasExpConc  = Regex.IsMatch(text, @"Data File,Type,RT,Resp\.,Final Conc\.,Exp\. Conc\.", RegexOptions.IgnoreCase);

        if (hasIstd)
        {
            // 성분 섹션 개수로 VocMulti vs VocSingle 판정
            int compoundCount = CountCompoundSections(text, istd: true);
            return compoundCount >= 2 ? GcFileFormat.VocMulti : GcFileFormat.VocSingle;
        }
        if (hasExpConc) return GcFileFormat.SingleExpConc;
        if (hasNoIstd)  return GcFileFormat.SingleNoIstd;
        return GcFileFormat.Unknown;
    }

    private static int CountCompoundSections(string text, bool istd)
    {
        var pattern = istd
            ? @"Data File,,Type,RT,,Resp\."
            : @"Data File,Type,RT,Resp\.";
        return Regex.Matches(text, pattern, RegexOptions.IgnoreCase).Count;
    }

    public static GcInstrumentFile Parse(string path)
    {
        var file = new GcInstrumentFile { SourcePath = path };
        string text;
        try { text = ReadAllTextEucKr(path); }
        catch { file.Format = GcFileFormat.Unknown; return file; }

        file.Format = DetectFormat(path);
        if (file.Format == GcFileFormat.Unknown) return file;

        var lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

        // 1) 배치 헤더
        ExtractBatchHeader(lines, file);

        // 2) 성분 섹션 파싱
        ParseCompoundSections(lines, file);

        // 3) Calibration 섹션 파싱 + 선형회귀
        ParseCalibrationSections(lines, file);

        // 4) 상세 블록에서 Data File → Dil. 조인
        ApplyDilutionFromDetailBlocks(lines, file);

        // 5) SN 정규화
        foreach (var c in file.Compounds)
            foreach (var r in c.Rows)
                r.SN = NormalizeSn(r.RawName);

        return file;
    }

    private static void ExtractBatchHeader(string[] lines, GcInstrumentFile file)
    {
        foreach (var raw in lines.Take(6))
        {
            var cells = SplitCsvLine(raw);
            if (cells.Length == 0) continue;
            if (cells[0].Equals("Batch Data Path", StringComparison.OrdinalIgnoreCase) && cells.Length > 1)
                file.BatchDataPath = cells[1];
            if (cells[0].Equals("Analysis Time", StringComparison.OrdinalIgnoreCase) && cells.Length > 1)
            {
                file.AnalysisTime = cells[1];
                for (int i = 2; i + 1 < cells.Length; i += 2)
                {
                    if (cells[i].Equals("Analyst Name", StringComparison.OrdinalIgnoreCase))
                        file.AnalystName = cells[i + 1];
                }
            }
        }
    }

    // 성분 섹션: 한 줄짜리 성분명 → 헤더 → 데이터 행 → (다음 성분 / Calibration 섹션)
    private static void ParseCompoundSections(string[] lines, GcInstrumentFile file)
    {
        bool istd       = file.Format is GcFileFormat.VocMulti or GcFileFormat.VocSingle;
        bool hasExpConc = file.Format is GcFileFormat.SingleExpConc;

        int i = 0;
        while (i < lines.Length)
        {
            var line = lines[i];
            var cells = SplitCsvLine(line);

            // 헤더 행 탐지
            bool isHeader =
                cells.Length > 0 &&
                cells[0].Equals("Data File", StringComparison.OrdinalIgnoreCase) &&
                line.Contains("Final Conc", StringComparison.OrdinalIgnoreCase);

            if (isHeader)
            {
                // 직전 성분명 추출 (바로 위 또는 두 줄 위)
                string compoundName = "";
                for (int back = 1; back <= 3 && i - back >= 0; back++)
                {
                    var prev = SplitCsvLine(lines[i - back]);
                    if (prev.Length == 0) continue;
                    var prevFirst = prev[0].Trim();
                    if (string.IsNullOrWhiteSpace(prevFirst)) continue;
                    // 헤더/Calibration 레이블이면 스킵
                    if (prevFirst.Equals("Calibration", StringComparison.OrdinalIgnoreCase)) continue;
                    if (prevFirst.StartsWith("Batch Data Path", StringComparison.OrdinalIgnoreCase)) continue;
                    if (prevFirst.StartsWith("Analysis Time", StringComparison.OrdinalIgnoreCase))   continue;
                    if (prevFirst.StartsWith("Calibration Last", StringComparison.OrdinalIgnoreCase)) continue;
                    compoundName = prevFirst;
                    break;
                }

                var compound = new GcCompound
                {
                    Name       = compoundName,
                    HasIstd    = istd,
                    HasExpConc = hasExpConc,
                };

                // 데이터 행 읽기 (빈 줄 또는 다른 섹션 시작까지)
                int j = i + 1;
                while (j < lines.Length)
                {
                    var rowLine  = lines[j];
                    var rowCells = SplitCsvLine(rowLine);

                    if (string.IsNullOrWhiteSpace(rowLine)) { j++; break; }
                    // 다음 섹션 전환 감지
                    if (rowCells.Length > 0)
                    {
                        var first = rowCells[0].Trim();
                        if (first.Equals("Calibration", StringComparison.OrdinalIgnoreCase)) break;
                        if (first.Equals("Data File",   StringComparison.OrdinalIgnoreCase)) break;
                        if (first.StartsWith("Batch Data Path", StringComparison.OrdinalIgnoreCase)) break;
                    }

                    var row = ParseDataRow(rowCells, istd, hasExpConc);
                    if (row != null) compound.Rows.Add(row);
                    j++;
                }

                file.Compounds.Add(compound);
                i = j;
                continue;
            }

            i++;
        }
    }

    private static GcInstrumentRow? ParseDataRow(string[] cells, bool istd, bool hasExpConc)
    {
        // ISTD 포맷 레이아웃 (공백 칼럼 포함):
        //   Data File, , Type, RT, , Resp., , ISTD Resp, , Final Conc.
        //   idx:      0  1   2    3  4   5   6   7        8   9
        // 비ISTD 포맷:
        //   Data File, Type, RT, Resp., Final Conc.[, Exp. Conc.]
        //   idx:      0      1    2    3              4

        if (cells.Length < 3) return null;
        var dataFile = cells[0].Trim();
        if (string.IsNullOrEmpty(dataFile)) return null;
        // Calibration 같은 섹션 레이블은 이미 필터됐지만 한 번 더 확인
        if (dataFile.Equals("Data File", StringComparison.OrdinalIgnoreCase)) return null;

        var row = new GcInstrumentRow { RawName = dataFile };

        if (istd)
        {
            if (cells.Length < 10) return null;
            row.Type      = cells[2].Trim();
            row.Rt        = cells[3].Trim();
            row.Resp      = cells[5].Trim();
            row.IstdResp  = cells[7].Trim();
            row.FinalConc = cells[9].Trim();
        }
        else
        {
            if (cells.Length < 5) return null;
            row.Type      = cells[1].Trim();
            row.Rt        = cells[2].Trim();
            row.Resp      = cells[3].Trim();
            row.FinalConc = cells[4].Trim();
            if (hasExpConc && cells.Length >= 6)
                row.ExpConc = cells[5].Trim();
        }

        row.IsControl = IsControlType(row.Type) || IsControlName(dataFile);
        return row;
    }

    // Calibration 섹션:
    //   Calibration,,
    //   <compound>,,
    //   <empty>
    //   Cal. Path,Type,Level,Enable,,Conc.,,Response   (ISTD 포맷)
    //   Cal. Path,Type,Level,Enable,Conc.,Response     (비ISTD)
    //   <경로>,Calibration,1,x,,<Conc>,,<Resp>
    private static void ParseCalibrationSections(string[] lines, GcInstrumentFile file)
    {
        bool istd = file.Format is GcFileFormat.VocMulti or GcFileFormat.VocSingle;

        for (int i = 0; i < lines.Length; i++)
        {
            var cells = SplitCsvLine(lines[i]);
            if (cells.Length == 0) continue;
            if (!cells[0].Trim().Equals("Cal. Path", StringComparison.OrdinalIgnoreCase)) continue;

            // 이 헤더 위로 성분명 찾기
            string compoundName = "";
            for (int back = 1; back <= 5 && i - back >= 0; back++)
            {
                var prev = SplitCsvLine(lines[i - back]);
                if (prev.Length == 0) continue;
                var first = prev[0].Trim();
                if (string.IsNullOrWhiteSpace(first)) continue;
                if (first.Equals("Calibration", StringComparison.OrdinalIgnoreCase)) continue;
                // "Naphthalene,,", "1,4-Dioxane" 같은 성분 라인 (쉼표 이후 비어있음)
                compoundName = first;
                break;
            }

            var compound = file.Compounds.FirstOrDefault(c =>
                string.Equals(c.Name, compoundName, StringComparison.OrdinalIgnoreCase));
            if (compound == null)
            {
                // 이름 불일치 시 첫 번째 Calibration 미설정 성분에 할당
                compound = file.Compounds.FirstOrDefault(c => c.Calibration.Count == 0);
            }
            if (compound == null) continue;

            int j = i + 1;
            while (j < lines.Length)
            {
                var cc = SplitCsvLine(lines[j]);
                if (cc.Length == 0) { j++; continue; }
                // 섹션 종료 감지: Batch/Data File 등이 오면 끝
                var head = cc[0].Trim();
                if (head.StartsWith("Batch Data Path", StringComparison.OrdinalIgnoreCase)) break;
                if (head.Equals("Data File", StringComparison.OrdinalIgnoreCase)) break;
                if (head.Equals("Cal. Path", StringComparison.OrdinalIgnoreCase)) break;

                // 예상 칼럼:
                //   ISTD:  Path, Calibration, Level, Enable, , Conc, , Response   (8)
                //   No:    Path, Calibration, Level, Enable, Conc, Response       (6)
                // Type이 "Cal"인 것만 취함 (QC 행은 스킵)
                if (cc.Length >= 2 && cc[1].Trim().Equals("Cal", StringComparison.OrdinalIgnoreCase))
                {
                    var pt = new GcCalibrationPoint();
                    if (int.TryParse(cc[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var lvl))
                        pt.Level = lvl;
                    if (cc.Length >= 4) pt.Enable = cc[3].Trim();
                    if (istd)
                    {
                        if (cc.Length >= 6)
                        {
                            pt.Conc         = cc[5].Trim(); // Final Conc.
                            pt.Response     = cc[3].Trim(); // Resp. (main compound)
                            pt.IstdResponse = cc[4].Trim(); // ISTD Resp

                            // 디버깅 로그
                            System.Diagnostics.Debug.WriteLine($"[GC Cal ISTD] {cc[0]}: Conc={cc[5]}, Resp={cc[3]}, ISTD={cc[4]}");
                        }
                    }
                    else
                    {
                        if (cc.Length >= 6)
                        {
                            pt.Conc     = cc[4].Trim();
                            pt.Response = cc[5].Trim();
                        }
                    }
                    compound.Calibration.Add(pt);
                }
                j++;
            }

            // 선형회귀 계산
            ComputeLinearRegression(compound);
            // 종료 라인(예: 다음 성분의 Cal. Path / Batch Data Path)을 for 루프가 다시 보도록
            i = j - 1;
        }
    }

    private static void ComputeLinearRegression(GcCompound compound)
    {
        var xs = new List<double>();
        var ys = new List<double>();
        foreach (var p in compound.Calibration)
        {
            if (double.TryParse(p.Conc,     NumberStyles.Float, CultureInfo.InvariantCulture, out var x) &&
                double.TryParse(p.Response, NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
            {
                xs.Add(x); ys.Add(y);
            }
        }
        int n = xs.Count;
        if (n < 2) return;

        double sx = xs.Sum(), sy = ys.Sum();
        double sxx = xs.Sum(v => v * v);
        double sxy = 0;
        for (int k = 0; k < n; k++) sxy += xs[k] * ys[k];
        double denom = n * sxx - sx * sx;
        if (Math.Abs(denom) < 1e-12) return;

        double a = (n * sxy - sx * sy) / denom;
        double b = (sy - a * sx) / n;

        double meanY = sy / n;
        double ssTot = ys.Sum(y => (y - meanY) * (y - meanY));
        double ssRes = 0;
        for (int k = 0; k < n; k++)
        {
            double yp = a * xs[k] + b;
            ssRes += (ys[k] - yp) * (ys[k] - yp);
        }
        double r2 = ssTot > 0 ? 1 - ssRes / ssTot : 0;
        compound.SlopeA    = a;
        compound.Intercept = b;
        compound.R         = Math.Sqrt(Math.Max(0, r2));
    }

    // 상세 블록: 반복되는 단락. 핵심은 두 줄:
    //   ..., Data File, <name>
    //   Dil.,<value>,Acq. Method File,<method>
    // 파일 전체 스윕으로 Data File → Dil. 매핑 추출
    private static void ApplyDilutionFromDetailBlocks(string[] lines, GcInstrumentFile file)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string? currentDataFile = null;

        for (int i = 0; i < lines.Length; i++)
        {
            var cells = SplitCsvLine(lines[i]);
            if (cells.Length == 0) continue;

            // "Acq. Date-Time,...,Data File,<name>" 같은 행
            for (int k = 0; k + 1 < cells.Length; k++)
            {
                if (cells[k].Trim().Equals("Data File", StringComparison.OrdinalIgnoreCase))
                {
                    var name = cells[k + 1].Trim();
                    if (!string.IsNullOrEmpty(name) &&
                        !name.Equals("Data File", StringComparison.OrdinalIgnoreCase))
                    {
                        currentDataFile = name;
                    }
                    break;
                }
            }

            // "Dil.,<v>,..." 행
            if (cells[0].Trim().Equals("Dil.", StringComparison.OrdinalIgnoreCase) && cells.Length >= 2)
            {
                if (!string.IsNullOrEmpty(currentDataFile))
                {
                    map[currentDataFile] = cells[1].Trim();
                }
            }
        }

        if (map.Count == 0) return;
        foreach (var c in file.Compounds)
            foreach (var r in c.Rows)
                if (map.TryGetValue(r.RawName, out var dil)) r.Dilution = dil;
    }

    // ─── 유틸 ──────────────────────────────────────────────────────────────

    private static bool IsControlType(string type)
    {
        if (string.IsNullOrWhiteSpace(type)) return false;
        var t = type.Trim();
        return t.Equals("Cal",   StringComparison.OrdinalIgnoreCase)
            || t.Equals("QC",    StringComparison.OrdinalIgnoreCase)
            || t.Equals("Blank", StringComparison.OrdinalIgnoreCase);
    }

    private static readonly string[] ControlNameKeywords = new[]
    {
        "ST-", "ST ", " ST", "CCV", "FBK", "MBK", "BLANK", "BK-", "BK.", "-CCV-", "-C-C-V-", "DW"
    };

    private static bool IsControlName(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return false;
        var upper = rawName.ToUpperInvariant();
        foreach (var k in ControlNameKeywords)
            if (upper.Contains(k)) return true;
        return false;
    }

    // SN 정규화: "RE-07.D" → "03-11-04" 같은 형식으로는 일단 변환 불가.
    // 여기서는 원본에서 ".D"/".d" 확장자만 제거한 코드 반환.
    // 실제 SN 매칭은 상위 레이어에서 WasteSample 테이블과 조인 시 수행.
    private static string NormalizeSn(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return "";
        var s = rawName.Trim();
        if (s.EndsWith(".D", StringComparison.OrdinalIgnoreCase))
            s = s[..^2];
        return s;
    }

    // 인용부호 포함 CSV 한 줄 분할 (TocInstrumentParser와 동일 규칙)
    private static string[] SplitCsvLine(string line)
    {
        if (string.IsNullOrEmpty(line)) return Array.Empty<string>();
        var list = new List<string>();
        var sb   = new StringBuilder();
        bool inQ = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (inQ && i + 1 < line.Length && line[i + 1] == '"')
                {
                    sb.Append('"'); i++;
                }
                else
                {
                    inQ = !inQ;
                }
            }
            else if (c == ',' && !inQ)
            {
                list.Add(sb.ToString());
                sb.Clear();
            }
            else
            {
                sb.Append(c);
            }
        }
        list.Add(sb.ToString());
        return list.ToArray();
    }
}
