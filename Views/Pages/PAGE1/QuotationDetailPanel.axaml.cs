using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

public partial class QuotationDetailPanel : UserControl
{
    private QuotationIssue?             _current;
    private Dictionary<string, string> _cachedRow = new(StringComparer.OrdinalIgnoreCase);
    public  QuotationCheckPanel?  CheckPanel    { get; set; }
    /// <summary>🥕 당근 — 이 건을 재활용해서 신규 작성 (항목 복사, 번호·날짜 신규)</summary>
    public event Action<QuotationIssue>? CarrotRequested;
    /// <summary>✏️ 오작성 수정 — 시료명·견적번호·발행일·적용구분·업체명 등 메타 수정</summary>
    public event Action<QuotationIssue>? CorrectRequested;

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // ── 고정 컬럼 (항목 순회 제외) — Trim 된 실제 컬럼명 기준 ────────────
    private static readonly HashSet<string> FixedCols =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "_id","rowid",
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자","담당자연락처","담당자 e-Mail","시료명","견적번호",
            "적용구분","견적작성","합계 금액",   // ← 공백 없이 Trim된 값
            "수량","단가","소계","수량2","단가3","소계4",
        };

    public QuotationDetailPanel() => InitializeComponent();

    // ══════════════════════════════════════════════════════════════════════
    //  외부 연동
    // ══════════════════════════════════════════════════════════════════════
    public void ShowIssue(QuotationIssue issue)
    {
        _current = issue;
        txbEmpty.IsVisible  = false;
        spContent.IsVisible = true;
        spButtons.IsVisible = true;

        txbCompany.Text    = issue.업체명;
        txbAbbr.Text       = issue.약칭;
        txbSampleName.Text = issue.시료명;
        txbNo.Text         = issue.견적번호;
        txbDate.Text       = issue.발행일;
        txbType.Text       = issue.견적구분;
        txbAmount.Text     = issue.총금액 > 0 ? $"{issue.총금액:#,0} 원" : "—";

        Log($"=== ShowIssue rowid={issue.Id}  {issue.업체명}  {issue.시료명} ===");

        Dictionary<string, string> row;
        try
        {
            row = QuotationService.GetIssueRow(issue.Id);
            Log($"row 컬럼수={row.Count}");

            // 고정 컬럼 제외한 항목 컬럼 목록만 로그
            var itemCols = row.Keys
                .Where(k => !k.EndsWith("단가") && !k.EndsWith("소계") && !FixedCols.Contains(k))
                .OrderBy(k => k)
                .ToList();
            Log($"항목 컬럼 후보({itemCols.Count}개): {string.Join(", ", itemCols.Take(20))}");
        }
        catch (Exception ex)
        {
            Log($"GetIssueRow 오류: {ex.Message}");
            return;
        }

        _cachedRow = row;

        // 담당자 정보
        txbManagerName.Text  = issue.담당자;
        txbManagerPhone.Text = row.TryGetValue("담당자연락처",  out var ph) ? ph : issue.담당자연락처;
        txbManagerEmail.Text = row.TryGetValue("담당자 e-Mail", out var em) ? em : issue.담당자이메일;

        BuildItemLines(row);

        if (CheckPanel != null)
            ApplyCheckStates(row);
        else
            Log("WARNING: CheckPanel == null");
    }

    public void Clear()
    {
        _current            = null;
        _cachedRow          = new(StringComparer.OrdinalIgnoreCase);
        txbEmpty.IsVisible  = true;
        spContent.IsVisible = false;
        spButtons.IsVisible = false;
        spItems.Children.Clear();
        txbManagerName.Text  = "";
        txbManagerPhone.Text = "";
        txbManagerEmail.Text = "";
    }

    /// <summary>체크박스 변경 실시간 반영 — 캐시된 행 기반으로 체크된 항목만 표시</summary>
    public void PreviewCheckedItems(IEnumerable<string> checkedAnalyteNames)
    {
        if (_cachedRow.Count == 0) return;
        var checkedSet = new HashSet<string>(checkedAnalyteNames, StringComparer.OrdinalIgnoreCase);
        // 캐시 복사 후 체크 해제된 메인 항목 컬럼값을 빈값으로 설정
        var preview = new Dictionary<string, string>(_cachedRow, StringComparer.OrdinalIgnoreCase);
        foreach (var key in preview.Keys.ToList())
        {
            if (FixedCols.Contains(key)) continue;
            if (key.EndsWith("단가") || key.EndsWith("소계")) continue;
            if (!checkedSet.Contains(key))
                preview[key] = "";  // HasNonZeroStr → false → 행 미표시
        }
        BuildItemLines(preview);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 라인 빌드
    // ══════════════════════════════════════════════════════════════════════
    private void BuildItemLines(Dictionary<string, string> row)
    {
        spItems.Children.Clear();
        bool odd = false;

        foreach (var kv in row)
        {
            var col = kv.Key;   // 이미 Trim된 상태 (GetIssueRow에서 Trim)

            if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
            if (FixedCols.Contains(col)) continue;
            if (!HasNonZeroStr(kv.Value)) continue;

            row.TryGetValue(col + "단가", out var priceStr);
            row.TryGetValue(col + "소계", out var subStr);

            var grid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,55,75,80"),
                Background        = Brush.Parse(odd ? "#1a1a28" : "#1e1e30"),
            };
            odd = !odd;

            grid.Children.Add(Cell(col,             AppFonts.MD,   "#cccccc", 0));
            grid.Children.Add(Cell(kv.Value,        AppFonts.Base, "#aaaaaa", 1, HorizontalAlignment.Right));
            grid.Children.Add(Cell(FmtNum(priceStr),AppFonts.Base, "#aaaaaa", 2, HorizontalAlignment.Right));
            grid.Children.Add(Cell(FmtNum(subStr),  AppFonts.Base, "#88cc88", 3, HorizontalAlignment.Right));

            spItems.Children.Add(grid);
        }

        int cnt = spItems.Children.Count;
        Log($"BuildItemLines → {cnt}행");

        if (cnt == 0)
        {
            spItems.Children.Add(new TextBlock
            {
                Text       = "항목 데이터 없음",
                FontSize   = AppFonts.Base,
                FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Avalonia.Thickness(12, 4),
            });
        }
    }

    private TextBlock Cell(string text, double size, string color,
                           int col, HorizontalAlignment ha = HorizontalAlignment.Left)
    {
        var tb = new TextBlock
        {
            Text                = text ?? "",
            FontSize            = size,
            FontFamily          = Font,
            Foreground          = Brush.Parse(color),
            Margin              = new Avalonia.Thickness(10, 3),
            HorizontalAlignment = ha,
            VerticalAlignment   = VerticalAlignment.Center,
            TextTrimming        = Avalonia.Media.TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(tb, col);
        return tb;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  체크박스 동기화
    // ══════════════════════════════════════════════════════════════════════
    private void ApplyCheckStates(Dictionary<string, string> row)
    {
        var names = CheckPanel!.GetAllAnalyteNames();
        Log($"ApplyCheckStates: CheckPanel항목={names.Count}개");

        int matched = 0, checkedCnt = 0;
        foreach (var name in names)
        {
            // row 키도 Trim됐고, name도 DB에서 온 값이므로 OrdinalIgnoreCase 비교로 충분
            bool has = HasNonZeroStr(GetVal(row, name))
                    || HasNonZeroStr(GetVal(row, name + "단가"))
                    || HasNonZeroStr(GetVal(row, name + "소계"));

            if (row.ContainsKey(name) || row.ContainsKey(name + "단가") || row.ContainsKey(name + "소계"))
                matched++;

            if (has) checkedCnt++;
            CheckPanel.SetChecked(name, has);
        }

        CheckPanel.SyncAllCategories();
        Log($"ApplyCheckStates 완료: 컬럼매칭={matched}  체크됨={checkedCnt}");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static string GetVal(Dictionary<string, string> row, string key)
    {
        if (row.TryGetValue(key, out var v)) return v ?? "";
        // Trim 후 재시도
        var found = row.FirstOrDefault(kv =>
            string.Equals(kv.Key.Trim(), key.Trim(), StringComparison.OrdinalIgnoreCase));
        return found.Value ?? "";
    }

    private static bool HasNonZeroStr(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        // ₩ 기호, 쉼표 제거 후 파싱
        var clean = val.Replace("₩","").Replace(",","").Trim();
        if (decimal.TryParse(clean, NumberStyles.Any,
                             CultureInfo.InvariantCulture, out var d))
            return d != 0;
        return true;
    }

    private static string FmtNum(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return "";
        var clean = val.Replace("₩","").Replace(",","").Trim();
        if (decimal.TryParse(clean, NumberStyles.Any,
                             CultureInfo.InvariantCulture, out var d))
            return d == 0 ? "" : $"{d:#,0}";
        return val;
    }

    // ── 버튼 ─────────────────────────────────────────────────────────────
    // 🥕 당근: 이 건 재활용 (항목 복사, 번호·날짜는 신규 생성)
    private void BtnCarrot_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;
        CarrotRequested?.Invoke(_current);
    }

    // ✏️ 오작성 수정: 시료명·견적번호·발행일·적용구분·업체명 수정
    private void BtnCorrect_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;
        CorrectRequested?.Invoke(_current);
    }

    /// <summary>의뢰서 편집 패널 전환 요청 — MainPage가 구독</summary>
    public event Action<QuotationIssue, List<string>, HashSet<string>>? OrderRequestEditRequested;

    // 📋 의뢰서 작성
    private async void BtnOrderRequest_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;

        // 견적서 분석항목 추출
        var quotedItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var row = QuotationService.GetIssueRow(_current.Id);
            foreach (var kv in row)
            {
                var col = kv.Key;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (FixedCols.Contains(col)) continue;
                if (HasNonZeroStr(kv.Value)) quotedItems.Add(col);
            }
            Log($"[의뢰서] 견적항목 {quotedItems.Count}개");
        }
        catch (Exception ex) { Log($"[의뢰서] 항목 조회 오류: {ex.Message}"); }

        // 시료명 선택 팝업 (이동 가능, 최소/최대/닫기)
        var owner = TopLevel.GetTopLevel(this) as Window;
        var popup = new ETA.Views.OrderRequestWindow(_current, quotedItems);
        if (owner != null) await popup.ShowDialog(owner);
        else popup.Show();

        if (!popup.Confirmed || popup.SelectedSamples.Count == 0) return;

        // MainPage에 편집 패널 전환 요청
        OrderRequestEditRequested?.Invoke(_current, popup.SelectedSamples, quotedItems);
    }

    private void BtnPrint_Click(object? sender, RoutedEventArgs e)
        => Log($"인쇄 → {_current?.견적번호}");

}
