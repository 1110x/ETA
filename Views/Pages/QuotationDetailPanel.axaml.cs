using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

public partial class QuotationDetailPanel : UserControl
{
    private QuotationIssue?      _current;
    public  QuotationCheckPanel?  CheckPanel    { get; set; }
    /// <summary>🥕 당근 — 이 건을 재활용해서 신규 작성 (항목 복사, 번호·날짜 신규)</summary>
    public event Action<QuotationIssue>? CarrotRequested;
    /// <summary>✏️ 오작성 수정 — 시료명·견적번호·발행일·적용구분·업체명 등 메타 수정</summary>
    public event Action<QuotationIssue>? CorrectRequested;

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    // ── 고정 컬럼 (항목 순회 제외) — Trim 된 실제 컬럼명 기준 ────────────
    private static readonly HashSet<string> FixedCols =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자연락처","담당자 e-Mail","시료명","견적번호",
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

        BuildItemLines(row);

        if (CheckPanel != null)
            ApplyCheckStates(row);
        else
            Log("WARNING: CheckPanel == null");
    }

    public void Clear()
    {
        _current            = null;
        txbEmpty.IsVisible  = true;
        spContent.IsVisible = false;
        spButtons.IsVisible = false;
        spItems.Children.Clear();
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

            grid.Children.Add(Cell(col,             12, "#cccccc", 0));
            grid.Children.Add(Cell(kv.Value,        11, "#aaaaaa", 1, HorizontalAlignment.Right));
            grid.Children.Add(Cell(FmtNum(priceStr),11, "#aaaaaa", 2, HorizontalAlignment.Right));
            grid.Children.Add(Cell(FmtNum(subStr),  11, "#88cc88", 3, HorizontalAlignment.Right));

            spItems.Children.Add(grid);
        }

        int cnt = spItems.Children.Count;
        Log($"BuildItemLines → {cnt}행");

        if (cnt == 0)
        {
            spItems.Children.Add(new TextBlock
            {
                Text       = "항목 데이터 없음",
                FontSize   = 11,
                FontFamily = Font,
                Foreground = Brush.Parse("#555"),
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

    // 📊 분석기록부 출력
    private async void BtnAnalysisRecord_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;
        if (sender is Button btn) btn.IsEnabled = false;

        try
        {
            Log($"분석기록부 출력 시작: {_current.견적번호}");

            // 1. 항목 그룹 조회 (DB만 읽기, 빠름)
            var groups = await Task.Run(
                () => AnalysisRecordService.GetAnalysisGroups(_current.견적번호));

            if (groups.Count == 0)
            {
                var dlg = new Window
                {
                    Title                 = "알림",
                    Width                 = 320,
                    Height                = 130,
                    CanResize             = false,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Background            = new SolidColorBrush(Color.Parse("#2d2d2d")),
                };
                var okBtn2 = new Button
                {
                    Content             = "확인",
                    Width               = 70,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Background          = new SolidColorBrush(Color.Parse("#3a3a4a")),
                    Foreground          = Brushes.WhiteSmoke,
                };
                okBtn2.Click += (_, _) => dlg.Close();
                dlg.Content = new StackPanel
                {
                    Margin = new Thickness(20), Spacing = 12,
                    Children =
                    {
                        new TextBlock
                        {
                            Text         = "생성할 분석기록부가 없습니다.\n의뢰서를 먼저 작성해 주세요.",
                            Foreground   = Brushes.WhiteSmoke, FontSize = 12,
                            TextWrapping = Avalonia.Media.TextWrapping.Wrap, FontFamily = Font,
                        },
                        okBtn2
                    }
                };
                var owner0 = TopLevel.GetTopLevel(this) as Window;
                if (owner0 != null) await dlg.ShowDialog(owner0); else dlg.Show();
                return;
            }

            // 2. 프로그래스 창 빌드
            var progressBars  = new Dictionary<string, ProgressBar>(StringComparer.OrdinalIgnoreCase);
            var statusBlocks  = new Dictionary<string, TextBlock>(StringComparer.OrdinalIgnoreCase);
            var itemStack     = new StackPanel { Spacing = 8, Margin = new Thickness(16, 12, 16, 8) };

            foreach (var colName in groups.Keys)
            {
                var pb = new ProgressBar
                {
                    Minimum         = 0,
                    Maximum         = 100,
                    Value           = 0,
                    Height          = 14,
                    Foreground      = new SolidColorBrush(Color.Parse("#4aaa4a")),
                    Background      = new SolidColorBrush(Color.Parse("#2a2a3a")),
                    BorderThickness = new Thickness(0),
                };
                var statusTb = new TextBlock
                {
                    Text              = "대기",
                    FontSize          = 10,
                    Width             = 36,
                    Foreground        = new SolidColorBrush(Color.Parse("#666")),
                    VerticalAlignment = VerticalAlignment.Center,
                    FontFamily        = Font,
                };
                progressBars[colName] = pb;
                statusBlocks[colName] = statusTb;

                var row = new Grid { ColumnDefinitions = new ColumnDefinitions("110,*,40") };
                row.Children.Add(new TextBlock
                {
                    Text              = colName,
                    FontSize          = 11,
                    Foreground        = Brushes.WhiteSmoke,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontFamily        = Font,
                    TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                    [Grid.ColumnProperty] = 0,
                });
                pb[Grid.ColumnProperty] = 1;
                row.Children.Add(pb);
                statusTb[Grid.ColumnProperty] = 2;
                row.Children.Add(statusTb);

                itemStack.Children.Add(row);
            }

            var countTb = new TextBlock
            {
                Text       = $"0 / {groups.Count} 완료",
                FontSize   = 10,
                Foreground = new SolidColorBrush(Color.Parse("#888")),
                FontFamily = Font,
                Margin     = new Thickness(16, 0, 16, 12),
            };

            var progressWin = new Window
            {
                Title                 = "📊 분석기록부 생성 중...",
                Width                 = 440,
                SizeToContent         = SizeToContent.Height,
                CanResize             = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background            = new SolidColorBrush(Color.Parse("#1e1e2e")),
                Content               = new StackPanel
                {
                    Children = { itemStack, countTb }
                }
            };

            var owner = TopLevel.GetTopLevel(this) as Window;
            if (owner != null) progressWin.Show(owner); else progressWin.Show();

            // 3. 파일 생성 (분析자별)
            var outputDir  = AnalysisRecordService.GetOutputDir();
            var 방류기준표 = await Task.Run(() => AnalysisRecordService.Load방류기준표());

            // 모든 항목 진행 바를 indeterminate로 전환
            foreach (var pb2 in progressBars.Values) pb2.IsIndeterminate = true;
            foreach (var stb in statusBlocks.Values) { stb.Text = "생성"; stb.Foreground = new SolidColorBrush(Color.Parse("#aaa")); }

            var results = await Task.Run(() =>
                AnalysisRecordService.GenerateByAssignee(groups, 방류기준표, outputDir, DateTime.Today));

            bool anySuccess = results.Count > 0;
            foreach (var colName in groups.Keys)
            {
                if (progressBars.TryGetValue(colName, out var pb3)) { pb3.IsIndeterminate = false; pb3.Value = anySuccess ? 100 : 0; }
                if (statusBlocks.TryGetValue(colName, out var stb2))
                {
                    stb2.Text = anySuccess ? "✅" : "❌";
                    stb2.Foreground = new SolidColorBrush(Color.Parse(anySuccess ? "#4aaa4a" : "#cc4444"));
                }
            }
            countTb.Text = anySuccess ? $"{results.Count}개 파일 생성 완료" : "생성 실패";

            // 4. 완료 후 창 닫고 폴더 열기
            await Task.Delay(800);
            progressWin.Close();
            if (anySuccess)
                Process.Start(new ProcessStartInfo { FileName = outputDir, UseShellExecute = true });

            Log($"분析기록부 {results.Count}개 생성 완료 → {outputDir}");
        }
        finally
        {
            if (sender is Button b) b.IsEnabled = true;
        }
    }

    private void BtnDelete_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;
        if (QuotationService.Delete(_current.Id))
        { Log($"삭제 완료 → {_current.견적번호}"); Clear(); }
    }
}
