using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

public class AnalysisRequestRecord
{
    public int    Id       { get; set; }
    public string 약칭     { get; set; } = "";
    public string 시료명   { get; set; } = "";
    public string 접수번호 { get; set; } = "";
    public string 의뢰일   { get; set; } = "";
    public string 연도     { get; set; } = "";
    public string 월       { get; set; } = "";
}

public partial class QuotationHistoryPanel : UserControl
{
    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Treeview.log"));
    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── 공개 이벤트 ───────────────────────────────────────────────────────
    public event Action<QuotationIssue>?        IssueSelected;
    public event Action<AnalysisRequestRecord>? AnalysisRequestSelected;
    public event Action?                        AnalysisTabActivated;
    public event Action?                        QuotationTabActivated;

    private bool _isAnalysisTab = false;

    // 검색 디바운스 타이머 (한글 IME 조합 완료 후 검색)
    private readonly DispatcherTimer _searchTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private string _pendingSearchText = "";

    // 검색용 전체 데이터 캐시
    private List<QuotationIssue>          _allIssues  = [];
    private List<AnalysisRequestRecord>   _allAnalysisRecords = [];

    // 분석의뢰 지연 로딩용 캐시
    private Dictionary<string, List<AnalysisRequestRecord>> _analysisByMonth = new();
    private readonly HashSet<string> _loadedMonths = new();

    private readonly TreeView _treeQuotation;
    private readonly TreeView _treeAnalysis;

    public QuotationHistoryPanel()
    {
        InitializeComponent();

        _treeQuotation = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
        };
        _treeQuotation.SelectionChanged += OnQuotationNodeSelected;
        _treeQuotation.KeyDown         += (_, e) => HandleTreeKeyDown(_treeQuotation, e);

        _treeAnalysis = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            IsVisible = false,
            SelectionMode = SelectionMode.Multiple,
        };
        _treeAnalysis.SelectionChanged += OnAnalysisNodeSelected;
        _treeAnalysis.KeyDown          += (_, e) => HandleTreeKeyDown(_treeAnalysis, e);

        treeHost.Children.Add(_treeQuotation);
        treeHost.Children.Add(_treeAnalysis);

        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            ApplySearch(_pendingSearchText);
        };

        Log("QuotationHistoryPanel 초기화 완료");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  [1] 토글 스위치 이벤트
    // ══════════════════════════════════════════════════════════════════════
    private void TglTab_Changed(object? sender, RoutedEventArgs e)
        => SwitchTab(tglTab.IsChecked == true);

    private void SwitchTab(bool toAnalysis)
    {
        _isAnalysisTab           = toAnalysis;
        _treeQuotation.IsVisible = !toAnalysis;
        _treeAnalysis.IsVisible  =  toAnalysis;
        txbTabLabel.Text         = "";   // ToggleSwitch 자체에 레이블 있으므로 빈값
        txbInfo.Text             = "";

        Log($"탭 전환 → {(toAnalysis ? "분석의뢰내역" : "견적발행내역")}");

        if (toAnalysis) { AnalysisTabActivated?.Invoke(); _ = LoadAnalysisTreeAsync(); }
        else            { QuotationTabActivated?.Invoke(); }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  [2][3] 견적발행내역 로드 — 년/월/일 3단계, Expanded 규칙 적용
    // ══════════════════════════════════════════════════════════════════════
    public void LoadData()
    {
        _treeQuotation.Items.Clear();
        Log("LoadData() 시작");

        var issues = QuotationService.GetAllIssues();
        _allIssues = issues;
        Log($"견적 발행건수={issues.Count}");

        var today      = DateTime.Today;
        var thisYear   = today.Year.ToString();
        var thisMonth  = today.Month.ToString("D2");
        var cutoff7    = today.AddDays(-6).ToString("yyyy-MM-dd");   // 최근 7일 기준

        // 년 그룹
        var byYear = issues
            .GroupBy(i => i.발행일.Length >= 4 ? i.발행일[..4] : "기타")
            .OrderByDescending(g => g.Key);

        foreach (var yearGroup in byYear)
        {
            string year      = yearGroup.Key;
            bool   isThisYear = year == thisYear;

            var yearNode = MakeParentNode($"📅  {year}년");
            yearNode.IsExpanded = isThisYear;   // 올해만 펼침

            // 월 그룹
            var byMonth = yearGroup
                .GroupBy(i => i.발행일.Length >= 7 ? i.발행일[5..7] : "??")
                .OrderByDescending(g => g.Key);

            foreach (var monthGroup in byMonth)
            {
                string month       = monthGroup.Key;
                bool   isThisMonth = isThisYear && month == thisMonth;

                var monthNode = MakeParentNode($"  {month}월  ({monthGroup.Count()}건)");
                monthNode.IsExpanded = isThisMonth;  // 올해 이번 달만 펼침

                // 일 그룹
                var byDay = monthGroup
                    .GroupBy(i => i.발행일.Length >= 10 ? i.발행일[8..10] : "??")
                    .OrderByDescending(g => g.Key);

                foreach (var dayGroup in byDay)
                {
                    string day     = dayGroup.Key;
                    string dateStr = $"{year}-{month}-{day}";

                    // 최근 7일 이내 여부
                    bool isRecent = string.Compare(dateStr, cutoff7,
                                       StringComparison.Ordinal) >= 0;

                    var dayNode = MakeParentNode($"    {day}일  ({dayGroup.Count()}건)");
                    dayNode.IsExpanded = isRecent;   // 최근 7일만 펼침

                    foreach (var issue in dayGroup.OrderByDescending(i => i.발행일))
                        dayNode.Items.Add(MakeIssueLeaf(issue));

                    monthNode.Items.Add(dayNode);
                }

                yearNode.Items.Add(monthNode);
            }

            _treeQuotation.Items.Add(yearNode);
        }

        txbInfo.Text = $"총 {issues.Count}건";
        Log($"LoadData() 완료 — 루트노드={_treeQuotation.Items.Count}");
    }

    private void TvHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e) { }

    // ══════════════════════════════════════════════════════════════════════
    //  키보드 상하 이동 — 펼쳐진 리프 노드만 순환
    // ══════════════════════════════════════════════════════════════════════
    private void HandleTreeKeyDown(TreeView tree, KeyEventArgs e)
    {
        if (e.Key != Key.Up && e.Key != Key.Down) return;

        var leaves = CollectVisibleLeaves(tree);
        if (leaves.Count == 0) return;

        var current = tree.SelectedItem as TreeViewItem;
        int idx     = current != null ? leaves.IndexOf(current) : -1;

        int next = e.Key == Key.Down
            ? (idx < 0 ? 0 : Math.Min(idx + 1, leaves.Count - 1))
            : (idx < 0 ? leaves.Count - 1 : Math.Max(idx - 1, 0));

        if (next == idx) { e.Handled = true; return; }

        var target = leaves[next];
        tree.SelectedItem = target;
        target.BringIntoView();
        e.Handled = true;
    }

    /// <summary>현재 펼쳐진 상태 기준으로 리프 노드(QuotationIssue/AnalysisRequestRecord 태그)만 수집</summary>
    private static List<TreeViewItem> CollectVisibleLeaves(ItemsControl parent)
    {
        var result = new List<TreeViewItem>();
        foreach (var obj in parent.Items)
        {
            if (obj is not TreeViewItem node) continue;
            if (node.Tag is QuotationIssue || node.Tag is AnalysisRequestRecord)
                result.Add(node);
            else if (node.IsExpanded)
                result.AddRange(CollectVisibleLeaves(node));
        }
        return result;
    }

    private void TxbSearch_TextChanged(object? sender, TextChangedEventArgs e)
    {
        _pendingSearchText = sender is TextBox tb ? tb.Text ?? "" : "";
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  검색 필터 — 키워드 있으면 평면 결과, 없으면 원래 트리 복원
    // ══════════════════════════════════════════════════════════════════════
    private void ApplySearch(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword))
        {
            // 검색어 없음 → 원래 트리 복원
            if (_isAnalysisTab) _ = LoadAnalysisTreeAsync();
            else                LoadData();
            return;
        }

        var kw = keyword.Trim();

        if (_isAnalysisTab)
        {
            var matched = _allAnalysisRecords
                .Where(r =>
                    r.약칭    .Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    r.시료명  .Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    r.접수번호.Contains(kw, StringComparison.OrdinalIgnoreCase))
                .ToList();

            _treeAnalysis.Items.Clear();
            foreach (var rec in matched)
                _treeAnalysis.Items.Add(MakeAnalysisLeaf(rec));

            txbInfo.Text = $"검색 결과 {matched.Count}건";
        }
        else
        {
            var matched = _allIssues
                .Where(i =>
                    (i.약칭     ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (i.시료명   ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (i.견적번호 ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase))
                .ToList();

            _treeQuotation.Items.Clear();
            foreach (var issue in matched)
                _treeQuotation.Items.Add(MakeIssueLeaf(issue));

            txbInfo.Text = $"검색 결과 {matched.Count}건";
        }
    }

    private void OnQuotationNodeSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (_treeQuotation.SelectedItem is TreeViewItem item &&
            item.Tag is QuotationIssue issue)
        {
            txbInfo.Text = $"{issue.약칭}  {issue.시료명}  |  {issue.견적번호}";
            Log($"견적 선택: {issue.약칭} {issue.시료명} [{issue.견적번호}]");
            IssueSelected?.Invoke(issue);
        }
    }

    private async void BtnDelete_Click(object? sender, RoutedEventArgs e)
    {
        var owner = TopLevel.GetTopLevel(this) as Window;

        // ── 분석의뢰내역 탭 삭제 (복수 선택 지원) ──────────────────────────
        if (_isAnalysisTab)
        {
            var recs = GetSelectedAnalysisRecords();
            if (recs.Count == 0) return;

            string msg = recs.Count == 1
                ? $"아래 분석의뢰를 DB에서 삭제하시겠습니까?\n\n{recs[0].약칭}  {recs[0].시료명}\n{recs[0].접수번호}"
                : $"선택된 {recs.Count}건을 DB에서 삭제하시겠습니까?\n\n"
                  + string.Join("\n", recs.Take(5).Select(r => $"  {r.약칭}  {r.시료명}"))
                  + (recs.Count > 5 ? $"\n  … 외 {recs.Count - 5}건" : "");

            bool confirmed = await ShowConfirmDialogAsync(owner, msg);
            if (!confirmed) return;

            int deleted = 0;
            await Task.Run(() =>
            {
                foreach (var r in recs)
                {
                    Log($"분석의뢰 삭제: {r.접수번호} id={r.Id}");
                    if (AnalysisRequestService.DeleteRecord(r.Id)) deleted++;
                }
            });

            _loadedMonths.Clear();
            _analysisByMonth.Clear();
            await LoadAnalysisTreeAsync();
            await ShowAlertDialogAsync(owner, $"{deleted}건 삭제되었습니다.");
            return;
        }

        // ── 견적발행내역 탭 삭제 ─────────────────────────────────────────
        if (!(_treeQuotation.SelectedItem is TreeViewItem qItem &&
              qItem.Tag is QuotationIssue issue)) return;

        bool qConfirmed = await ShowConfirmDialogAsync(owner,
            $"아래 견적을 삭제하시겠습니까?\n\n{issue.약칭}  {issue.시료명}\n{issue.견적번호}");
        if (!qConfirmed) return;

        Log($"삭제: {issue.견적번호}");
        bool qOk = QuotationService.Delete(issue.Id);
        if (qOk)
        {
            LoadData();
            await ShowAlertDialogAsync(owner, "삭제되었습니다.");
        }
    }

    // ── 다이얼로그 헬퍼 ──────────────────────────────────────────────────
    private static Task<bool> ShowConfirmDialogAsync(Window? owner, string message)
    {
        var tcs = new System.Threading.Tasks.TaskCompletionSource<bool>();

        var dlg = new Window
        {
            Title                 = "삭제 확인",
            Width                 = 340,
            SizeToContent         = SizeToContent.Height,
            CanResize             = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background            = new SolidColorBrush(Color.Parse("#1e1e2e")),
        };

        var yesBtn = new Button
        {
            Content             = "삭제",
            Width               = 70,
            Background          = new SolidColorBrush(Color.Parse("#4a1a1a")),
            Foreground          = new SolidColorBrush(Color.Parse("#ff8888")),
            FontFamily          = Font,
            FontSize            = AppFonts.Base,
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        var noBtn = new Button
        {
            Content             = "취소",
            Width               = 70,
            Background          = new SolidColorBrush(Color.Parse("#2a2a3a")),
            Foreground          = new SolidColorBrush(Color.Parse("#aaaaaa")),
            FontFamily          = Font,
            FontSize            = AppFonts.Base,
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };

        yesBtn.Click += (_, _) => { tcs.TrySetResult(true);  dlg.Close(); };
        noBtn.Click  += (_, _) => { tcs.TrySetResult(false); dlg.Close(); };
        dlg.Closed   += (_, _) => tcs.TrySetResult(false);

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20, 16, 20, 16),
            Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text         = message,
                    Foreground   = new SolidColorBrush(Color.Parse("#dddddd")),
                    FontSize     = AppFonts.MD,
                    FontFamily   = Font,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                new StackPanel
                {
                    Orientation         = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing             = 8,
                    Children = { noBtn, yesBtn },
                }
            }
        };

        if (owner != null) dlg.ShowDialog(owner);
        else dlg.Show();

        return tcs.Task;
    }

    private static Task ShowAlertDialogAsync(Window? owner, string message)
    {
        var tcs = new System.Threading.Tasks.TaskCompletionSource<bool>();

        var dlg = new Window
        {
            Title                 = "알림",
            Width                 = 280,
            SizeToContent         = SizeToContent.Height,
            CanResize             = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background            = new SolidColorBrush(Color.Parse("#1e1e2e")),
        };

        var okBtn = new Button
        {
            Content             = "확인",
            Width               = 70,
            Background          = new SolidColorBrush(Color.Parse("#2a4a2a")),
            Foreground          = new SolidColorBrush(Color.Parse("#aef0ae")),
            FontFamily          = Font,
            FontSize            = AppFonts.Base,
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };

        okBtn.Click += (_, _) => dlg.Close();
        dlg.Closed  += (_, _) => tcs.TrySetResult(true);

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20, 16, 20, 16),
            Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text       = message,
                    Foreground = new SolidColorBrush(Color.Parse("#dddddd")),
                    FontSize   = AppFonts.MD,
                    FontFamily = Font,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                new StackPanel
                {
                    Orientation         = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children = { okBtn },
                }
            }
        };

        if (owner != null) dlg.ShowDialog(owner);
        else dlg.Show();

        return tcs.Task;
    }

    private void BtnRefresh_Click(object? sender, RoutedEventArgs e)
    {
        Log("새로고침");
        if (_isAnalysisTab) _ = LoadAnalysisTreeAsync();
        else LoadData();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  분석의뢰내역 — 지연 로딩 (년/월/일 구조)
    // ══════════════════════════════════════════════════════════════════════
    private async Task LoadAnalysisTreeAsync()
    {
        _treeAnalysis.Items.Clear();
        _loadedMonths.Clear();
        _analysisByMonth.Clear();
        txbInfo.Text = "로딩 중...";
        Log("LoadAnalysisTreeAsync() 시작");

        List<AnalysisRequestRecord> records;
        try
        {
            records = await Task.Run(() => AnalysisRequestService.GetAllRecords());
            Log($"DB 조회 완료: {records.Count}건");
        }
        catch (Exception ex)
        {
            Log($"DB 오류: {ex.Message}"); txbInfo.Text = "로드 실패"; return;
        }

        if (records.Count == 0) { txbInfo.Text = "분석의뢰 데이터 없음"; return; }
        _allAnalysisRecords = records;

        var today     = DateTime.Today;
        var thisYear  = today.Year.ToString();
        var thisMonth = today.Month.ToString("D2");
        var cutoff7   = today.AddDays(-6).ToString("yyyy-MM-dd");

        _analysisByMonth = records
            .GroupBy(r => $"{r.연도}-{r.월}")
            .ToDictionary(g => g.Key, g => g.OrderByDescending(r => r.의뢰일).ToList());

        var byYear = records.GroupBy(r => r.연도).OrderByDescending(g => g.Key).ToList();

        bool isFirstMonth = true;

        foreach (var yearGroup in byYear)
        {
            string year       = yearGroup.Key;
            bool   isThisYear = year == thisYear;

            var yearNode = MakeParentNode($"📅  {year}년");
            yearNode.IsExpanded = isThisYear;

            var byMonth = yearGroup.GroupBy(r => r.월).OrderByDescending(g => g.Key).ToList();

            foreach (var monthGroup in byMonth)
            {
                string monthKey    = $"{year}-{monthGroup.Key}";
                int    count       = _analysisByMonth.TryGetValue(monthKey, out var ml)
                                     ? ml.Count : monthGroup.Count();
                bool   isThisMonth = isThisYear && monthGroup.Key == thisMonth;

                var monthNode = MakeParentNode($"  {monthGroup.Key}월  ({count}건)");
                monthNode.Tag         = monthKey;
                monthNode.IsExpanded  = isThisMonth;

                if (isFirstMonth)
                {
                    FillMonthNodeWithDays(monthNode, monthKey, cutoff7);
                    isFirstMonth = false;
                }
                else
                {
                    // 플레이스홀더 + 지연 로딩
                    monthNode.Items.Add(new TreeViewItem
                    {
                        Header = Fs(new TextBlock
                        {
                            FontFamily = Font, Text = "...",
                            Foreground = Brush.Parse("#444455"),
                        }, "FontSizeXS"),
                        Tag = "__placeholder__",
                    });
                    var capturedKey  = monthKey;
                    var capturedNode = monthNode;
                    IDisposable? sub = null;
                    sub = monthNode.GetObservable(TreeViewItem.IsExpandedProperty)
                        .Subscribe(expanded =>
                        {
                            if (!expanded) return;
                            if (_loadedMonths.Contains(capturedKey)) return;
                            FillMonthNodeWithDays(capturedNode, capturedKey, cutoff7);
                            sub?.Dispose();
                        });
                }

                yearNode.Items.Add(monthNode);
            }

            _treeAnalysis.Items.Add(yearNode);
        }

        txbInfo.Text = $"총 {records.Count}건";
        Log($"트리 완료 — 년노드={_treeAnalysis.Items.Count}");
    }

    // 월 노드 안에 일 노드 + 리프 생성
    private void FillMonthNodeWithDays(TreeViewItem monthNode, string monthKey, string cutoff7)
    {
        monthNode.Items.Clear();
        _loadedMonths.Add(monthKey);

        if (!_analysisByMonth.TryGetValue(monthKey, out var recs)) return;

        var byDay = recs
            .GroupBy(r => r.의뢰일.Length >= 10 ? r.의뢰일[8..10] : "??")
            .OrderByDescending(g => g.Key);

        // monthKey = "YYYY-MM"
        string yearMonth = monthKey;   // e.g. "2026-03"

        foreach (var dayGroup in byDay)
        {
            string day     = dayGroup.Key;
            string dateStr = $"{yearMonth}-{day}";
            bool   isRecent = string.Compare(dateStr, cutoff7,
                                  StringComparison.Ordinal) >= 0;

            var dayNode = MakeParentNode($"    {day}일  ({dayGroup.Count()}건)");
            dayNode.IsExpanded = isRecent;

            foreach (var rec in dayGroup)
                dayNode.Items.Add(MakeAnalysisLeaf(rec));

            monthNode.Items.Add(dayNode);
        }

        Log($"FillMonthNodeWithDays: {monthKey} → {recs.Count}건");
    }

    private void OnAnalysisNodeSelected(object? sender, SelectionChangedEventArgs e)
    {
        var selected = GetSelectedAnalysisRecords();

        if (selected.Count == 1)
        {
            var rec = selected[0];
            txbInfo.Text = $"{rec.약칭}  {rec.시료명}  |  {rec.접수번호}";
            Log($"의뢰 선택: {rec.약칭} {rec.시료명} [{rec.접수번호}]");
            AnalysisRequestSelected?.Invoke(rec);
        }
        else if (selected.Count > 1)
        {
            txbInfo.Text = $"{selected.Count}건 선택됨";
        }
    }

    private List<AnalysisRequestRecord> GetSelectedAnalysisRecords()
        => _treeAnalysis.SelectedItems
            .OfType<TreeViewItem>()
            .Where(item => item.Tag is AnalysisRequestRecord)
            .Select(item => (AnalysisRequestRecord)item.Tag!)
            .ToList();

    // ══════════════════════════════════════════════════════════════════════
    //  노드 헬퍼
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>TextBlock 의 FontSize 를 DynamicResource 에 바인딩한다 (슬라이더 실시간 반영)</summary>
    private TextBlock Fs(TextBlock tb, string resourceKey)
    {
        tb.Bind(TextBlock.FontSizeProperty, this.GetResourceObservable(resourceKey));
        return tb;
    }

    private TreeViewItem MakeParentNode(string header)
    {
        var tb = new TextBlock
        {
            Text = header, FontFamily = Font,
            Foreground = Brush.Parse("#aaaacc"),
            Margin = new Thickness(2, 1),
        };
        Fs(tb, "FontSizeBase");
        return new TreeViewItem { Header = tb, IsExpanded = false };
    }

    private TreeViewItem MakeIssueLeaf(QuotationIssue issue)
    {
        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(issue.약칭);
        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };
        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
        topRow.Children.Add(new Border
        {
            Background = Brush.Parse(ibg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 0,
            Child = Fs(new TextBlock { Text = issue.약칭, FontFamily = Font,
                                    Foreground = Brush.Parse(ifg) }, "FontSizeXS"),
        });
        topRow.Children.Add(Fs(new TextBlock
        {
            Text = issue.시료명, FontFamily = Font,
            Foreground = Brush.Parse("#dddddd"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        }, "FontSizeBase"));
        sp.Children.Add(topRow);
        var bottomRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = issue.견적번호, FontFamily = Font,
            Foreground = Brush.Parse("#445566"),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        }, "FontSizeXS"));
        if (!string.IsNullOrEmpty(issue.담당자))
            bottomRow.Children.Add(Fs(new TextBlock
            {
                Text = issue.담당자, FontFamily = Font,
                Foreground = Brush.Parse("#778899"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0),
                [Grid.ColumnProperty] = 1,
            }, "FontSizeXS"));
        sp.Children.Add(bottomRow);
        return new TreeViewItem { Header = sp, Tag = issue };
    }

    private TreeViewItem MakeAnalysisLeaf(AnalysisRequestRecord rec)
    {
        var (rbg, rfg) = BadgeColorHelper.GetBadgeColor(rec.약칭);
        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };
        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
        topRow.Children.Add(new Border
        {
            Background = Brush.Parse(rbg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 0,
            Child = Fs(new TextBlock { Text = rec.약칭, FontFamily = Font,
                                    Foreground = Brush.Parse(rfg) }, "FontSizeXS"),
        });
        topRow.Children.Add(Fs(new TextBlock
        {
            Text = rec.시료명, FontFamily = Font,
            Foreground = Brush.Parse("#dddddd"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        }, "FontSizeBase"));
        sp.Children.Add(topRow);
        sp.Children.Add(Fs(new TextBlock
        {
            Text = rec.접수번호, FontFamily = Font,
            Foreground = Brush.Parse("#445566"), Margin = new Thickness(0, 0, 0, 1),
        }, "FontSizeXS"));

        var node = new TreeViewItem { Header = sp, Tag = rec };

        // 더블탭 → 삭제 확인
        node.DoubleTapped += async (_, _) =>
        {
            var owner = TopLevel.GetTopLevel(node) as Window;
            bool confirmed = await ShowConfirmDialogAsync(owner,
                $"아래 분석의뢰를 DB에서 삭제하시겠습니까?\n\n{rec.약칭}  {rec.시료명}\n{rec.접수번호}");
            if (!confirmed) return;
            bool ok = await Task.Run(() => AnalysisRequestService.DeleteRecord(rec.Id));
            if (ok)
            {
                _loadedMonths.Clear();
                _analysisByMonth.Clear();
                await LoadAnalysisTreeAsync();
                await ShowAlertDialogAsync(owner, "삭제되었습니다.");
            }
        };

        return node;
    }
}
