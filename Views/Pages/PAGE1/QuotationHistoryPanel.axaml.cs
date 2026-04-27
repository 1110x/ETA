using Avalonia;
using ETA.Views;
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
    public string 채취일자   { get; set; } = "";
    public string 정도보증   { get; set; } = "";
    public string 채수담당자 { get; set; } = "";
}

public class TradeStatementRecord
{
    public int     Id             { get; set; }
    public string  발행일         { get; set; } = "";
    public string  업체명         { get; set; } = "";
    public string  약칭           { get; set; } = "";
    public string  거래명세서번호 { get; set; } = "";
    public decimal 공급가액       { get; set; }
    public decimal 부가세         { get; set; }
    public decimal 합계금액       { get; set; }
    public string  연도 { get; set; } = "";
    public string  월   { get; set; } = "";
}

public partial class QuotationHistoryPanel : UserControl
{
    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath =
        Path.Combine(ETA.Services.Common.AppPaths.LogsDir, "Treeview.log");
    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 공개 이벤트 ───────────────────────────────────────────────────────
    public event Action<QuotationIssue>?        IssueSelected;
    public event Action<QuotationIssue>?        IssueAddedToList;   // 발행 모드: 노드 클릭 시
    public event Action<AnalysisRequestRecord>? AnalysisRequestSelected;
    public event Action<AnalysisRequestRecord>? AnalysisRequestAddedToList; // 발행 모드: 의뢰 노드 클릭 시
    public event Action<TradeStatementRecord>? TradeStatementSelected;
    public event Action?                        AnalysisTabActivated;
    public event Action?                        QuotationTabActivated;
    public event Action?                        StatementTabActivated;

    /// <summary>true이면 노드 클릭 시 IssueAddedToList 이벤트를 발생시키고 Show2 상세표시 안 함.
    /// false=견적/의뢰서(보라색), true=거래명세서 발행(청색)</summary>
    private bool _issuingMode = false;
    public bool IssuingMode
    {
        get => _issuingMode;
        set { _issuingMode = value; ApplyTheme(); }
    }

    private void ApplyTheme()
    {
        bool blue = _issuingMode;
        headerBorder.Background    = Brush.Parse(blue ? "#181e28" : "#1e1828");
        separatorBorder.Background = Brush.Parse(blue ? "#2a3a5a" : "#3a2a5a");
        treeHostBorder.Background  = Brush.Parse(blue ? "#0e1322" : "#130e1e");
        this.Resources["TreeItemSelectedBg"] = Brush.Parse(blue ? "#1a2a4a" : "#2a1a4a");

        // 거래명세발행내역 탭은 발행모드에서만 노출
        if (btnTabStatement != null)
            btnTabStatement.IsVisible = blue;
    }

    private enum TabKind { Quotation, Analysis, Statement }
    private TabKind _currentTab = TabKind.Quotation;
    private bool _isAnalysisTab => _currentTab == TabKind.Analysis;

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
    private readonly TreeView _treeStatement;

    public QuotationHistoryPanel()
    {
        // DynamicResource 초기값 설정 (InitializeComponent 전에 필요)
        this.Resources["TreeItemSelectedBg"] = Brush.Parse("#2a1a4a");
        InitializeComponent();
        ApplyTheme();

        _treeQuotation = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            SelectionMode = SelectionMode.Multiple,  // Ctrl+클릭 다중선택
            ClipToBounds = false,
        };
        _treeQuotation.SelectionChanged += OnQuotationNodeSelected;
        _treeQuotation.KeyDown         += (_, e) => HandleTreeKeyDown(_treeQuotation, e);

        _treeAnalysis = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            IsVisible = false,
            SelectionMode = SelectionMode.Multiple,
            ClipToBounds = false,
        };
        _treeAnalysis.SelectionChanged += OnAnalysisNodeSelected;
        _treeAnalysis.KeyDown          += (_, e) => HandleTreeKeyDown(_treeAnalysis, e);

        _treeStatement = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            IsVisible = false,
            SelectionMode = SelectionMode.Single,
            ClipToBounds = false,
        };
        _treeStatement.KeyDown += (_, e) => HandleTreeKeyDown(_treeStatement, e);
        _treeStatement.SelectionChanged += OnStatementNodeSelected;

        treeHost.Children.Add(_treeQuotation);
        treeHost.Children.Add(_treeAnalysis);
        treeHost.Children.Add(_treeStatement);

        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            ApplySearch(_pendingSearchText);
        };

        Log("QuotationHistoryPanel 초기화 완료");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  [1] 탭 버튼 이벤트
    // ══════════════════════════════════════════════════════════════════════
    private void BtnTabQuotation_Click(object? sender, RoutedEventArgs e) => SwitchTab(TabKind.Quotation);
    private void BtnTabAnalysis_Click(object? sender, RoutedEventArgs e)  => SwitchTab(TabKind.Analysis);
    private void BtnTabStatement_Click(object? sender, RoutedEventArgs e) => SwitchTab(TabKind.Statement);

    private void SwitchTab(TabKind kind)
    {
        _currentTab = kind;
        _treeQuotation.IsVisible = kind == TabKind.Quotation;
        _treeAnalysis.IsVisible  = kind == TabKind.Analysis;
        if (_treeStatement != null) _treeStatement.IsVisible = kind == TabKind.Statement;

        UpdateTabButtonStyles();
        txbTabLabel.Text = "";
        txbInfo.Text     = "";

        Log($"탭 전환 → {kind}");

        switch (kind)
        {
            case TabKind.Quotation: QuotationTabActivated?.Invoke(); break;
            case TabKind.Analysis:  AnalysisTabActivated?.Invoke();  _ = LoadAnalysisTreeAsync(); break;
            case TabKind.Statement: StatementTabActivated?.Invoke(); _ = LoadStatementTreeAsync(); break;
        }
    }

    private void UpdateTabButtonStyles()
    {
        void Set(Button b, bool active)
        {
            if (active) { if (!b.Classes.Contains("active")) b.Classes.Add("active"); }
            else        { b.Classes.Remove("active"); }
        }
        Set(btnTabQuotation, _currentTab == TabKind.Quotation);
        Set(btnTabAnalysis,  _currentTab == TabKind.Analysis);
        Set(btnTabStatement, _currentTab == TabKind.Statement);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  [2][3] 견적발행내역 로드 — 년/월 2단계, 리프에 (MM/DD) 표시
    // ══════════════════════════════════════════════════════════════════════
    public void LoadData()
    {
        _treeQuotation.Items.Clear();
        Log("LoadData() 시작");

        var issues = QuotationService.GetAllIssues();
        _allIssues = issues;
        Log($"견적 발행건수={issues.Count}");

        var today     = DateTime.Today;
        var thisYear  = today.Year.ToString();
        var thisMonth = today.Month.ToString("D2");

        var byYear = issues
            .GroupBy(i => i.발행일.Length >= 4 ? i.발행일[..4] : "기타")
            .OrderByDescending(g => g.Key);

        foreach (var yearGroup in byYear)
        {
            string year       = yearGroup.Key;
            bool   isThisYear = year == thisYear;

            var yearNode = MakeParentNode($"📅  {year}년");
            yearNode.IsExpanded = isThisYear;

            var byMonth = yearGroup
                .GroupBy(i => i.발행일.Length >= 7 ? i.발행일[5..7] : "??")
                .OrderByDescending(g => g.Key);

            foreach (var monthGroup in byMonth)
            {
                string month       = monthGroup.Key;
                bool   isThisMonth = isThisYear && month == thisMonth;

                var monthNode = MakeParentNode($"  {month}월  ({monthGroup.Count()}건)");
                monthNode.IsExpanded = isThisMonth;

                foreach (var issue in monthGroup.OrderByDescending(i => i.발행일))
                    monthNode.Items.Add(MakeIssueLeaf(issue));

                yearNode.Items.Add(monthNode);
            }

            _treeQuotation.Items.Add(yearNode);
        }

        txbInfo.Text = $"총 {issues.Count}건";
        Log($"LoadData() 완료 — 루트노드={_treeQuotation.Items.Count}");
    }

    private void TvHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e) { }

    /// <summary>현재 Ctrl+클릭으로 선택된 QuotationIssue 목록 반환</summary>
    public List<QuotationIssue> GetSelectedIssues()
    {
        var result = new List<QuotationIssue>();
        foreach (var item in _treeQuotation.SelectedItems)
        {
            if (item is TreeViewItem tvi && tvi.Tag is QuotationIssue issue)
                result.Add(issue);
        }
        return result;
    }

    /// <summary>특정 견적 목록의 트리 아이콘을 파란색으로 갱신</summary>
    public void RefreshIssueIcons(HashSet<int> issuedIds)
    {
        RefreshIconsInTree(_treeQuotation, issuedIds);
    }

    private static void RefreshIconsInTree(ItemsControl parent, HashSet<int> issuedIds)
    {
        foreach (var obj in parent.Items)
        {
            if (obj is not TreeViewItem node) continue;
            if (node.Tag is QuotationIssue issue)
            {
                // 아이콘 TextBlock은 topRow의 첫 번째 자식
                if (node.Header is StackPanel sp &&
                    sp.Children.Count > 0 &&
                    sp.Children[0] is Grid topRow &&
                    topRow.Children.Count > 0 &&
                    topRow.Children[0] is TextBlock iconBlock)
                {
                    iconBlock.Text = issuedIds.Contains(issue.Id) ? "🔵" : "🔴";
                }
            }
            else
            {
                RefreshIconsInTree(node, issuedIds);
            }
        }
    }

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
            if (IssuingMode) IssueAddedToList?.Invoke(issue);
            else             IssueSelected?.Invoke(issue);
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
        var selected = GetSelectedIssues();
        if (selected.Count == 0) return;

        // 거래명세서번호가 있는 건이 포함된 경우 → 거래명세서 삭제 옵션 제공
        var withStatement = selected.Where(i => !string.IsNullOrEmpty(i.거래명세서번호)).ToList();
        if (withStatement.Count > 0)
        {
            var stNos = withStatement.Select(i => i.거래명세서번호!).Distinct().ToList();
            string stMsg = stNos.Count == 1
                ? $"거래명세서 [{stNos[0]}]을(를) 삭제하시겠습니까?\n\n연결된 견적 {withStatement.Count}건의 거래명세서번호가 초기화됩니다."
                : $"거래명세서 {stNos.Count}건을 삭제하시겠습니까?\n\n{string.Join("\n", stNos.Take(5))}"
                  + (stNos.Count > 5 ? $"\n… 외 {stNos.Count - 5}건" : "");

            bool stConfirmed = await ShowConfirmDialogAsync(owner, stMsg);
            if (!stConfirmed) return;

            await Task.Run(() =>
            {
                foreach (var no in stNos)
                    QuotationService.DeleteTradeStatement(no);
            });

            LoadData();
            await ShowAlertDialogAsync(owner, $"거래명세서 {stNos.Count}건 삭제되었습니다.");
            return;
        }

        // 거래명세서번호 없는 일반 견적 삭제
        string qMsg = selected.Count == 1
            ? $"아래 견적을 삭제하시겠습니까?\n\n{selected[0].약칭}  {selected[0].시료명}\n{selected[0].견적번호}"
            : $"선택된 {selected.Count}건을 삭제하시겠습니까?\n\n"
              + string.Join("\n", selected.Take(5).Select(i => $"  {i.약칭}  {i.시료명}"))
              + (selected.Count > 5 ? $"\n  … 외 {selected.Count - 5}건" : "");

        bool qConfirmed = await ShowConfirmDialogAsync(owner, qMsg);
        if (!qConfirmed) return;

        int qDeleted = 0;
        await Task.Run(() =>
        {
            foreach (var issue in selected)
            {
                Log($"삭제: {issue.견적번호}");
                if (QuotationService.Delete(issue.Id)) qDeleted++;
            }
        });

        LoadData();
        await ShowAlertDialogAsync(owner, $"{qDeleted}건 삭제되었습니다.");
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
            Background            = AppTheme.BgPrimary,
        };

        var yesBtn = new Button
        {
            Content             = "삭제",
            Width               = 70,
            Background          = new SolidColorBrush(Color.Parse("#4a1a1a")),
            Foreground          = AppTheme.FgDanger,
            FontFamily          = Font,
            FontSize            = AppFonts.Base,
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        var noBtn = new Button
        {
            Content             = "취소",
            Width               = 70,
            Background          = AppTheme.BgCard,
            Foreground          = AppTheme.FgMuted,
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
                    Foreground   = AppTheme.FgPrimary,
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
            Background            = AppTheme.BgPrimary,
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
                    Foreground = AppTheme.FgPrimary,
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
    //  거래명세발행내역 — 년/월 2단계, 리프에 (MM/DD) 표시
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

        if (records.Count == 0) { txbInfo.Text = "데이터 없음"; return; }
        _allAnalysisRecords = records;

        var today     = DateTime.Today;
        var thisYear  = today.Year.ToString();
        var thisMonth = today.Month.ToString("D2");

        var byYear = records.GroupBy(r => r.연도).OrderByDescending(g => g.Key).ToList();

        foreach (var yearGroup in byYear)
        {
            string year       = yearGroup.Key;
            bool   isThisYear = year == thisYear;

            var yearNode = MakeParentNode($"📅  {year}년");
            yearNode.IsExpanded = isThisYear;

            var byMonth = yearGroup.GroupBy(r => r.월).OrderByDescending(g => g.Key).ToList();

            foreach (var monthGroup in byMonth)
            {
                bool isThisMonth = isThisYear && monthGroup.Key == thisMonth;
                var monthNode = MakeParentNode($"  {monthGroup.Key}월  ({monthGroup.Count()}건)");
                monthNode.IsExpanded = isThisMonth;

                foreach (var rec in monthGroup.OrderByDescending(r => r.의뢰일))
                    monthNode.Items.Add(MakeAnalysisLeaf(rec));

                yearNode.Items.Add(monthNode);
            }

            _treeAnalysis.Items.Add(yearNode);
        }

        txbInfo.Text = $"총 {records.Count}건";
        Log($"트리 완료 — 년노드={_treeAnalysis.Items.Count}");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  거래명세서발행내역 로드 — 년/월 2단계
    // ══════════════════════════════════════════════════════════════════════
    private async Task LoadStatementTreeAsync()
    {
        _treeStatement.Items.Clear();
        txbInfo.Text = "로딩 중...";
        Log("LoadStatementTreeAsync() 시작");

        List<TradeStatementRecord> records;
        try
        {
            records = await Task.Run(GetAllTradeStatements);
            Log($"DB 조회 완료: {records.Count}건");
        }
        catch (Exception ex)
        {
            Log($"DB 오류: {ex.Message}"); txbInfo.Text = "로드 실패"; return;
        }

        if (records.Count == 0) { txbInfo.Text = "발행된 거래명세서 없음"; return; }

        var today     = DateTime.Today;
        var thisYear  = today.Year.ToString();
        var thisMonth = today.Month.ToString("D2");

        var byYear = records.GroupBy(r => r.연도).OrderByDescending(g => g.Key).ToList();
        foreach (var yearGroup in byYear)
        {
            string year = yearGroup.Key;
            var yearNode = MakeParentNode($"📅  {year}년");
            yearNode.IsExpanded = year == thisYear;

            var byMonth = yearGroup.GroupBy(r => r.월).OrderByDescending(g => g.Key);
            foreach (var monthGroup in byMonth)
            {
                var monthNode = MakeParentNode($"  {monthGroup.Key}월  ({monthGroup.Count()}건)");
                monthNode.IsExpanded = year == thisYear && monthGroup.Key == thisMonth;

                foreach (var rec in monthGroup.OrderByDescending(r => r.발행일))
                    monthNode.Items.Add(MakeStatementLeaf(rec));

                yearNode.Items.Add(monthNode);
            }

            _treeStatement.Items.Add(yearNode);
        }

        txbInfo.Text = $"총 {records.Count}건";
        Log($"트리 완료 — 년노드={_treeStatement.Items.Count}");
    }

    private static List<TradeStatementRecord> GetAllTradeStatements()
    {
        var list = new List<TradeStatementRecord>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "거래명세서발행내역"))
            {
                Log("GetAllTradeStatements: 테이블 없음");
                return list;
            }

            bool hasId = DbConnectionFactory.ColumnExists(conn, "거래명세서발행내역", DbConnectionFactory.RowId);
            string idExpr = hasId ? $"`{DbConnectionFactory.RowId}`" : "0";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = $@"
                SELECT {idExpr},
                       COALESCE(`발행일`, ''),
                       COALESCE(`업체명`, ''),
                       COALESCE(`약칭`,   ''),
                       COALESCE(`거래명세서번호`, ''),
                       COALESCE(`공급가액`, 0),
                       COALESCE(`부가세`,   0),
                       COALESCE(`합계금액`, 0)
                FROM   `거래명세서발행내역`
                ORDER  BY `발행일` DESC";
            Log($"GetAllTradeStatements 쿼리: hasId={hasId}");

            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string date = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
                string y = "", m = "";
                if (date.Length >= 10 && (date[4] == '-' || date[4] == '/'))
                { y = date[..4]; m = date[5..7]; }
                else if (date.Length >= 8 && int.TryParse(date[..4], out _))
                { y = date[..4]; m = date[4..6]; }
                else if (date.Length >= 4) { y = date[..4]; }

                list.Add(new TradeStatementRecord
                {
                    Id             = rdr.IsDBNull(0) ? 0 : Convert.ToInt32(rdr.GetValue(0)),
                    발행일         = date,
                    업체명         = rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                    약칭           = rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                    거래명세서번호 = rdr.IsDBNull(4) ? "" : rdr.GetString(4),
                    공급가액       = rdr.IsDBNull(5) ? 0 : Convert.ToDecimal(rdr.GetValue(5)),
                    부가세         = rdr.IsDBNull(6) ? 0 : Convert.ToDecimal(rdr.GetValue(6)),
                    합계금액       = rdr.IsDBNull(7) ? 0 : Convert.ToDecimal(rdr.GetValue(7)),
                    연도 = y, 월 = m,
                });
            }
            Log($"GetAllTradeStatements: {list.Count}건 로드");
        }
        catch (Exception ex)
        {
            Log($"GetAllTradeStatements 오류: {ex.Message}");
        }
        return list;
    }

    private void OnStatementNodeSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (_treeStatement.SelectedItem is TreeViewItem item &&
            item.Tag is TradeStatementRecord rec)
        {
            txbInfo.Text = $"{rec.업체명}  |  {rec.거래명세서번호}  |  {rec.합계금액:N0}원";
            Log($"거래명세서 선택: {rec.거래명세서번호}");
            TradeStatementSelected?.Invoke(rec);
        }
    }

    private TreeViewItem MakeStatementLeaf(TradeStatementRecord rec)
    {
        var (bg, fg) = BadgeColorHelper.GetBadgeColor(rec.약칭);
        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };

        // 1줄: [약칭 뱃지]  업체명
        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
        topRow.Children.Add(new Border
        {
            Background = Brush.Parse(bg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 0,
            Child = Fs(new TextBlock { Text = rec.약칭, FontFamily = Font,
                                       Foreground = Brush.Parse(fg) }, "FontSizeXS"),
        });
        topRow.Children.Add(Fs(new TextBlock
        {
            Text = rec.업체명, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        }, "FontSizeBase"));
        sp.Children.Add(topRow);

        // 2줄: 거래명세서번호  ·  (MM/DD)  ·  합계금액
        string mmdd = rec.발행일.Length >= 10 ? $"({rec.발행일[5..10].Replace("-","/")})" : "";
        var bottomRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = rec.거래명세서번호, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        }, "FontSizeXS"));
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = mmdd, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        }, "FontSizeXS"));
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = $"{rec.합계금액:N0}원", FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 0, 0),
            [Grid.ColumnProperty] = 2,
        }, "FontSizeXS"));
        sp.Children.Add(bottomRow);

        return new TreeViewItem { Header = sp, Tag = rec };
    }


    private void OnAnalysisNodeSelected(object? sender, SelectionChangedEventArgs e)
    {
        // 발행 모드: 추가된 항목마다 이벤트 발생 (다중선택 지원)
        if (IssuingMode)
        {
            int added = 0;
            foreach (var obj in e.AddedItems)
            {
                if (obj is TreeViewItem tvi && tvi.Tag is AnalysisRequestRecord rec)
                {
                    AnalysisRequestAddedToList?.Invoke(rec);
                    added++;
                }
            }
            if (added > 0)
            {
                var all = GetSelectedAnalysisRecords();
                txbInfo.Text = $"{all.Count}건 선택 ({added}건 추가)";
            }
            return;
        }

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
            Foreground = AppTheme.FgInfo,
            Margin = new Thickness(2, 1),
        };
        Fs(tb, "FontSizeBase");
        var tvi = new TreeViewItem { Header = tb, IsExpanded = false };
        TextShimmer.AttachHover(tvi);
        return tvi;
    }

    private TreeViewItem MakeIssueLeaf(QuotationIssue issue)
    {
        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(issue.약칭);
        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };
        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,Auto,*") };

        // 거래명세서 발행 여부 아이콘 (약칭 앞)
        bool issued = !string.IsNullOrEmpty(issue.거래명세서번호);
        topRow.Children.Add(Fs(new TextBlock
        {
            Text              = issued ? "🔵" : "🔴",
            FontFamily        = Font,
            VerticalAlignment = VerticalAlignment.Center,
            MinWidth          = 16,
            Margin            = new Thickness(0, 0, 6, 0),
            [Grid.ColumnProperty] = 0,
        }, "FontSizeSM"));

        topRow.Children.Add(new Border
        {
            Background = Brush.Parse(ibg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 1,
            Child = Fs(new TextBlock { Text = issue.약칭, FontFamily = Font,
                                    Foreground = Brush.Parse(ifg) }, "FontSizeXS"),
        });
        // 시료명 + (MM/DD)
        string mmdd = issue.발행일.Length >= 10 ? $"({issue.발행일[5..10].Replace("-","/")})" : "";
        topRow.Children.Add(Fs(new TextBlock
        {
            Text = issue.시료명, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 2,
        }, "FontSizeBase"));
        sp.Children.Add(topRow);
        var bottomRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = issue.견적번호, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        }, "FontSizeXS"));
        if (!string.IsNullOrEmpty(issue.담당자))
            bottomRow.Children.Add(Fs(new TextBlock
            {
                Text = issue.담당자, FontFamily = Font,
                Foreground = AppTheme.FgMuted,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0),
                [Grid.ColumnProperty] = 1,
            }, "FontSizeXS"));
        bottomRow.Children.Add(Fs(new TextBlock
        {
            Text = mmdd, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 0, 0),
            [Grid.ColumnProperty] = 2,
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
            Foreground = AppTheme.FgPrimary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        }, "FontSizeBase"));
        sp.Children.Add(topRow);
        string recMmdd = rec.의뢰일.Length >= 10 ? $"({rec.의뢰일[5..10].Replace("-","/")})" : "";
        var bottomAnalysis = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        bottomAnalysis.Children.Add(Fs(new TextBlock
        {
            Text = rec.접수번호, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        }, "FontSizeXS"));
        bottomAnalysis.Children.Add(Fs(new TextBlock
        {
            Text = recMmdd, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 0, 1),
            [Grid.ColumnProperty] = 1,
        }, "FontSizeXS"));
        sp.Children.Add(bottomAnalysis);

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
