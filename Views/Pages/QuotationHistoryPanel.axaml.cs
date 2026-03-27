using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

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
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Treeview.log"));
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

        _treeAnalysis = new TreeView
        {
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            IsVisible = false,
        };
        _treeAnalysis.SelectionChanged += OnAnalysisNodeSelected;

        treeHost.Children.Add(_treeQuotation);
        treeHost.Children.Add(_treeAnalysis);

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

    private void BtnDelete_Click(object? sender, RoutedEventArgs e)
    {
        if (_isAnalysisTab) return;
        if (_treeQuotation.SelectedItem is TreeViewItem item &&
            item.Tag is QuotationIssue issue)
        {
            Log($"삭제: {issue.견적번호}");
            if (QuotationService.Delete(issue.Id)) LoadData();
        }
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
                        Header = new TextBlock
                        {
                            Text = "...", FontSize = 9, FontFamily = Font,
                            Foreground = Brush.Parse("#444455"),
                        },
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
        if (_treeAnalysis.SelectedItem is TreeViewItem item &&
            item.Tag is AnalysisRequestRecord rec)
        {
            txbInfo.Text = $"{rec.약칭}  {rec.시료명}  |  {rec.접수번호}";
            Log($"의뢰 선택: {rec.약칭} {rec.시료명} [{rec.접수번호}]");
            AnalysisRequestSelected?.Invoke(rec);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  노드 헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private TreeViewItem MakeParentNode(string header) => new()
    {
        Header = new TextBlock
        {
            Text = header, FontSize = 11, FontFamily = Font,
            Foreground = Brush.Parse("#aaaacc"),
            Margin = new Thickness(2, 1),
        },
        IsExpanded = false,
    };

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
            Child = new TextBlock { Text = issue.약칭, FontSize = 9, FontFamily = Font,
                                    Foreground = Brush.Parse(ifg) },
        });
        topRow.Children.Add(new TextBlock
        {
            Text = issue.시료명, FontSize = 11, FontFamily = Font,
            Foreground = Brush.Parse("#dddddd"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        });
        sp.Children.Add(topRow);
        sp.Children.Add(new TextBlock
        {
            Text = issue.견적번호, FontSize = 9, FontFamily = Font,
            Foreground = Brush.Parse("#445566"), Margin = new Thickness(0, 0, 0, 1),
        });
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
            Child = new TextBlock { Text = rec.약칭, FontSize = 9, FontFamily = Font,
                                    Foreground = Brush.Parse(rfg) },
        });
        topRow.Children.Add(new TextBlock
        {
            Text = rec.시료명, FontSize = 11, FontFamily = Font,
            Foreground = Brush.Parse("#dddddd"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        });
        sp.Children.Add(topRow);
        sp.Children.Add(new TextBlock
        {
            Text = rec.접수번호, FontSize = 9, FontFamily = Font,
            Foreground = Brush.Parse("#445566"), Margin = new Thickness(0, 0, 0, 1),
        });
        return new TreeViewItem { Header = sp, Tag = rec };
    }
}
