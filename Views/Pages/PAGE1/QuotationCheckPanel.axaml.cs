using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Layout;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq;
using ETA.Views;

namespace ETA.Views.Pages.PAGE1;

/// <summary>Content3 — 분석항목 체크박스 패널 (간격 축소 + 외부 체크 동기화)</summary>
public partial class QuotationCheckPanel : UserControl
{
    public event Action<List<AnalysisItem>>? SelectionChanged;

    // ── 편집 모드 이벤트 (저장 완료 시 발생) ────────────────────────────
    public event Action<AnalysisRequestRecord>? AnalysisRecordSaved;
    public event Action<QuotationIssue>?        IssueSaved;

    private List<AnalysisItem>                        _allItems   = [];
    private Dictionary<string, CheckBox>              _cbMap      = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, CheckBox>              _catMap     = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, List<CheckBox>>        _catChildren = new(StringComparer.OrdinalIgnoreCase);

    private bool _suspendEvents = false;
    private bool _isDirty       = false;

    // ── 편집 대상 (둘 중 하나만 non-null) ───────────────────────────────
    private AnalysisRequestRecord? _currentAnalysisRecord;
    private QuotationIssue?        _currentIssue;

    public AnalysisRequestRecord? CurrentAnalysisRecord
    {
        get => _currentAnalysisRecord;
        set
        {
            _currentAnalysisRecord = value;
            _currentIssue          = null;
            RefreshSaveBar();
        }
    }

    public QuotationIssue? CurrentIssue
    {
        get => _currentIssue;
        set
        {
            _currentIssue          = value;
            _currentAnalysisRecord = null;
            RefreshSaveBar();
        }
    }

    /// <summary>편집 모드 해제</summary>
    public void ClearEditTarget()
    {
        _currentAnalysisRecord = null;
        _currentIssue          = null;
        _isDirty               = false;
        RefreshSaveBar();
    }

    private void RefreshSaveBar()
    {
        _isDirty = false;
        var border = this.FindControl<Border>("saveBorder");
        var dirty  = this.FindControl<TextBlock>("txbDirty");
        if (border != null) border.IsVisible = _currentAnalysisRecord != null || _currentIssue != null;
        if (dirty  != null) dirty.IsVisible  = false;
    }

    private void SetDirty()
    {
        _isDirty = true;
        var dirty = this.FindControl<TextBlock>("txbDirty");
        if (dirty != null) dirty.IsVisible = true;
    }

    public QuotationCheckPanel()
    {
        InitializeComponent();
    }

    // ── 로드 ─────────────────────────────────────────────────────────────
    public void LoadData()
    {
        _allItems    = AnalysisService.GetAllItems();
        _cbMap       = new(StringComparer.OrdinalIgnoreCase);
        _catMap      = new(StringComparer.OrdinalIgnoreCase);
        _catChildren = new(StringComparer.OrdinalIgnoreCase);
        BuildList();
        UpdateCount();
    }

    // ── 7개 컬럼 독립 스크롤 레이아웃 빌드 ─────────────────────────────────
    private void BuildList()
    {
        // 카테고리 순서: 일반항목 → DB 분석정보 등장 순 (첫 출현 기준)
        var catOrder = _allItems
            .Select(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .Distinct()
            .ToList();
        // 일반항목을 맨 앞으로
        if (catOrder.Remove("일반항목"))
            catOrder.Insert(0, "일반항목");

        var groupDict = _allItems
            .GroupBy(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .ToDictionary(g => g.Key);

        var groups = catOrder
            .Where(k => groupDict.ContainsKey(k))
            .Select(k => groupDict[k])
            .ToList();

        // ColumnsGrid 초기화
        var columnsGrid = this.FindControl<Grid>("ColumnsGrid");
        if (columnsGrid == null) return;
        columnsGrid.Children.Clear();

        int columnIndex = 0;
        foreach (var group in groups)
        {
            if (columnIndex >= 7) break; // 최대 7개 컬럼만 표시

            var catKey = group.Key;
            var children = new List<CheckBox>();
            _catChildren[catKey] = children;

            // ═══════════════════════════════════════════════════════════════
            // 컬럼 레이아웃 구조:
            // Grid (RowDefinitions: Auto, *)
            //   ├─ Row=0: TextBlock (Category Header)
            //   └─ Row=1: Border
            //        └─ ScrollViewer (독립 수직 스크롤)
            //             └─ StackPanel (Items)
            // ═══════════════════════════════════════════════════════════════

            // 컬럼 컨테이너 그리드
            var columnContainer = new Grid
            {
                RowDefinitions = new RowDefinitions("Auto,*"),
                Margin = new Avalonia.Thickness(0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch,
            };
            Grid.SetColumn(columnContainer, columnIndex);

            // 헤더: 카테고리명 + 전체 체크박스 (가로 배치)
            var catChk = new CheckBox
            {
                Content = $"전체 ({group.Count()})",
                IsChecked = false,
                FontSize = AppFonts.Base,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse("#88bb88"),
                Margin = new Avalonia.Thickness(6, 0, 0, 0),
                Padding = new Avalonia.Thickness(2, 2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            var headerPanel = new StackPanel
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                Margin = new Avalonia.Thickness(0, 0, 0, 6),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            var headerText = new TextBlock
            {
                Text = $"{catKey}",
                FontSize = AppFonts.MD,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgInfo,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            headerPanel.Children.Add(headerText);
            headerPanel.Children.Add(catChk);
            Grid.SetRow(headerPanel, 0);
            columnContainer.Children.Add(headerPanel);

            // Border: 컬럼 시각적 구분
            var columnBorder = new Border
            {
                Background = AppTheme.BgSecondary,
                CornerRadius = new Avalonia.CornerRadius(4),
                BorderThickness = new Avalonia.Thickness(1),
                BorderBrush = AppTheme.BorderSubtle,
                Padding = new Avalonia.Thickness(8),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch,
            };
            Grid.SetRow(columnBorder, 1);

            // ScrollViewer: 독립 수직 스크롤 (기본 동작: 필요할 때만 scrollbar 표시)
            var scrollViewer = new ScrollViewer
            {
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch,
            };

            // ItemsPanel: StackPanel (항목)
            var itemsStack = new StackPanel
            {
                Spacing = 4,
                Orientation = Avalonia.Layout.Orientation.Vertical,
            };

            // (catChk는 헤더 패널로 이동됨)

            // 항목 체크박스들
            foreach (var item in group.OrderBy(a => a.ES))
            {
                var cb = new CheckBox
                {
                    Content = item.Analyte,
                    IsChecked = false,
                    Tag = item,
                    FontSize = AppFonts.SM,
                    Foreground = AppTheme.FgPrimary,
                    Margin = new Avalonia.Thickness(0, 0, 0, 2),
                    Padding = new Avalonia.Thickness(2, 1),
                };
                Avalonia.Controls.ToolTip.SetTip(cb,
                    $"ES: {item.ES}\n단위: {item.unit}\n방법: {item.Method}");

                cb.IsCheckedChanged += (_, _) =>
                {
                    if (_suspendEvents) return;
                    SyncCategory(catChk, children);
                    UpdateCount();
                    if (_currentAnalysisRecord != null || _currentIssue != null) SetDirty();
                    SelectionChanged?.Invoke(GetSelected());
                };

                children.Add(cb);
                _cbMap[item.Analyte] = cb;
                _catMap[item.Analyte] = catChk;
                itemsStack.Children.Add(cb);
            }

            // 카테고리 헤더 전체 선택/해제 로직
            catChk.IsCheckedChanged += (_, _) =>
            {
                if (_suspendEvents || catChk.IsChecked == null) return;
                _suspendEvents = true;
                foreach (var cb in children)
                    cb.IsChecked = catChk.IsChecked == true;
                _suspendEvents = false;
                UpdateCount();
                if (_currentAnalysisRecord != null || _currentIssue != null) SetDirty();
                SelectionChanged?.Invoke(GetSelected());
            };

            // ScrollViewer 설정
            scrollViewer.Content = itemsStack;

            // Border 설정
            columnBorder.Child = scrollViewer;

            // 컬럼 컨테이너에 추가
            columnContainer.Children.Add(columnBorder);

            // 메인 ColumnsGrid에 추가
            columnsGrid.Children.Add(columnContainer);

            columnIndex++;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  외부 연동 API
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>DetailPanel 에서 rowid row 기준으로 체크 상태를 일괄 세팅</summary>
    public void SetChecked(string analyteName, bool isChecked)
    {
        if (_cbMap.TryGetValue(analyteName, out var cb))
        {
            _suspendEvents = true;
            cb.IsChecked   = isChecked;
            _suspendEvents = false;
        }
    }

    /// <summary>체크 동기화 완료 후 카테고리 헤더 일괄 갱신</summary>
    public void SyncAllCategories()
    {
        // 이벤트 일시 중단 — catChk.IsChecked 변경이 IsCheckedChanged 핸들러를 통해
        // 모든 자식 체크박스를 덮어쓰는 것을 방지
        _suspendEvents = true;
        foreach (var (catKey, children) in _catChildren)
        {
            var first = children.FirstOrDefault();
            if (first == null) continue;
            if (_catMap.TryGetValue((first.Tag as AnalysisItem)?.Analyte ?? "", out var catChk))
                SyncCategory(catChk, children);
        }
        _suspendEvents = false;
        UpdateCount();
    }

    /// <summary>모든 분석항목명 반환 (DetailPanel 이 체크 동기화에 사용)</summary>
    public List<string> GetAllAnalyteNames()
        => _cbMap.Keys.ToList();

    /// <summary>항목명으로 AnalysisItem 메타 반환 (배지 표시용)</summary>
    public AnalysisItem? GetItem(string analyteName)
    {
        if (_cbMap.TryGetValue(analyteName, out var cb) && cb.Tag is AnalysisItem item)
            return item;
        return null;
    }

    /// <summary>현재 체크된 항목 반환</summary>
    public List<AnalysisItem> GetSelected()
        => _cbMap.Values
            .Where(c => c.IsChecked == true && c.Tag is AnalysisItem)
            .Select(c => (AnalysisItem)c.Tag!)
            .ToList();

    // ── 전체 선택 / 해제 버튼 ────────────────────────────────────────────
    private void BtnSelectAll_Click(object? sender, RoutedEventArgs e)
    {
        _suspendEvents = true;
        foreach (var cb in _cbMap.Values) cb.IsChecked = true;
        _suspendEvents = false;
        SyncAllCategories();
        SelectionChanged?.Invoke(GetSelected());
    }

    public void ClearAll()
    {
        _suspendEvents = true;
        foreach (var cb in _cbMap.Values) cb.IsChecked = false;
        _suspendEvents = false;
        SyncAllCategories();
        SelectionChanged?.Invoke(GetSelected());
    }

    private void BtnClearAll_Click(object? sender, RoutedEventArgs e) => ClearAll();

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static void SyncCategory(CheckBox cat, List<CheckBox> children)
    {
        int n = children.Count(c => c.IsChecked == true);
        cat.IsChecked = n == 0 ? false : n == children.Count ? (bool?)true : null;
    }

    private void UpdateCount()
    {
        int sel = _cbMap.Values.Count(c => c.IsChecked == true);
        txbCount.Text = $"선택 {sel}개 / 전체 {_allItems.Count}개";
    }

    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [CheckPanel] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  저장 / 취소 버튼
    // ══════════════════════════════════════════════════════════════════════
    private async void BtnSaveAnalytes_Click(object? sender, RoutedEventArgs e)
    {
        var allNames     = _cbMap.Keys.ToList();
        var checkedNames = GetSelected().Select(a => a.Analyte).ToList();

        Log($"저장 클릭 — rec={_currentAnalysisRecord?.Id.ToString() ?? "null"}  issue={_currentIssue?.Id.ToString() ?? "null"}  allNames={allNames.Count}  checked={checkedNames.Count}");

        bool ok = false;
        if (_currentAnalysisRecord != null)
        {
            var rec = _currentAnalysisRecord;
            ok = await System.Threading.Tasks.Task.Run(() =>
                AnalysisRequestService.UpdateAnalyteValues(rec.Id, allNames, checkedNames));
            if (ok) AnalysisRecordSaved?.Invoke(rec);
        }
        else if (_currentIssue != null)
        {
            var issue = _currentIssue;
            ok = await System.Threading.Tasks.Task.Run(() =>
                QuotationService.UpdateIssueAnalytes(issue.Id, allNames, checkedNames));
            if (ok) IssueSaved?.Invoke(issue);
        }

        Log($"저장 완료 — ok={ok}");

        if (ok)
        {
            _isDirty = false;
            var dirty = this.FindControl<TextBlock>("txbDirty");
            if (dirty != null) dirty.IsVisible = false;
        }
    }

    private void BtnCancelEdit_Click(object? sender, RoutedEventArgs e)
    {
        // 체크 상태를 DB에서 다시 불러오도록 이벤트 발생
        if (_currentAnalysisRecord != null)
            AnalysisRecordSaved?.Invoke(_currentAnalysisRecord);
        else if (_currentIssue != null)
            IssueSaved?.Invoke(_currentIssue);

        _isDirty = false;
        var dirty = this.FindControl<TextBlock>("txbDirty");
        if (dirty != null) dirty.IsVisible = false;
    }
}
