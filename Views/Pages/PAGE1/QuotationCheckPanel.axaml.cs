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
    private Contract?              _selectedCompany;  // Show3에서 선택한 업체

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

    /// <summary>Show3에서 선택한 업체</summary>
    public Contract? SelectedCompany
    {
        get => _selectedCompany;
        set
        {
            _selectedCompany = value;
            LoadData();  // 업체 변경 시 데이터 다시 로드
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

    // ── 카테고리별 커스텀 트리뷰 빌드 ────────────────────────────────────────
    private void BuildList()
    {
        // 카테고리 순서: 일반항목 → DB 등장 순 (첫 출현 기준)
        var catOrder = _allItems
            .Select(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .Distinct()
            .ToList();
        if (catOrder.Remove("일반항목"))
            catOrder.Insert(0, "일반항목");

        var groupDict = _allItems
            .GroupBy(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .ToDictionary(g => g.Key);

        var groups = catOrder
            .Where(k => groupDict.ContainsKey(k))
            .Select(k => groupDict[k])
            .ToList();

        var container = this.FindControl<StackPanel>("ItemsContainer");
        if (container == null) return;
        container.Children.Clear();

        foreach (var group in groups)
        {
            var catKey = group.Key;
            var children = new List<CheckBox>();
            _catChildren[catKey] = children;

            // ── 전체 선택 체크박스 ───────────────────────────────────────────
            var catChk = new CheckBox
            {
                Content = $"전체 ({group.Count()})",
                IsChecked = false,
                FontSize = AppFonts.XS,
                Foreground = Brush.Parse("#88bb88"),
                Padding = new Avalonia.Thickness(2, 0),
                Margin = new Avalonia.Thickness(0),
                MinHeight = 18,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };

            // ── 펼침/접힘 화살표 ─────────────────────────────────────────────
            var chevron = new TextBlock
            {
                Text = "▾",
                FontSize = 10,
                Foreground = AppTheme.FgMuted,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(0, 0, 4, 0),
            };

            // ── 카테고리 헤더 Border ─────────────────────────────────────────
            var headerRow = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"),
                Margin = new Avalonia.Thickness(0),
            };
            headerRow.Children.Add(chevron);
            Grid.SetColumn(chevron, 0);

            var catLabel = new TextBlock
            {
                Text = catKey,
                FontSize = AppFonts.SM,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgInfo,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            headerRow.Children.Add(catLabel);
            Grid.SetColumn(catLabel, 1);

            headerRow.Children.Add(catChk);
            Grid.SetColumn(catChk, 2);

            var headerBorder = new Border
            {
                Classes = { "cat-header" },
                Background = AppTheme.BgSecondary,
                Padding = new Avalonia.Thickness(8, 5),
                Margin = new Avalonia.Thickness(0, 0, 0, 1),
                CornerRadius = new Avalonia.CornerRadius(4),
                Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child = headerRow,
            };

            // ── 항목 목록 패널 ───────────────────────────────────────────────
            var itemsPanel = new Border
            {
                Background = AppTheme.BgPrimary,
                Padding = new Avalonia.Thickness(10, 4, 4, 4),
                Margin = new Avalonia.Thickness(0, 0, 0, 4),
                CornerRadius = new Avalonia.CornerRadius(0, 0, 4, 4),
                IsVisible = true,
            };
            var itemsStack = new StackPanel { Spacing = 0 };
            itemsPanel.Child = itemsStack;

            // 금액 캐시 (선택된 업체가 있을 경우)
            var priceDict = new Dictionary<string, decimal>();
            if (_selectedCompany != null)
            {
                try
                {
                    priceDict = QuotationService.GetPricesByCompany(_selectedCompany.C_CompanyName);
                }
                catch { }
            }

            foreach (var item in group.OrderBy(a => a.ES))
            {
                // 배지 (약칭 표시 - BadgeHelper.Abbr() 스타일)
                var (bgColor, fgColor) = BadgeColorHelper.GetBadgeColor(item.약칭);
                var badge = new Border
                {
                    Background = new SolidColorBrush(Color.Parse(bgColor)),
                    BorderBrush = new SolidColorBrush(Color.Parse(fgColor)),
                    BorderThickness = new Avalonia.Thickness(1),
                    CornerRadius = new Avalonia.CornerRadius(10),
                    Padding = new Avalonia.Thickness(6, 1, 8, 1),
                    Margin = new Avalonia.Thickness(0, 0, 4, 0),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = string.IsNullOrEmpty(item.약칭) ? "✓" : item.약칭,
                        FontSize = AppFonts.SM,
                        FontWeight = FontWeight.Medium,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        Foreground = new SolidColorBrush(Color.Parse(fgColor)),
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    }
                };

                // 금액 조회
                var priceText = "";
                if (priceDict.TryGetValue(item.Analyte, out var p))
                    priceText = $"{p:N0}원";

                // 레이아웃: [배지] 항목명 금액
                var contentGrid = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"),
                    Margin = new Avalonia.Thickness(0),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                };

                var itemNameTb = new TextBlock
                {
                    Text = item.Analyte,
                    FontSize = AppFonts.SM,
                    Foreground = AppTheme.FgPrimary,
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                };
                Grid.SetColumn(itemNameTb, 1);

                var priceTb = new TextBlock
                {
                    Text = priceText,
                    FontSize = AppFonts.XS,
                    Foreground = AppTheme.FgMuted,
                    Margin = new Avalonia.Thickness(8, 0, 0, 0),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                };
                Grid.SetColumn(priceTb, 2);

                contentGrid.Children.Add(badge);
                Grid.SetColumn(badge, 0);
                contentGrid.Children.Add(itemNameTb);
                contentGrid.Children.Add(priceTb);

                var cb = new CheckBox
                {
                    Content = contentGrid,
                    IsChecked = false,
                    Tag = item,
                    FontSize = AppFonts.SM,
                    Foreground = AppTheme.FgPrimary,
                    Padding = new Avalonia.Thickness(2, 1),
                    Margin = new Avalonia.Thickness(0),
                    MinHeight = 22,
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

            // ── 카테고리 전체 선택 로직 ──────────────────────────────────────
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

            // ── 헤더 클릭 → 펼침/접힘 토글 ─────────────────────────────────
            headerBorder.PointerPressed += (_, e) =>
            {
                // 체크박스 클릭은 전파시키고 토글 안 함
                if (e.Source is CheckBox) return;
                itemsPanel.IsVisible = !itemsPanel.IsVisible;
                chevron.Text = itemsPanel.IsVisible ? "▾" : "▸";
                // 헤더 하단 모서리: 접혔을 때 둥글게
                headerBorder.CornerRadius = itemsPanel.IsVisible
                    ? new Avalonia.CornerRadius(4, 4, 0, 0)
                    : new Avalonia.CornerRadius(4);
            };

            // ── 섹션 조합 ────────────────────────────────────────────────────
            var section = new StackPanel { Margin = new Avalonia.Thickness(0, 0, 0, 2) };
            section.Children.Add(headerBorder);
            section.Children.Add(itemsPanel);
            container.Children.Add(section);
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
