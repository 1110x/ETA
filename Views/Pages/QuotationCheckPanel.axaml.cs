using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Layout;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages;

/// <summary>Content3 — 분석항목 체크박스 패널 (간격 축소 + 외부 체크 동기화)</summary>
public partial class QuotationCheckPanel : UserControl
{
    public event Action<List<AnalysisItem>>? SelectionChanged;

    private List<AnalysisItem>                        _allItems   = [];
    // Analyte명 → CheckBox 빠른 조회
    private Dictionary<string, CheckBox>              _cbMap      = new(StringComparer.OrdinalIgnoreCase);
    // Analyte명 → 카테고리 CheckBox (상위 동기화용)
    private Dictionary<string, CheckBox>              _catMap     = new(StringComparer.OrdinalIgnoreCase);
    // 카테고리 → 자식 체크박스 목록
    private Dictionary<string, List<CheckBox>>        _catChildren = new(StringComparer.OrdinalIgnoreCase);

    private bool _suspendEvents = false;

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
        // 카테고리별 그룹화 (최대 7개)
        var groups = _allItems
            .GroupBy(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .OrderBy(g => g.Key)
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

            // 헤더: 카테고리명 + 항목수
            var headerText = new TextBlock
            {
                Text = $"{catKey}",
                FontSize = 12,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse("#a0d060"),
                Margin = new Avalonia.Thickness(0, 0, 0, 6),
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
            };
            Grid.SetRow(headerText, 0);
            columnContainer.Children.Add(headerText);

            // Border: 컬럼 시각적 구분
            var columnBorder = new Border
            {
                Background = Brush.Parse("#252530"),
                CornerRadius = new Avalonia.CornerRadius(4),
                BorderThickness = new Avalonia.Thickness(1),
                BorderBrush = Brush.Parse("#3a3a45"),
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

            // 카테고리 헤더 체크박스 (StackPanel 최상단)
            var catChk = new CheckBox
            {
                Content = $"전체 ({group.Count()})",
                IsChecked = false,
                FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse("#88bb88"),
                Margin = new Avalonia.Thickness(0, 0, 0, 4),
                Padding = new Avalonia.Thickness(2, 2),
            };
            itemsStack.Children.Add(catChk);

            // 항목 체크박스들
            foreach (var item in group.OrderBy(a => a.ES))
            {
                var cb = new CheckBox
                {
                    Content = item.Analyte,
                    IsChecked = false,
                    Tag = item,
                    FontSize = 10,
                    Foreground = Brush.Parse("#ddd"),
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
        foreach (var (catKey, children) in _catChildren)
        {
            // 카테고리 헤더 체크박스 찾기 (catMap 에서 첫 자식으로 역추적)
            var first = children.FirstOrDefault();
            if (first == null) continue;
            if (_catMap.TryGetValue((first.Tag as AnalysisItem)?.Analyte ?? "", out var catChk))
                SyncCategory(catChk, children);
        }
        UpdateCount();
    }

    /// <summary>모든 분석항목명 반환 (DetailPanel 이 체크 동기화에 사용)</summary>
    public List<string> GetAllAnalyteNames()
        => _cbMap.Keys.ToList();

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
}
