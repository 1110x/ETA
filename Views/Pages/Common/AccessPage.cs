using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.SERVICE3;
using ETA.Services.Common;
using ETA.Views;
using ETA.Views.Controls;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 관리자 페이지
///   탭 1 — 권한관리: 직원별 메뉴 접근 설정
///   탭 2 — 처리시설 매칭: 외부 시스템(WAYBLE / 측정인.kr) ↔ 처리시설_마스터 매핑
///   탭 3 — 분석항목 관리: 분석정보 CRUD + 시험기록부 테이블 생성
///   탭 4 — 분석장비 관리: 측정인_분석장비 테이블 수동 CRUD
///   탭 5 — 측정인계약: add_meas_cont_no + cmb_emis_cmpy_plc_no 스크랩
/// </summary>
public class AccessPage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly string[] DeptItems = { "전체", "수질분석센터", "처리시설", "일반업무", "기타" };

    // ── 외부 연결 ─────────────────────────────────────────────────────────────
    public Control Show1 { get; }
    public Control Show2 { get; }
    public Control Show4 { get; }

    // ── 공통 상태 ─────────────────────────────────────────────────────────────
    private readonly ContentControl _show2Content;
    private readonly ContentControl _show1Switcher;
    private readonly ContentControl _show4Content;
    private string _mode = "access"; // "access" | "facilityMatch" | "analyte" | "equipment" | "mcontract"

    // ── 탭 버튼 ───────────────────────────────────────────────────────────────
    private Border? _tabAccess;
    private Border? _tabFacilityMatch;
    private Border? _tabAnalyte;
    private Border? _tabEquipment;
    private Border? _tabMeasurerContract;

    // ── 권한관리 상태 ─────────────────────────────────────────────────────────
    private Agent?           _selectedAgent;
    private string           _deptFilter = "전체";
    private string           _searchText = "";
    private readonly TreeView _treeView;
    private DispatcherTimer?  _debounce;

    // ── 분석항목 관리 상태 ────────────────────────────────────────────────────
    private AnalysisItem?    _selectedAnalyte;
    private readonly TreeView _analyteTree;

    // ── 공식 빌더 상태 ────────────────────────────────────────────────────────
    private List<string> _formulaTokens = new();
    private string _formulaAnalyte = "";
    private WrapPanel? _formulaDisplayPanel;

    // ── 컬럼 편집 모드 ────────────────────────────────────────────────────────
    private bool _colEditMode = false;

    // ── 분석장비 관리 상태 ────────────────────────────────────────────────────
    private (string 장비명, string 코드값)? _selectedEquipment;
    private readonly StackPanel _equipmentListPanel; // 장비 목록 행들

    // ── 측정인계약 상태 ───────────────────────────────────────────────────────
    private readonly StackPanel _measurerContractListPanel;

    // ── 처리시설 매칭 상태 ────────────────────────────────────────────────────
    private readonly StackPanel _facilityMatchListPanel;
    private (string 시설명, string 시료명)? _selectedFacilityPair;

    // ── 패널 캐시 ─────────────────────────────────────────────────────────────
    private readonly Control _employeePanelRef;
    private readonly Control _analytePanelRef;
    private readonly Control _equipmentPanelRef;
    private readonly Control _measurerContractPanelRef;
    private readonly Control _facilityMatchPanelRef;

    // 표시 플래그 ── 측정인관리 메뉴에서 진입할 때 measurerOnly=true 로 호출
    private readonly bool _showAccessTab;
    private readonly bool _showMeasurerTabs;

    // =========================================================================
    public AccessPage() : this("access", showAccessTab: true, showMeasurerTabs: false) { }

    /// <param name="initialMode">"access" | "analyte" | "equipment" | "mcontract"</param>
    /// <param name="showAccessTab">권한관리 탭 노출 여부</param>
    /// <param name="showMeasurerTabs">분석항목/분석장비/측정인계약 탭 노출 여부 (측정인관리 진입 시 true)</param>
    public AccessPage(string initialMode, bool showAccessTab, bool showMeasurerTabs)
    {
        _showAccessTab    = showAccessTab;
        _showMeasurerTabs = showMeasurerTabs;

        AccessService.EnsureTable();
        MeasurerService.EnsureEquipmentTable();
        MeasurerService.EnsureMeasurerContractTables();
        FacilityMappingService.EnsureTable();

        _treeView          = BuildEmployeeTreeView();
        _analyteTree       = BuildAnalyteTreeView();
        _equipmentListPanel = new StackPanel { Spacing = 1 };
        _measurerContractListPanel = new StackPanel { Spacing = 1 };
        _facilityMatchListPanel    = new StackPanel { Spacing = 1 };

        _show2Content = new ContentControl
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Stretch,
        };
        _show1Switcher = new ContentControl
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Stretch,
        };
        _show4Content = new ContentControl
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Stretch,
        };

        _employeePanelRef  = BuildEmployeePanel();
        _analytePanelRef   = BuildAnalytePanel();
        _equipmentPanelRef = BuildEquipmentPanel();
        _measurerContractPanelRef = BuildMeasurerContractPanel();
        _facilityMatchPanelRef    = BuildFacilityMatchPanel();

        Show1 = BuildShow1Shell();
        Show2 = new Border { Padding = new Thickness(6), Child = _show2Content };
        Show4 = new Border { Padding = new Thickness(6), Child = _show4Content };

        SwitchMode(initialMode);
        if (_showAccessTab) LoadEmployeeTree();
    }

    // =========================================================================
    // Show1 외곽: 탭 바 + 스위처
    // =========================================================================
    private Control BuildShow1Shell()
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        root.RowDefinitions.Add(new RowDefinition(GridLength.Star));
        var tabBar = BuildTabBar();
        Grid.SetRow(tabBar, 0);
        Grid.SetRow(_show1Switcher, 1);
        root.Children.Add(tabBar);
        root.Children.Add(_show1Switcher);
        return root;
    }

    private Border BuildTabBar()
    {
        var stack = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 2 };
        if (_showAccessTab)
        {
            _tabAccess = MakeTabBtn("🔐 권한관리", active: true, () => SwitchMode("access"));
            stack.Children.Add(_tabAccess);
            _tabFacilityMatch = MakeTabBtn("🏭 처리시설 매칭", active: false, () => SwitchMode("facilityMatch"));
            stack.Children.Add(_tabFacilityMatch);
        }
        if (_showMeasurerTabs)
        {
            _tabAnalyte          = MakeTabBtn("🧪 분석항목",    active: false, () => SwitchMode("analyte"));
            _tabEquipment        = MakeTabBtn("🔧 분석장비",    active: false, () => SwitchMode("equipment"));
            _tabMeasurerContract = MakeTabBtn("🧾 계약·측정시설", active: false, () => SwitchMode("mcontract"));
            stack.Children.Add(_tabAnalyte);
            stack.Children.Add(_tabEquipment);
            stack.Children.Add(_tabMeasurerContract);
        }
        return new Border
        {
            Background      = AppRes("PanelBg"),
            BorderBrush     = AppTheme.BorderDefault,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding         = new Thickness(6, 4, 6, 0),
            Child           = stack,
        };
    }

    private static Border MakeTabBtn(string text, bool active, Action onClick)
    {
        var tb = new TextBlock
        {
            Text       = text,
            FontSize   = AppTheme.FontMD,
            FontFamily = Font,
            Foreground = active ? AppTheme.FgInfo : AppRes("FgMuted"),
            FontWeight = active ? FontWeight.SemiBold : FontWeight.Normal,
            Padding    = new Thickness(10, 6),
        };
        var btn = new Border
        {
            Background   = active ? AppTheme.BgSecondary : Brushes.Transparent,
            CornerRadius = new CornerRadius(4, 4, 0, 0),
            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child        = tb,
        };
        btn.PointerPressed += (_, _) => onClick();
        return btn;
    }

    private void SwitchMode(string mode)
    {
        _mode = mode;
        UpdateTabStyle(_tabAccess,           mode == "access");
        UpdateTabStyle(_tabFacilityMatch,    mode == "facilityMatch");
        UpdateTabStyle(_tabAnalyte,          mode == "analyte");
        UpdateTabStyle(_tabEquipment,        mode == "equipment");
        UpdateTabStyle(_tabMeasurerContract, mode == "mcontract");

        switch (mode)
        {
            case "facilityMatch":
                _show1Switcher.Content = _facilityMatchPanelRef;
                _show2Content.Content  = BuildFacilityMatchHint();
                _show4Content.Content  = BuildWaybleCatalogPanel();
                LoadFacilityMatchList();
                break;
            case "analyte":
                _show1Switcher.Content = _analytePanelRef;
                _show2Content.Content  = BuildAnalyteHint();
                _show4Content.Content  = null;
                LoadAnalyteTree();
                break;
            case "equipment":
                _show1Switcher.Content = _equipmentPanelRef;
                _show2Content.Content  = BuildEquipmentHint();
                _show4Content.Content  = null;
                LoadEquipmentList();
                break;
            case "mcontract":
                _show1Switcher.Content = _measurerContractPanelRef;
                _show2Content.Content  = BuildMeasurerContractHint();
                _show4Content.Content  = null;
                LoadMeasurerContractList();
                break;
            default:
                _show1Switcher.Content = _employeePanelRef;
                _show2Content.Content  = BuildEmptyHint();
                _show4Content.Content  = null;
                break;
        }

        Show4VisibleChanged?.Invoke(mode == "facilityMatch");
    }

    /// <summary>Show4 콘텐츠 가시성 변경 통지 (MainPage 가 content4Star 조절용)</summary>
    public event Action<bool>? Show4VisibleChanged;

    private void UpdateTabStyle(Border? btn, bool active)
    {
        if (btn == null) return;
        btn.Background = active ? AppTheme.BgSecondary : Brushes.Transparent;
        if (btn.Child is TextBlock tb)
        {
            tb.Foreground = active ? AppTheme.FgInfo : AppRes("FgMuted");
            tb.FontWeight = active ? FontWeight.SemiBold : FontWeight.Normal;
        }
    }

    // =========================================================================
    // 탭 1 — 권한관리 직원 패널
    // =========================================================================
    private Control BuildEmployeePanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var header = new Border
        {
            Background = AppRes("PanelBg"),
            Padding    = new Thickness(10, 6),
            Child      = new TextBlock { Text = "직원 목록", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"), FontFamily = Font }
        };
        Grid.SetRow(header, 0);

        var searchBox = new TextBox
        {
            Watermark = "🔍 직원 검색...", FontSize = AppTheme.FontMD, FontFamily = Font,
            Height = 32, Margin = new Thickness(5, 4, 5, 2), CornerRadius = new CornerRadius(6),
            Padding = new Thickness(8, 5), BorderThickness = new Thickness(1),
            BorderBrush = AppRes("InputBorder"), Background = AppRes("PanelBg"),
        };
        searchBox.TextChanged += (_, _) =>
        {
            _debounce?.Stop();
            _debounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _debounce.Tick += (_, _) => { _debounce!.Stop(); _searchText = searchBox.Text?.Trim().ToLower() ?? ""; ApplyEmployeeFilters(); };
            _debounce.Start();
        };
        Grid.SetRow(searchBox, 1);

        var deptBox = new ComboBox
        {
            ItemsSource = DeptItems, SelectedIndex = 0,
            FontSize = AppTheme.FontMD, FontFamily = Font, Height = 30,
            Margin = new Thickness(5, 0, 5, 3), Padding = new Thickness(8, 4),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            BorderThickness = new Thickness(1), BorderBrush = AppRes("InputBorder"), Background = AppRes("PanelBg"),
        };
        deptBox.SelectionChanged += (_, _) => { _deptFilter = deptBox.SelectedItem?.ToString() ?? "전체"; ApplyEmployeeFilters(); };
        Grid.SetRow(deptBox, 2);
        Grid.SetRow(_treeView, 3);

        grid.Children.Add(header); grid.Children.Add(searchBox); grid.Children.Add(deptBox); grid.Children.Add(_treeView);
        return grid;
    }

    private TreeView BuildEmployeeTreeView()
    {
        var tv = new TreeView { Margin = new Thickness(5) };
        tv.SelectionChanged += EmployeeTree_SelectionChanged;
        return tv;
    }

    private void LoadEmployeeTree()
    {
        _treeView.Items.Clear();
        _selectedAgent = null;
        try
        {
            foreach (var a in AgentService.GetAllItems().OrderBy(x => x.입사일))
                _treeView.Items.Add(CreateEmployeeTreeItem(a));
        }
        catch { }
    }

    private TreeViewItem CreateEmployeeTreeItem(Agent agent)
    {
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
        row.Children.Add(MakePhotoCircle(agent, 26));
        if (!string.IsNullOrWhiteSpace(agent.직급))
        {
            var (bbg, bfg) = BadgeColorHelper.GetBadgeColor(agent.직급);
            row.Children.Add(new Border { Background = new SolidColorBrush(Color.Parse(bbg)), CornerRadius = new CornerRadius(3), Padding = new Thickness(5, 2), VerticalAlignment = VerticalAlignment.Center, Child = new TextBlock { Text = agent.직급, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse(bfg)), FontFamily = Font } });
        }
        row.Children.Add(new TextBlock { Text = agent.성명, FontSize = AppTheme.FontLG, FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center });
        if (!string.IsNullOrWhiteSpace(agent.부서))
            row.Children.Add(new TextBlock { Text = $"({agent.부서})", FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
        return new TreeViewItem { Header = row, Tag = agent };
    }

    private void ApplyEmployeeFilters()
    {
        foreach (var tvi in _treeView.Items.OfType<TreeViewItem>())
        {
            if (tvi.Tag is not Agent a) continue;
            bool nameOk = string.IsNullOrEmpty(_searchText) || (a.성명?.ToLower().Contains(_searchText) ?? false) || (a.사번?.ToLower().Contains(_searchText) ?? false);
            bool deptOk = _deptFilter == "전체" || a.부서 == _deptFilter;
            tvi.IsVisible = nameOk && deptOk;
        }
    }

    private void EmployeeTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0) return;
        if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is Agent a)
        {
            _selectedAgent = a;
            _show2Content.Content = BuildAccessPanel(a);
        }
    }

    // =========================================================================
    // 탭 1 — Show2 접근 권한 패널
    // =========================================================================
    private Control BuildAccessPanel(Agent agent)
    {
        var accessMap = AccessService.GetAccessMap(agent.사번);
        var root = new StackPanel { Spacing = 0, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock { Text = $"🔐 {agent.성명} ({agent.사번}) — 메뉴 접근 권한", FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 6) });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 8) });
        root.Children.Add(new TextBlock { Text = "비활성화된 메뉴는 해당 직원에게 숨김 처리됩니다.", FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 12) });
        foreach (var cat in AccessService.AllMenus.GroupBy(m => m.Category))
        {
            root.Children.Add(new Border { Background = AppTheme.BgSecondary, CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4), Margin = new Thickness(0, 4, 0, 2), Child = new TextBlock { Text = cat.Key, FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppTheme.FgInfo } });
            foreach (var (key, label, _) in cat)
            {
                bool allowed = accessMap.TryGetValue(key, out var v) ? v : true;
                root.Children.Add(BuildToggleRow(agent.사번, key, label, allowed));
            }
        }
        return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = root };
    }

    private Control BuildToggleRow(string 사번, string key, string label, bool allowed)
    {
        var row = new Grid { Margin = new Thickness(0, 1, 0, 1) };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        var lbl = new TextBlock { Text = label, FontSize = AppTheme.FontLG, FontFamily = Font, Foreground = allowed ? AppRes("AppFg") : AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Padding = new Thickness(8, 5) };
        Grid.SetColumn(lbl, 0);
        var toggle = BuildToggleSwitch(allowed, isOn => { AccessService.SetAccess(사번, key, isOn); lbl.Foreground = isOn ? AppRes("AppFg") : AppRes("FgMuted"); });
        Grid.SetColumn(toggle, 1);
        row.Children.Add(lbl); row.Children.Add(toggle);
        var wrap = new StackPanel { Spacing = 0 };
        wrap.Children.Add(row);
        wrap.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(8, 0) });
        return wrap;
    }

    private static Control BuildToggleSwitch(bool initialOn, Action<bool> onChange)
    {
        bool isOn = initialOn;
        var knob = new Border { Width = 18, Height = 18, CornerRadius = new CornerRadius(9), Background = Brushes.White, Margin = isOn ? new Thickness(22, 0, 0, 0) : new Thickness(2, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Left };
        var track = new Border { Width = 44, Height = 24, CornerRadius = new CornerRadius(12), Background = isOn ? new SolidColorBrush(Color.Parse("#2e7d32")) : new SolidColorBrush(Color.Parse("#555566")), Margin = new Thickness(8, 0), VerticalAlignment = VerticalAlignment.Center, Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand), Child = knob };
        track.PointerPressed += (_, _) =>
        {
            isOn = !isOn;
            track.Background = isOn ? new SolidColorBrush(Color.Parse("#2e7d32")) : new SolidColorBrush(Color.Parse("#555566"));
            knob.Margin = isOn ? new Thickness(22, 0, 0, 0) : new Thickness(2, 0, 0, 0);
            onChange(isOn);
        };
        return track;
    }

    // =========================================================================
    // 탭 2 — 분석항목 관리 Show1
    // =========================================================================
    private Control BuildAnalytePanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var hdrGrid = new Grid();
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        var hdrTitle = new TextBlock { Text = "분석정보 항목 목록", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"), FontFamily = Font, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 6) };
        Grid.SetColumn(hdrTitle, 0);
        var addBtn = new Button { Content = "+ 항목 추가", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusOkBg, Foreground = AppTheme.StatusOkFg, BorderBrush = AppTheme.StatusOkBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        addBtn.Click += (_, _) => { _selectedAnalyte = null; _show2Content.Content = BuildAnalyteEditPanel(null, isNew: true); };
        Grid.SetColumn(addBtn, 1);
        hdrGrid.Children.Add(hdrTitle); hdrGrid.Children.Add(addBtn);
        var headerBorder = new Border { Background = AppRes("PanelBg"), Child = hdrGrid };
        Grid.SetRow(headerBorder, 0);

        var treeScroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = _analyteTree };
        Grid.SetRow(treeScroll, 1);
        grid.Children.Add(headerBorder); grid.Children.Add(treeScroll);
        return grid;
    }

    private TreeView BuildAnalyteTreeView()
    {
        var tv = new TreeView { Margin = new Thickness(4) };
        tv.SelectionChanged += AnalyteTree_SelectionChanged;
        return tv;
    }

    private void LoadAnalyteTree()
    {
        _analyteTree.Items.Clear();
        try
        {
            var items  = AnalysisService.GetAllItems();
            // 그룹: 그룹 내 최소 ES 기준 → ES 번호 빠른 그룹부터.
            // 항목: 그룹 내에서 ES 번호 숫자 정렬 (자릿수/소수점 분리).
            var groups = items
                .GroupBy(a => string.IsNullOrWhiteSpace(a.Category) ? "기타" : a.Category)
                .OrderBy(g => g.Min(i => EsSortKey(i.ES)), StringComparer.Ordinal)
                .ThenBy(g => g.Key, StringComparer.CurrentCulture);
            foreach (var grp in groups)
            {
                var groupNode = new TreeViewItem
                {
                    Header = new TextBlock { Text = grp.Key, FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppTheme.FgInfo, Margin = new Thickness(2, 2) },
                    IsExpanded = true,
                };
                foreach (var item in grp.OrderBy(i => EsSortKey(i.ES), StringComparer.Ordinal))
                    groupNode.Items.Add(CreateAnalyteTreeItem(item));
                _analyteTree.Items.Add(groupNode);
            }
        }
        catch (Exception ex) { Console.WriteLine($"[LoadAnalyteTree] {ex.Message}"); }
    }

    /// <summary>"ES 04311.1d" 같은 ES 코드를 숫자 우선 정렬용 키로 변환 — 자릿수/소수점 분리, 빈 값 뒤로.</summary>
    private static string EsSortKey(string es)
    {
        if (string.IsNullOrWhiteSpace(es)) return "zzzz_";
        var m = System.Text.RegularExpressions.Regex.Match(es, @"(\d{4,5})(?:\.(\d+))?");
        if (!m.Success) return "zzz_" + es;
        var main = m.Groups[1].Value.PadLeft(5, '0');
        var sub  = (m.Groups[2].Success ? m.Groups[2].Value : "").PadLeft(3, '0');
        return $"{main}.{sub}_{es}";
    }

    private static TreeViewItem CreateAnalyteTreeItem(AnalysisItem item)
    {
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5, VerticalAlignment = VerticalAlignment.Center };
        if (!string.IsNullOrWhiteSpace(item.약칭))
        {
            var (bbg, bfg) = BadgeColorHelper.GetBadgeColor(item.약칭);
            row.Children.Add(new Border { Background = new SolidColorBrush(Color.Parse(bbg)), CornerRadius = new CornerRadius(3), Padding = new Thickness(4, 1), VerticalAlignment = VerticalAlignment.Center, Child = new TextBlock { Text = item.약칭, FontSize = AppTheme.FontXS, Foreground = new SolidColorBrush(Color.Parse(bfg)), FontFamily = Font } });
        }
        row.Children.Add(new TextBlock { Text = item.Analyte, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center });
        return new TreeViewItem { Header = row, Tag = item };
    }

    private void AnalyteTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0) return;
        if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is AnalysisItem item)
        {
            _selectedAnalyte = item;
            _show2Content.Content = BuildAnalyteEditPanel(item, isNew: false);
            // Show4 — 파서 컬럼 매핑 (해당 항목의 시험기록부 컬럼 + 스키마별 기본 매핑)
            try
            {
                var tableName = WaterCenterDbMigration.SafeName(item.Analyte) + "_시험기록부";
                var dbCols = new List<string>();
                using (var conn = DbConnectionFactory.CreateConnection())
                {
                    conn.Open();
                    if (DbConnectionFactory.TableExists(conn, tableName))
                        dbCols = DbConnectionFactory.GetColumnNames(conn, tableName);
                }
                _show4Content.Content = new ScrollViewer
                {
                    Content = BuildParserMapSection(item.Analyte, dbCols),
                    Padding = new Thickness(8),
                };
                Show4VisibleChanged?.Invoke(true);
            }
            catch (Exception ex)
            {
                _show4Content.Content = new TextBlock { Text = $"파서 매핑 오류: {ex.Message}", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = Brushes.Red };
                Show4VisibleChanged?.Invoke(true);
            }
        }
    }

    // =========================================================================
    // 탭 2 — 분석항목 관리 Show2
    // =========================================================================
    private Control BuildAnalyteEditPanel(AnalysisItem? item, bool isNew)
    {
        var origAnalyte = item?.Analyte;

        TextBox MkTb(string val, string hint) => new TextBox
        {
            Text = val, Watermark = hint, FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(6, 4), Margin = new Thickness(0, 2),
            Background = AppRes("PanelBg"), BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
        };

        var tbAnalyte  = MkTb(item?.Analyte ?? "",                    "항목명 (예: 생물학적산소요구량)");
        var tbAlias    = MkTb(item?.약칭 ?? "",                        "약칭 (예: BOD)");
        var tbAliasX   = MkTb(item?.AliasX ?? "",                     "파서 키워드 별칭 — 쉼표 구분 (예: T-N,TN,총 질소)");
        var tbCategory = MkTb(item?.Category ?? "",                    "카테고리 (예: 유기물류)");
        var tbUnit     = MkTb(item?.unit ?? "",                        "단위 (예: mg/L)");
        var tbDP       = MkTb((item?.DecimalPlaces ?? 0).ToString(),   "소수점 자리 (예: 1)");
        var tbES       = MkTb(item?.ES ?? "",                          "정렬번호 (예: 01)");
        var tbParts    = MkTb(item?.Parts ?? "",                       "구성 (예: 1)");
        var tbMethod   = MkTb(item?.Method ?? "",                      "시험방법 (예: 자외선/가시선분광법)");
        var tbInst     = MkTb(item?.instrument ?? "",                  "기기 (예: ICP-OES)");
        var tbLoQ      = MkTb(item?.정량한계?.ToString("0.##########") ?? "", "정량한계 (예: 0.05) — 결과 < LoQ 면 ND 표시");

        // ─ 기기 선택: 체크박스 드롭다운(다중 선택) ────────────────────
        // 현장측정 분석항목이면 현장측정장비 테이블 사용, 그 외는 전체 분석장비
        bool isFieldMeasAnalyteEarly =
            (item?.Category?.Contains("현장측정") == true) ||
            (item?.Analyte?.Contains("현장측정") == true);
        var equipments = new List<(string 장비명, string 코드값)>();
        try
        {
            equipments = isFieldMeasAnalyteEarly
                ? MeasurerService.GetFieldMeasEquipment()
                : MeasurerService.GetAllEquipments();
        }
        catch { }

        Control instField;
        if (equipments.Count > 0)
        {
            // 현재값 → 선택된 장비명 집합
            // 저장 포맷이 다양해도 매칭: 쉼표/슬래시/| 구분 토큰 + 대소문자/특수문자 무시 + 양방향 부분매칭
            static string NormEquip(string s)
                => new string((s ?? "").ToLowerInvariant()
                    .Where(char.IsLetterOrDigit).ToArray());

            var selected = new HashSet<string>();
            if (!string.IsNullOrWhiteSpace(item?.instrument))
            {
                var raw = item!.instrument!;
                var tokens = raw
                    .Split(new[] { ',', ';', '|', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Trim())
                    .Where(t => t.Length > 0)
                    .ToList();
                if (tokens.Count == 0) tokens.Add(raw.Trim());

                var normTokens = tokens.Select(NormEquip).Where(t => t.Length > 0).ToList();
                var normRawAll = NormEquip(raw);

                foreach (var (name, _) in equipments)
                {
                    var nName = NormEquip(name);
                    if (nName.Length == 0) continue;

                    bool hit = false;
                    // 1) 토큰별 양방향 부분매칭
                    foreach (var tok in normTokens)
                    {
                        if (tok.Length == 0) continue;
                        if (nName.Contains(tok) || tok.Contains(nName)) { hit = true; break; }
                    }
                    // 2) 전체 원문 안에 장비명(정규화)이 포함되는지
                    if (!hit && nName.Length >= 3 && normRawAll.Contains(nName)) hit = true;

                    if (hit) selected.Add(name);
                }
            }

            string SummaryText() => selected.Count == 0
                ? "장비 선택 ▾"
                : (selected.Count == 1 ? selected.First() : $"{selected.First()} 외 {selected.Count - 1}개") + " ▾";

            var dropdownBtn = new Button
            {
                Content                    = SummaryText(),
                FontSize                   = AppTheme.FontSM,
                FontFamily                 = Font,
                Height                     = 26,
                Padding                    = new Thickness(8, 0),
                Margin                     = new Thickness(0, 2, 0, 0),
                HorizontalAlignment        = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background                 = AppRes("PanelBg"),
                Foreground                 = AppRes("AppFg"),
                BorderBrush                = AppRes("InputBorder"),
                BorderThickness            = new Thickness(1),
            };

            var listPanel = new StackPanel { Spacing = 2, Margin = new Thickness(6) };
            foreach (var (name, _) in equipments)
            {
                var cb = new CheckBox
                {
                    Content    = name,
                    IsChecked  = selected.Contains(name),
                    FontSize   = AppTheme.FontSM,
                    FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                };
                var capturedName = name;
                cb.IsCheckedChanged += (_, _) =>
                {
                    if (cb.IsChecked == true) selected.Add(capturedName);
                    else                      selected.Remove(capturedName);
                    tbInst.Text        = string.Join(", ", selected);
                    dropdownBtn.Content = SummaryText();
                };
                listPanel.Children.Add(cb);
            }

            var sv = new ScrollViewer
            {
                MaxHeight                     = 320,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                Content                       = listPanel,
            };
            dropdownBtn.Flyout = new Flyout { Content = sv, Placement = PlacementMode.Bottom };

            // 초기 text 동기화 (최초 진입 시 빈 문자열 보정)
            tbInst.Text = string.Join(", ", selected);

            instField = dropdownBtn;
        }
        else
        {
            instField = tbInst;
        }

        // ─ 폼 레이아웃 ────────────────────────────────────────────────
        Grid TwoCol(Control l, Control r)
        {
            var g = new Grid();
            g.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
            g.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(8)));
            g.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
            Grid.SetColumn(l, 0); Grid.SetColumn(r, 2);
            g.Children.Add(l); g.Children.Add(r);
            return g;
        }
        Control Labeled(string lbl, Control ctrl)
        {
            var s = new StackPanel { Spacing = 0 };
            s.Children.Add(new TextBlock { Text = lbl, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 4, 0, 1) });
            s.Children.Add(ctrl);
            return s;
        }

        var form = new StackPanel { Spacing = 2, Margin = new Thickness(0, 6) };
        form.Children.Add(Labeled("항목명 (Analyte)", tbAnalyte));
        form.Children.Add(TwoCol(Labeled("약칭", tbAlias), Labeled("카테고리", tbCategory)));
        form.Children.Add(Labeled("파서 별칭 (AliasX) — 쉼표 구분, 파서 키워드 자동 매핑에 사용", tbAliasX));
        form.Children.Add(TwoCol(Labeled("단위", tbUnit), Labeled("소수점 자리", tbDP)));
        form.Children.Add(TwoCol(Labeled("정렬번호 (ES)", tbES), Labeled("구성 (Parts)", tbParts)));
        form.Children.Add(Labeled("정량한계 (LoQ) — 결과 < LoQ 면 시험기록부에 ND 표시", tbLoQ));
        form.Children.Add(Labeled("시험방법", tbMethod));
        form.Children.Add(Labeled("기기 (instrument)", instField));

        // ─ 동시분석그룹 (번호 같은 항목끼리 한번에 분석) ────────────────────
        int curGroup = 0;
        if (!string.IsNullOrWhiteSpace(origAnalyte))
        {
            try { curGroup = MeasurerService.GetConcurrentGroup(origAnalyte!); } catch { }
        }
        var cbGroup = new ComboBox
        {
            FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 28, Margin = new Thickness(0, 2, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Background = AppRes("PanelBg"),
            BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
            Tag = curGroup,
        };
        // 각 그룹에 속한 항목 목록 조회
        var groupMembers = new Dictionary<int, List<string>>();
        try
        {
            var all = MeasurerService.GetAllConcurrentGroups();
            foreach (var kv in all)
            {
                if (kv.Value < 1 || kv.Value > 10) continue;
                if (!groupMembers.ContainsKey(kv.Value)) groupMembers[kv.Value] = new List<string>();
                groupMembers[kv.Value].Add(kv.Key);
            }
        }
        catch (Exception ex) { Console.WriteLine($"[GetAllConcurrentGroups] {ex.Message}"); }

        cbGroup.Items.Add("(없음)");
        for (int g = 1; g <= 10; g++)
        {
            string label;
            if (groupMembers.TryGetValue(g, out var members) && members.Count > 0)
            {
                var joined = string.Join(", ", members.OrderBy(s => s));
                if (joined.Length > 60) joined = joined.Substring(0, 57) + "…";
                label = $"그룹 {g} — {joined} ({members.Count})";
            }
            else
            {
                label = $"그룹 {g} (비어있음)";
            }
            cbGroup.Items.Add(label);
        }
        cbGroup.SelectedIndex = Math.Max(0, Math.Min(10, curGroup));
        form.Children.Add(Labeled(
            "동시분석그룹 (같이 분석되는 항목끼리 같은 번호 — 예: 금속류 1, 음이온 2)",
            cbGroup));

        // ─ 버튼 행 ───────────────────────────────────────────────────
        var btnSave = new Button { Content = isNew ? "✚ 신규 저장" : "💾 저장", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        btnSave.Click += (_, _) =>
        {
            var updated = new AnalysisItem
            {
                Category      = tbCategory.Text?.Trim() ?? "",
                Analyte       = tbAnalyte.Text?.Trim() ?? "",
                약칭           = tbAlias.Text?.Trim() ?? "",
                AliasX        = tbAliasX.Text?.Trim() ?? "",
                Parts         = tbParts.Text?.Trim() ?? "",
                DecimalPlaces = int.TryParse(tbDP.Text?.Trim(), out var dp) ? dp : 0,
                unit          = tbUnit.Text?.Trim() ?? "",
                ES            = tbES.Text?.Trim() ?? "",
                Method        = tbMethod.Text?.Trim() ?? "",
                instrument    = tbInst.Text?.Trim() ?? "",
                정량한계      = double.TryParse(tbLoQ.Text?.Trim(), out var loq) ? loq : (double?)null,
            };
            if (string.IsNullOrWhiteSpace(updated.Analyte)) return;
            SaveAnalyteToDB(updated, isNew ? null : origAnalyte);

            // 동시분석그룹 저장 (SelectedIndex 0 = 없음, 1~ = 그룹번호)
            try { MeasurerService.SaveConcurrentGroup(updated.Analyte, cbGroup.SelectedIndex); }
            catch (Exception ex) { Console.WriteLine($"[SaveConcurrentGroup] {ex.Message}"); }

            LoadAnalyteTree();
            _selectedAnalyte = updated;
            _show2Content.Content = BuildAnalyteEditPanel(updated, isNew: false);
        };

        var btnDelete = new Button { Content = "🗑 삭제", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = AppTheme.StatusBadBg, Foreground = AppTheme.StatusBadFg, BorderBrush = AppTheme.StatusBadBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold, IsVisible = !isNew };
        btnDelete.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(origAnalyte)) return;
            DeleteAnalyteFromDB(origAnalyte);
            LoadAnalyteTree();
            _selectedAnalyte = null;
            _show2Content.Content = BuildAnalyteHint();
        };

        var btnCreateTable = new Button { Content = "🗄 시험기록부 테이블 생성", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = AppTheme.StatusAccentBg, Foreground = AppTheme.StatusAccentFg, BorderBrush = AppTheme.StatusAccentBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        var statusText = new TextBlock { FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 4), IsVisible = false };
        btnCreateTable.Click += (_, _) =>
        {
            try
            {
                FacilityDbMigration.EnsureAnalysisRecordTables();
                statusText.Text      = "✓ 시험기록부 테이블 생성 완료";
                statusText.Foreground = new SolidColorBrush(Color.Parse("#4caf50"));
                statusText.IsVisible  = true;
                var cur = new AnalysisItem { Category = tbCategory.Text?.Trim() ?? "", Analyte = tbAnalyte.Text?.Trim() ?? "", 약칭 = tbAlias.Text?.Trim() ?? "", Method = tbMethod.Text?.Trim() ?? "", instrument = tbInst.Text?.Trim() ?? "" };
                if (!string.IsNullOrWhiteSpace(cur.Analyte))
                    _show2Content.Content = BuildAnalyteEditPanel(cur, isNew: false);
            }
            catch (Exception ex) { statusText.Text = $"오류: {ex.Message}"; statusText.Foreground = Brushes.Red; statusText.IsVisible = true; }
        };

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0, 10, 0, 4) };
        btnRow.Children.Add(btnSave);
        if (!isNew) btnRow.Children.Add(btnDelete);
        btnRow.Children.Add(btnCreateTable);

        // ─ 시험기록부 컬럼 섹션 ─────────────────────────────────────
        var colSection = BuildColumnHeadersSection(item);

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock { Text = isNew ? "✚ 새 분석항목 추가" : $"분석항목 편집 — {item?.Analyte}", FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 4) });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 6) });
        root.Children.Add(form);
        root.Children.Add(btnRow);
        root.Children.Add(statusText);
        if (!isNew) { root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0, 10) }); root.Children.Add(colSection); }
        return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = root };
    }

    // 시험기록부 컬럼 헤더 섹션 ─────────────────────────────────────────────
    private Control BuildColumnHeadersSection(AnalysisItem? item)
    {
        var stack = new StackPanel { Spacing = 4 };
        stack.Children.Add(new TextBlock { Text = "📋 시험기록부 컬럼", FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppTheme.FgInfo, Margin = new Thickness(0, 0, 0, 4) });
        if (item == null) { stack.Children.Add(new TextBlock { Text = "항목을 선택하세요.", FontFamily = Font, FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted") }); return stack; }
        try
        {
            string autoSchema = WaterCenterDbMigration.DetermineSchema(item.Analyte, item.Category, item.Method, item.instrument);
            string? overrideSchema = ETA.Services.SERVICE3.AnalysisNoteService.GetSchemaOverride(item.Analyte);
            string schema   = overrideSchema ?? autoSchema;
            string safeName = WaterCenterDbMigration.SafeName(item.Analyte);
            bool   isToc    = item.Analyte.Equals("TOC", StringComparison.OrdinalIgnoreCase);
            string displayTable = isToc ? "TOC_NPOC_시험기록부 / TOC_TCIC_시험기록부" : $"{safeName}_시험기록부";

            (string Bg, string Fg) SchemaColor(string s) => s switch
            {
                "BOD"     => ("#1b5e20", "#a5d6a7"), "SS"      => ("#4a148c", "#ce93d8"),
                "UvVis"   => ("#0d47a1", "#90caf9"), "TocTcic" => ("#006064", "#80deea"),
                "GcLc"    => ("#e65100", "#ffcc80"), "Metal"   => ("#bf360c", "#ffab91"),
                "Cr6"     => ("#880e4f", "#f48fb1"), "VOC"     => ("#33691e", "#dcedc8"),
                "NHexan"  => ("#4e342e", "#ffccbc"), "Ecotox"  => ("#1a237e", "#9fa8da"),
                _         => ("#424242", "#eeeeee"),
            };

            var (schemaBg, schemaFg) = SchemaColor(schema);

            // 스키마 드롭다운 패널 (처음엔 숨김)
            var schemaDropdown = new StackPanel { IsVisible = false, Spacing = 2, Margin = new Thickness(0, 4, 0, 0) };
            var schemaBadgesWrap = new WrapPanel { Orientation = Orientation.Horizontal };
            schemaDropdown.Children.Add(schemaBadgesWrap);
            var capturedAnalyte = item.Analyte;
            // DB에서 실제 사용 중인 표준형식 목록 동적 조회
            var dbSchemas = ETA.Services.SERVICE3.AnalysisNoteService.GetDistinctSchemas();
            // 수동 추가 스키마가 DB 목록에 없으면 맨 앞에 표시
            var schemaList = overrideSchema != null && !dbSchemas.Contains(overrideSchema, StringComparer.OrdinalIgnoreCase)
                ? (IEnumerable<string>)new[] { overrideSchema }.Concat(dbSchemas)
                : (IEnumerable<string>)dbSchemas;
            foreach (var s in schemaList)
            {
                var sc = s;
                var (sbg, sfg) = SchemaColor(sc);
                bool isCurrent = sc == schema;
                var optBorder = new Border
                {
                    Background   = new SolidColorBrush(Color.Parse(sbg)),
                    CornerRadius = new CornerRadius(4),
                    Padding      = new Thickness(8, 3),
                    Margin       = new Thickness(0, 2, 4, 2),
                    Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Opacity      = isCurrent ? 1.0 : 0.55,
                    Child        = new TextBlock
                    {
                        Text       = isCurrent ? $"✓ {sc}" : sc,
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = new SolidColorBrush(Color.Parse(sfg)),
                        FontWeight = isCurrent ? FontWeight.Bold : FontWeight.Normal,
                    }
                };
                optBorder.PointerReleased += (_, _) =>
                {
                    ETA.Services.SERVICE3.AnalysisNoteService.SaveSchemaOverride(capturedAnalyte, sc == autoSchema ? null : sc);
                    _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                };
                schemaBadgesWrap.Children.Add(optBorder);
            }
            // 자동(초기화) 버튼
            if (overrideSchema != null)
            {
                var resetBorder = new Border
                {
                    Background   = AppTheme.BgSecondary,
                    CornerRadius = new CornerRadius(4),
                    Padding      = new Thickness(8, 3),
                    Margin       = new Thickness(0, 2, 4, 2),
                    Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Child        = new TextBlock { Text = "↺ 자동", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("FgMuted") }
                };
                resetBorder.PointerReleased += (_, _) =>
                {
                    ETA.Services.SERVICE3.AnalysisNoteService.SaveSchemaOverride(capturedAnalyte, null);
                    _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                };
                schemaBadgesWrap.Children.Add(resetBorder);
            }

            // 직접 입력 행 (구분선 + TextBox + 적용 버튼)
            schemaDropdown.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 4, 0, 4) });
            var customRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };

            var tbCustom = new TextBox
            {
                Watermark = "직접 입력 (예: NHexan2, Custom...)",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Width = 220, Padding = new Thickness(6, 3),
            };
            var addSchemaBtn = new Border
            {
                Background   = new SolidColorBrush(Color.Parse("#1a3a5c")),
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(10, 3),
                Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child        = new TextBlock { Text = "적용", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#90caf9")) }
            };
            addSchemaBtn.PointerReleased += (_, _) =>
            {
                var custom = tbCustom.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(custom)) return;
                ETA.Services.SERVICE3.AnalysisNoteService.SaveSchemaOverride(capturedAnalyte, custom);
                _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
            };
            // Enter 키로도 적용
            tbCustom.KeyDown += (_, e) =>
            {
                if (e.Key != Avalonia.Input.Key.Enter) return;
                var custom = tbCustom.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(custom)) return;
                ETA.Services.SERVICE3.AnalysisNoteService.SaveSchemaOverride(capturedAnalyte, custom);
                _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
            };
            customRow.Children.Add(tbCustom);
            customRow.Children.Add(addSchemaBtn);
            schemaDropdown.Children.Add(customRow);

            // 스키마 배지 (클릭 → 드롭다운 토글)
            var schemaBadgeText = new TextBlock
            {
                Text       = overrideSchema != null ? $"{schema} ✎" : $"{schema} ▾",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = new SolidColorBrush(Color.Parse(schemaFg)),
                FontWeight = FontWeight.SemiBold,
            };
            var schemaBadge = new Border
            {
                Background   = new SolidColorBrush(Color.Parse(schemaBg)),
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(8, 3),
                Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child        = schemaBadgeText,
            };
            schemaBadge.PointerReleased += (_, _) => schemaDropdown.IsVisible = !schemaDropdown.IsVisible;

            var schemaRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 0, 0, 4) };
            schemaRow.Children.Add(new TextBlock { Text = "표준형식:", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
            schemaRow.Children.Add(schemaBadge);
            if (overrideSchema != null)
                schemaRow.Children.Add(new TextBlock { Text = $"(자동: {autoSchema})", FontFamily = Font, FontSize = AppTheme.FontXS, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
            stack.Children.Add(schemaRow);
            stack.Children.Add(schemaDropdown);
            stack.Children.Add(new TextBlock { Text = $"테이블: {displayTable}", FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 6), TextWrapping = TextWrapping.Wrap });

            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            var actualTable = isToc ? "TOC_NPOC_시험기록부" : $"{safeName}_시험기록부";
            var skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "id", "분석일", "SN", "업체명", "시료명", "구분", "소스구분", "비고", "등록일시" };
            if (DbConnectionFactory.TableExists(conn, actualTable))
            {
                var cols = DbConnectionFactory.GetColumnNames(conn, actualTable).Where(c => !skip.Contains(c)).ToList();

                // 노트 컬럼 선택 UI
                var selectedCols = new HashSet<string>(
                    ETA.Services.SERVICE3.AnalysisNoteService.GetNoteColumns(item.Analyte),
                    StringComparer.OrdinalIgnoreCase);

                // 노트 포함 컬럼 헤더 행 (레이블 + 편집 토글 버튼)
                var noteColHeader = new Grid { Margin = new Thickness(0, 4, 0, 2) };
                noteColHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Star });
                noteColHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                var noteColLabel = new TextBlock
                {
                    Text = "노트 포함 컬럼 (클릭으로 선택/해제)",
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(noteColLabel, 0);
                var toggleEditBtn = new Border
                {
                    Background = _colEditMode
                        ? new SolidColorBrush(Color.Parse("#1a3a5c"))
                        : AppTheme.BgSecondary,
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 3),
                    Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Child = new TextBlock
                    {
                        Text = _colEditMode ? "✕ 편집 닫기" : "✎ 컬럼 편집",
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = _colEditMode
                            ? new SolidColorBrush(Color.Parse("#90caf9"))
                            : AppRes("AppFg")
                    }
                };
                toggleEditBtn.PointerReleased += (_, _) =>
                {
                    _colEditMode = !_colEditMode;
                    _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                };
                Grid.SetColumn(toggleEditBtn, 1);
                noteColHeader.Children.Add(noteColLabel);
                noteColHeader.Children.Add(toggleEditBtn);
                stack.Children.Add(noteColHeader);

                var sampleCols = cols.Where(c => !AnalysisNoteService.IsCalibrationCol(c)).ToList();
                var calCols    = cols.Where(AnalysisNoteService.IsCalibrationCol).ToList();

                void AddColBadges(System.Collections.Generic.IList<string> colGroup, Avalonia.Controls.WrapPanel wp)
                {
                    foreach (var col in colGroup)
                    {
                        var colLocal = col;
                        bool isSelected = selectedCols.Contains(colLocal);

                        var tb = new TextBlock
                        {
                            Text = colLocal, FontFamily = Font, FontSize = AppTheme.FontSM,
                            Foreground = isSelected
                                ? new SolidColorBrush(Color.Parse("#a5d6a7"))
                                : AppRes("AppFg")
                        };
                        var badge = new Border
                        {
                            Background   = isSelected
                                ? new SolidColorBrush(Color.Parse("#1b5e20"))
                                : AppTheme.BgSecondary,
                            CornerRadius = new CornerRadius(4),
                            Padding      = new Thickness(8, 4),
                            Margin       = new Thickness(0, 2, 4, 2),
                            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                            Child        = tb
                        };
                        var dragged = new bool[1];
                        badge.PointerPressed += (_, _) => dragged[0] = false;
                        badge.PointerMoved += async (_, e) =>
                        {
                            if (dragged[0]) return;
                            if (!e.GetCurrentPoint(badge).Properties.IsLeftButtonPressed) return;
                            dragged[0] = true;
                            var dObj = new DataObject();
                            dObj.Set("formula_col", colLocal);
                            await DragDrop.DoDragDrop(e, dObj, DragDropEffects.Copy);
                        };
                        badge.PointerReleased += (_, _) =>
                        {
                            if (dragged[0]) return;
                            if (selectedCols.Contains(colLocal)) selectedCols.Remove(colLocal);
                            else selectedCols.Add(colLocal);
                            bool sel = selectedCols.Contains(colLocal);
                            badge.Background = sel
                                ? new SolidColorBrush(Color.Parse("#1b5e20"))
                                : AppTheme.BgSecondary;
                            tb.Foreground = sel
                                ? new SolidColorBrush(Color.Parse("#a5d6a7"))
                                : AppRes("AppFg");
                            ETA.Services.SERVICE3.AnalysisNoteService.SaveNoteColumns(item.Analyte, selectedCols);
                        };
                        wp.Children.Add(badge);
                    }
                }

                // 시료 측정 컬럼
                if (sampleCols.Count > 0)
                {
                    stack.Children.Add(new TextBlock
                    {
                        Text = "🧪 시료 측정 컬럼", FontFamily = Font, FontSize = AppTheme.FontXS,
                        FontWeight = Avalonia.Media.FontWeight.SemiBold,
                        Foreground = new SolidColorBrush(Color.Parse("#a5d6a7")),
                        Margin = new Thickness(0, 6, 0, 2)
                    });
                    var sampleWrap = new WrapPanel { Orientation = Orientation.Horizontal };
                    AddColBadges(sampleCols, sampleWrap);
                    stack.Children.Add(sampleWrap);
                }

                // 검정곡선 컬럼
                if (calCols.Count > 0)
                {
                    stack.Children.Add(new TextBlock
                    {
                        Text = "📈 검정곡선 컬럼", FontFamily = Font, FontSize = AppTheme.FontXS,
                        FontWeight = Avalonia.Media.FontWeight.SemiBold,
                        Foreground = new SolidColorBrush(Color.Parse("#90caf9")),
                        Margin = new Thickness(0, 8, 0, 2)
                    });
                    var calWrap = new WrapPanel { Orientation = Orientation.Horizontal };
                    AddColBadges(calCols, calWrap);
                    stack.Children.Add(calWrap);
                }

                // 컬럼 편집 모드 UI
                if (_colEditMode)
                {
                    var editSection = new Border
                    {
                        Background = new SolidColorBrush(Color.Parse("#0d1f33")),
                        CornerRadius = new CornerRadius(6),
                        Padding = new Thickness(10),
                        Margin = new Thickness(0, 6, 0, 0)
                    };
                    var editStack = new StackPanel { Spacing = 4 };
                    editStack.Children.Add(new TextBlock
                    {
                        Text = "컬럼 편집",
                        FontFamily = Font, FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
                        Foreground = new SolidColorBrush(Color.Parse("#90caf9")),
                        Margin = new Thickness(0, 0, 0, 6)
                    });

                    foreach (var colItem in cols)
                    {
                        var colCapture = colItem;
                        var row = new Grid { Margin = new Thickness(0, 0, 0, 4) };
                        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
                        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Star });
                        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                        var colLabel = new TextBlock
                        {
                            Text = colCapture, FontFamily = Font, FontSize = AppTheme.FontSM,
                            Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center
                        };
                        Grid.SetColumn(colLabel, 0);

                        var renameBox = new TextBox
                        {
                            Text = colCapture, FontFamily = Font, FontSize = AppTheme.FontSM,
                            Margin = new Thickness(4, 0), Padding = new Thickness(6, 3)
                        };
                        Grid.SetColumn(renameBox, 1);

                        var renameBtn = new Border
                        {
                            Background = new SolidColorBrush(Color.Parse("#1a3a1a")),
                            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 3),
                            Margin = new Thickness(4, 0),
                            Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                            Child = new TextBlock { Text = "이름변경", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#a5d6a7")) }
                        };
                        renameBtn.PointerReleased += (_, _) =>
                        {
                            var newName = renameBox.Text?.Trim() ?? "";
                            if (string.IsNullOrEmpty(newName) || newName == colCapture) return;
                            try
                            {
                                using var c = DbConnectionFactory.CreateConnection(); c.Open();
                                // INFORMATION_SCHEMA로 컬럼 타입 조회 (CHANGE COLUMN은 전 버전 호환)
                                string colDef = "TEXT DEFAULT ''";
                                using (var typeCmd = c.CreateCommand())
                                {
                                    typeCmd.CommandText =
                                        "SELECT COLUMN_TYPE, IS_NULLABLE, COLUMN_DEFAULT " +
                                        "FROM INFORMATION_SCHEMA.COLUMNS " +
                                        "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = @t AND COLUMN_NAME = @cn";
                                    typeCmd.Parameters.AddWithValue("@t",  actualTable);
                                    typeCmd.Parameters.AddWithValue("@cn", colCapture);
                                    using var tr = typeCmd.ExecuteReader();
                                    if (tr.Read())
                                    {
                                        string colType  = tr.IsDBNull(0) ? "text"  : tr.GetString(0);
                                        string nullable = tr.IsDBNull(1) ? "YES"   : tr.GetString(1);
                                        string? defVal  = tr.IsDBNull(2) ? null    : tr.GetString(2);
                                        string nullStr  = nullable == "NO" ? "NOT NULL" : "NULL";
                                        string defStr   = defVal != null ? $" DEFAULT '{defVal}'" : " DEFAULT ''";
                                        colDef = $"{colType} {nullStr}{defStr}";
                                    }
                                }
                                using var cmd2 = c.CreateCommand();
                                cmd2.CommandText = $"ALTER TABLE `{actualTable}` CHANGE COLUMN `{colCapture}` `{newName}` {colDef}";
                                cmd2.ExecuteNonQuery();
                                _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                            }
                            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[ColEdit] Rename: {ex.Message}"); }
                        };
                        Grid.SetColumn(renameBtn, 2);

                        var deleteBtn = new Border
                        {
                            Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 3),
                            Margin = new Thickness(4, 0),
                            Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                            Child = new TextBlock { Text = "삭제", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#ef9a9a")) }
                        };
                        deleteBtn.PointerReleased += (_, _) =>
                        {
                            try
                            {
                                using var c = DbConnectionFactory.CreateConnection(); c.Open();
                                using var cmd2 = c.CreateCommand();
                                cmd2.CommandText = $"ALTER TABLE `{actualTable}` DROP COLUMN `{colCapture}`";
                                cmd2.ExecuteNonQuery();
                                _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                            }
                            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[ColEdit] Drop: {ex.Message}"); }
                        };
                        Grid.SetColumn(deleteBtn, 3);

                        row.Children.Add(colLabel); row.Children.Add(renameBox);
                        row.Children.Add(renameBtn); row.Children.Add(deleteBtn);
                        editStack.Children.Add(row);
                    }

                    // 새 컬럼 추가 행
                    editStack.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 4, 0, 4) });
                    var addRow = new Grid();
                    addRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
                    addRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Star });
                    addRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    var addLabel = new TextBlock
                    {
                        Text = "+ 새 컬럼", FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetColumn(addLabel, 0);

                    var newColBox = new TextBox
                    {
                        Watermark = "새 컬럼명", FontFamily = Font, FontSize = AppTheme.FontSM,
                        Margin = new Thickness(4, 0), Padding = new Thickness(6, 3)
                    };
                    Grid.SetColumn(newColBox, 1);

                    var addColBtn = new Border
                    {
                        Background = new SolidColorBrush(Color.Parse("#1a3a5c")),
                        CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 3),
                        Margin = new Thickness(4, 0),
                        Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                        Child = new TextBlock { Text = "추가", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#90caf9")) }
                    };
                    addColBtn.PointerReleased += (_, _) =>
                    {
                        var newName = newColBox.Text?.Trim() ?? "";
                        if (string.IsNullOrEmpty(newName)) return;
                        try
                        {
                            using var c = DbConnectionFactory.CreateConnection(); c.Open();
                            using var cmd2 = c.CreateCommand();
                            cmd2.CommandText = $"ALTER TABLE `{actualTable}` ADD COLUMN `{newName}` TEXT DEFAULT ''";
                            cmd2.ExecuteNonQuery();
                            _show2Content.Content = BuildAnalyteEditPanel(_selectedAnalyte, isNew: false);
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[ColEdit] Add: {ex.Message}"); }
                    };
                    Grid.SetColumn(addColBtn, 2);

                    addRow.Children.Add(addLabel); addRow.Children.Add(newColBox); addRow.Children.Add(addColBtn);
                    editStack.Children.Add(addRow);

                    editSection.Child = editStack;
                    stack.Children.Add(editSection);
                }

                // 농도 계산 공식 빌더
                stack.Children.Add(BuildFormulaSection(item.Analyte, cols));
                // 파서 컬럼 매핑은 Show4 로 이동 (좁은 Show2 가독성 확보)
            }
            else
                stack.Children.Add(new TextBlock { Text = "⚠ 시험기록부 테이블 없음 — 위 버튼으로 생성하세요", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#f57f17")), TextWrapping = TextWrapping.Wrap });
        }
        catch (Exception ex) { stack.Children.Add(new TextBlock { Text = $"오류: {ex.Message}", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = Brushes.Red }); }
        return stack;
    }

    /// <summary>스키마별 기본 파서→DB 매핑 — 기존 Upsert*Data 의 하드코딩 룰을 그대로 노출.
    ///   사용자는 이 표시를 보고 잘못 매핑됐으면 우측 콤보로 바꾸면 끝.</summary>
    private static Dictionary<string, string> GetDefaultParserMap(string schema)
    {
        var d = new Dictionary<string, string>(StringComparer.Ordinal);
        switch (schema?.ToLowerInvariant())
        {
            case "uvvis":
            case "metal":
            case "cr6":
                d["흡광도"]   = "흡광도";
                d["시료량"]   = "시료량";
                d["희석배수"] = "희석배수";
                d["농도"]     = "농도";
                d["결과"]     = "결과";
                break;
            case "voc":
            case "gclc":
                d["Area"]     = "Area";
                d["ISTD"]     = "ISTD";
                d["시료량"]   = "시료량";
                d["희석배수"] = "희석배수";
                d["농도"]     = "농도";
                d["결과"]     = "결과";
                break;
            case "ss":
                d["시료량"]   = "시료량";
                d["전무게"]   = "전무게";
                d["후무게"]   = "후무게";
                d["무게차"]   = "무게차";
                d["희석배수"] = "희석배수";
                d["결과"]     = "결과";
                break;
            case "nhexan":
                d["시료량"]   = "시료량";
                d["전무게"]   = "전무게";
                d["후무게"]   = "후무게";
                d["무게차"]   = "무게차";
                d["희석배수"] = "희석배수";
                d["결과"]     = "결과";
                break;
            case "bod":
                d["시료량"]       = "시료량";
                d["D1"]           = "D1";
                d["D2"]           = "D2";
                d["희석배수"]     = "희석배수";
                d["결과"]         = "결과";
                d["식종시료량"]   = "식종시료량";
                d["식종BOD"]      = "식종BOD";
                d["15min_DO"]     = "15min_DO";
                d["5Day_DO"]      = "5Day_DO";
                break;
            case "toctcic":
                d["TCAU"]     = "TCAU";
                d["TCcon"]    = "TCcon";
                d["ICAU"]     = "ICAU";
                d["ICcon"]    = "ICcon";
                d["희석배수"] = "희석배수";
                d["농도"]     = "농도";
                d["결과"]     = "결과";
                break;
            case "ecotox":
                d["LC50"]            = "LC50";
                d["TU"]              = "TU";
                d["대조군_생물수"]   = "대조군_생물수";
                d["대조군_사망수"]   = "대조군_사망수";
                break;
            case "coliform":
                d["A"]            = "A";
                d["B"]            = "B";
                d["시료량"]       = "시료량";
                d["희석배수"]     = "희석배수";
                d["사용희석배수"] = "사용희석배수";
                d["결과"]         = "결과";
                break;
            default:
                d["시료량"]   = "시료량";
                d["희석배수"] = "희석배수";
                d["결과"]     = "결과";
                break;
        }
        return d;
    }

    // ── 파서 출력 필드 후보 (각 파서 코드에서 추출되는 대표 필드들) ────────
    private static readonly List<string> _parserFieldCandidates = new()
    {
        // 공통
        "SN","시료명","분석일",
        // GC/GCMS/PFAS
        "Area","Resp.","Response","ISTD","ISTD Resp.","ISTD Area","CompoundName","성분명",
        // UV-VIS / 분광광도계
        "흡광도","Abs","Conc.","농도","계산농도",
        // TOC NPOC/TCIC
        "AU","TC AU","IC AU","TC Conc.","IC Conc.","TCAU","ICAU","TCcon","ICcon",
        // BOD
        "D1","D2","15min_DO","5Day_DO","f(x/y)","식종시료량","식종BOD","식종함유량",
        // SS / NHexan (중량법)
        "전무게","후무게","무게차",
        // 공통 측정 메타
        "시료량","희석배수","결과","결과값",
        // 검정곡선
        "기울기","절편","R²","R2","검량선_a","검량선_b",
    };

    // ── 파서 컬럼 매핑 빌더 (드롭다운) ────────────────────────────────────────
    //   좌측: 파서 필드 후보 (편집 가능)
    //   우측: 현재 분석항목의 시험기록부 DB 컬럼 (편집 가능)
    //   분석정보.parser_column_map (JSON) 에 저장 → UpsertRowData 가 참조.
    private Control BuildParserMapSection(string analyte, List<string> availableCols)
    {
        var section = new StackPanel { Spacing = 4, Margin = new Thickness(0, 12, 0, 0) };
        section.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 0, 0, 4) });

        section.Children.Add(new TextBlock
        {
            Text = "파서 컬럼 매핑  (파서 필드 → DB 컬럼)",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = AppRes("FgMuted"),
            Margin = new Thickness(0, 0, 0, 4),
        });

        // 후보 목록 — DB 컬럼은 항목 스키마, 파서 필드는 공통 후보
        var dbCandidates = new List<string>(availableCols.Where(c => !string.IsNullOrWhiteSpace(c)));
        // 누락 흔한 컬럼 보강
        foreach (var must in new[] { "Area", "ISTD", "농도", "희석배수", "시료량", "결과", "흡광도" })
            if (!dbCandidates.Contains(must, StringComparer.OrdinalIgnoreCase)) dbCandidates.Add(must);

        var listPanel = new StackPanel { Spacing = 3 };
        section.Children.Add(listPanel);

        var saved = ETA.Services.SERVICE3.AnalysisNoteService.GetParserColumnMap(analyte)
                    ?? new Dictionary<string, string>();
        // 스키마별 기본 매핑 — Upsert*Data 의 기존 하드코딩 룰 그대로 노출
        var schema = ETA.Services.SERVICE3.AnalysisNoteService.GetSchemaOverride(analyte) ?? "";
        if (string.IsNullOrEmpty(schema))
            schema = WaterCenterDbMigration.DetermineSchema(
                _selectedAnalyte?.Analyte ?? analyte,
                _selectedAnalyte?.Category ?? "",
                _selectedAnalyte?.Method ?? "",
                _selectedAnalyte?.instrument ?? "");
        var defaults = GetDefaultParserMap(schema);
        // 기본값 + 저장된 값 병합 — 저장된 값 우선
        var current = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var kv in defaults) current[kv.Key] = kv.Value;
        foreach (var kv in saved)    current[kv.Key] = kv.Value;

        void SaveMap()
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var child in listPanel.Children)
            {
                if (child is Grid g && g.Children.Count >= 2 &&
                    g.Children[0] is ComboBox kCb && g.Children[1] is ComboBox vCb)
                {
                    var k = (kCb.SelectedItem?.ToString() ?? kCb.Text ?? "").Trim();
                    var v = (vCb.SelectedItem?.ToString() ?? vCb.Text ?? "").Trim();
                    if (!string.IsNullOrEmpty(k) && !string.IsNullOrEmpty(v))
                        map[k] = v;
                }
            }
            ETA.Services.SERVICE3.AnalysisNoteService.SaveParserColumnMap(analyte, map);
        }

        Grid AddRow(string parserField = "", string dbCol = "")
        {
            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,Auto,*,Auto"),
            };
            var keyCb = new ComboBox
            {
                ItemsSource = _parserFieldCandidates,
                SelectedItem = string.IsNullOrEmpty(parserField) ? null : parserField,
                FontFamily = Font, FontSize = AppTheme.FontSM,
                MinHeight = 28,
                PlaceholderText = "파서 필드 선택…",
            };
            keyCb.SelectionChanged += (_, _) => SaveMap();
            var arrow = new TextBlock
            {
                Text = " → ", FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            };
            var valCb = new ComboBox
            {
                ItemsSource = dbCandidates,
                SelectedItem = string.IsNullOrEmpty(dbCol) ? null : dbCol,
                FontFamily = Font, FontSize = AppTheme.FontSM,
                MinHeight = 28,
                PlaceholderText = "DB 컬럼 선택…",
            };
            valCb.SelectionChanged += (_, _) => SaveMap();
            var delBtn = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#4a0000")),
                CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 2),
                Margin = new Thickness(4, 0, 0, 0),
                Cursor = new Cursor(StandardCursorType.Hand),
                Child = new TextBlock { Text = "✕", FontFamily = Font, FontSize = AppTheme.FontXS, Foreground = new SolidColorBrush(Color.Parse("#ef9a9a")) },
                VerticalAlignment = VerticalAlignment.Center,
            };
            delBtn.PointerPressed += (_, _) => { listPanel.Children.Remove(rowGrid); SaveMap(); };
            Grid.SetColumn(keyCb, 0); rowGrid.Children.Add(keyCb);
            Grid.SetColumn(arrow, 1); rowGrid.Children.Add(arrow);
            Grid.SetColumn(valCb, 2); rowGrid.Children.Add(valCb);
            Grid.SetColumn(delBtn, 3); rowGrid.Children.Add(delBtn);
            listPanel.Children.Add(rowGrid);
            return rowGrid;
        }

        foreach (var kv in current) AddRow(kv.Key, kv.Value);

        var addBtn = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1a3a5c")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 4),
            Margin = new Thickness(0, 4, 0, 0),
            Cursor = new Cursor(StandardCursorType.Hand),
            HorizontalAlignment = HorizontalAlignment.Left,
            Child = new TextBlock { Text = "＋ 매핑 추가", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#90caf9")) },
        };
        addBtn.PointerPressed += (_, _) => AddRow();
        section.Children.Add(addBtn);

        section.Children.Add(new TextBlock
        {
            Text = $"DB 컬럼 후보 {dbCandidates.Count}개 · 파서 필드 후보 {_parserFieldCandidates.Count}개. 양쪽에서 선택만 하면 자동 저장됩니다.",
            FontFamily = Font, FontSize = AppTheme.FontXS,
            Foreground = AppRes("FgDimmed"),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 6, 0, 0),
        });

        return section;
    }

    // ── 농도 계산 공식 빌더 ──────────────────────────────────────────────────
    private Control BuildFormulaSection(string analyte, List<string> availableCols)
    {
        if (_formulaAnalyte != analyte)
        {
            _formulaAnalyte = analyte;
            var raw = ETA.Services.SERVICE3.AnalysisNoteService.GetFormula(analyte);
            _formulaTokens = string.IsNullOrWhiteSpace(raw)
                ? new List<string>()
                : raw.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        var section = new StackPanel { Spacing = 4, Margin = new Thickness(0, 10, 0, 0) };
        section.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 0, 0, 4) });

        // 헤더 행: 제목 + 전체 적용 버튼
        var formulaHeaderGrid = new Grid { Margin = new Thickness(0, 0, 0, 2) };
        formulaHeaderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Star });
        formulaHeaderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var formulaHeaderLabel = new TextBlock
        {
            Text = "농도 계산 공식  (컬럼 드래그 또는 연산자/상수 추가)",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(formulaHeaderLabel, 0);
        var bulkStatusText = new TextBlock { FontFamily = Font, FontSize = AppTheme.FontXS, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center };
        var bulkBtnText = new TextBlock { FontFamily = Font, FontSize = AppTheme.FontXS, Foreground = new SolidColorBrush(Color.Parse("#88ccaa")), VerticalAlignment = VerticalAlignment.Center };
        void RefreshBulkLabel() {
            string? bov = ETA.Services.SERVICE3.AnalysisNoteService.GetSchemaOverride(analyte);
            string bAutoSch = WaterCenterDbMigration.DetermineSchema(_selectedAnalyte?.Analyte ?? analyte, _selectedAnalyte?.Category ?? "", _selectedAnalyte?.Method ?? "", _selectedAnalyte?.instrument ?? "");
            bulkBtnText.Text = $"▶ {bov ?? bAutoSch} 전체 적용";
        }
        RefreshBulkLabel();
        var bulkBtnInner = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        bulkBtnInner.Children.Add(bulkBtnText);
        bulkBtnInner.Children.Add(bulkStatusText);
        var bulkBtnBorder = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1a2e1a")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 2),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = bulkBtnInner
        };
        bulkBtnBorder.PointerReleased += (_, _) =>
        {
            if (_formulaTokens.Count == 0) { bulkStatusText.Text = "공식 없음"; return; }
            string? bov = ETA.Services.SERVICE3.AnalysisNoteService.GetSchemaOverride(analyte);
            string bAutoSch = WaterCenterDbMigration.DetermineSchema(_selectedAnalyte?.Analyte ?? analyte, _selectedAnalyte?.Category ?? "", _selectedAnalyte?.Method ?? "", _selectedAnalyte?.instrument ?? "");
            string bSch = bov ?? bAutoSch;
            int n = ETA.Services.SERVICE3.AnalysisNoteService.BulkSaveFormulaBySchema(bSch, string.Join(" ", _formulaTokens));
            bulkStatusText.Text = n > 0 ? $"✓ {n}개" : "0개 (스키마 불일치)";
        };
        Grid.SetColumn(bulkBtnBorder, 1);
        formulaHeaderGrid.Children.Add(formulaHeaderLabel);
        formulaHeaderGrid.Children.Add(bulkBtnBorder);
        section.Children.Add(formulaHeaderGrid);

        // 공식 표시 (드롭 존)
        var formulaDisplay = new WrapPanel { Orientation = Orientation.Horizontal, MinHeight = 34 };
        _formulaDisplayPanel = formulaDisplay;

        var dropZoneBorder = new Border
        {
            BorderBrush = new SolidColorBrush(Color.Parse("#555555")),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(6, 4),
            Margin = new Thickness(0, 0, 0, 6),
            Child = formulaDisplay
        };

        DragDrop.SetAllowDrop(dropZoneBorder, true);
        dropZoneBorder.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
        {
            if (e.Data.Contains("formula_col"))
            {
                e.DragEffects = DragDropEffects.Copy;
                dropZoneBorder.BorderBrush = new SolidColorBrush(Color.Parse("#4caf50"));
            }
            else e.DragEffects = DragDropEffects.None;
            e.Handled = true;
        });
        dropZoneBorder.AddHandler(DragDrop.DragLeaveEvent, (object? s, RoutedEventArgs e) =>
        {
            dropZoneBorder.BorderBrush = new SolidColorBrush(Color.Parse("#555555"));
        });
        dropZoneBorder.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
        {
            dropZoneBorder.BorderBrush = new SolidColorBrush(Color.Parse("#555555"));
            if (e.Data.Contains("formula_col") && e.Data.Get("formula_col") is string col)
            {
                _formulaTokens.Add(col);
                RefreshFormulaDisplay(formulaDisplay, analyte);
            }
            e.Handled = true;
        });

        RefreshFormulaDisplay(formulaDisplay, analyte);
        section.Children.Add(dropZoneBorder);

        // 연산자 버튼 행
        var opRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, Margin = new Thickness(0, 0, 0, 4) };
        opRow.Children.Add(new TextBlock { Text = "연산자:", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
        var ops = new[] { ("+", "+"), ("-", "-"), ("×", "*"), ("÷", "/"), ("(", "("), (")", ")") };
        foreach (var (label, token) in ops)
        {
            var tokenLocal = token;
            var opBtn = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#5d4037")),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 4),
                Margin = new Thickness(0, 0, 4, 0),
                Cursor = new Cursor(StandardCursorType.Hand),
                Child = new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#ffcc80")) }
            };
            opBtn.PointerPressed += (_, _) => { _formulaTokens.Add(tokenLocal); RefreshFormulaDisplay(formulaDisplay, analyte); };
            opRow.Children.Add(opBtn);
        }
        section.Children.Add(opRow);

        // 상수 입력 행
        var constRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        constRow.Children.Add(new TextBlock { Text = "상수:", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
        var constBox = new TextBox { Width = 90, Height = 28, FontFamily = Font, FontSize = AppTheme.FontSM, Watermark = "예: 1000", VerticalContentAlignment = VerticalAlignment.Center };
        var addConstBtn = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1a3a5c")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 4),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new TextBlock { Text = "추가", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#90caf9")) }
        };
        addConstBtn.PointerPressed += (_, _) =>
        {
            var val = constBox.Text?.Trim() ?? "";
            if (!string.IsNullOrEmpty(val)) { _formulaTokens.Add(val); constBox.Text = ""; RefreshFormulaDisplay(formulaDisplay, analyte); }
        };
        var clearBtn = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#4a0000")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 4),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new TextBlock { Text = "초기화", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#ef9a9a")) }
        };
        clearBtn.PointerPressed += (_, _) => { _formulaTokens.Clear(); RefreshFormulaDisplay(formulaDisplay, analyte); };
        constRow.Children.Add(constBox);
        constRow.Children.Add(addConstBtn);
        constRow.Children.Add(clearBtn);
        section.Children.Add(constRow);


        return section;
    }

    private void RefreshFormulaDisplay(WrapPanel panel, string analyte)
    {
        panel.Children.Clear();

        if (_formulaTokens.Count == 0)
        {
            panel.Children.Add(new TextBlock
            {
                Text = "컬럼을 드래그하거나 연산자/상수를 추가하세요  = 농도",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = new SolidColorBrush(Color.Parse("#666666")),
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 2)
            });
        }
        else
        {
            for (int i = 0; i < _formulaTokens.Count; i++)
            {
                var idx = i;
                var token = _formulaTokens[i];
                bool isOp  = "+-*/()".Contains(token) && token.Length == 1;
                bool isNum = double.TryParse(token, System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture, out _);
                var (bg, fg) = isOp  ? ("#5d4037", "#ffcc80") :
                               isNum ? ("#1a3a5c", "#90caf9") :
                                       ("#1b5e20", "#a5d6a7");
                var displayLabel = token switch { "*" => "×", "/" => "÷", _ => token };

                var inner = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 3 };
                inner.Children.Add(new TextBlock { Text = displayLabel, FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse(fg)), VerticalAlignment = VerticalAlignment.Center });
                var del = new TextBlock { Text = "\u00d7", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse(fg)), Margin = new Thickness(3, 0, 0, 0), Cursor = new Cursor(StandardCursorType.Hand), VerticalAlignment = VerticalAlignment.Center, IsVisible = false };
                del.PointerPressed += (s, e) =>
                {
                    e.Handled = true; // \ubd80\ubaa8\uc758 \ub4dc\ub798\uadf8 \uc2dc\uc791 \ub9c9\uae30
                    if (idx < _formulaTokens.Count) _formulaTokens.RemoveAt(idx);
                    RefreshFormulaDisplay(panel, analyte);
                };
                inner.Children.Add(del);
                var tokenBorder = new Border
                {
                    Background = new SolidColorBrush(Color.Parse(bg)),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(6, 3),
                    Margin = new Thickness(0, 2, 4, 2),
                    BorderBrush = Brushes.Transparent,
                    BorderThickness = new Thickness(2),
                    Cursor = new Cursor(StandardCursorType.SizeAll),
                    Child = inner
                };
                tokenBorder.PointerEntered += (_, _) => del.IsVisible = true;
                tokenBorder.PointerExited  += (_, _) => del.IsVisible = false;

                // \u2500\u2500 \ub4dc\ub798\uadf8 \uc18c\uc2a4: \uc790\uae30 \uc778\ub371\uc2a4\ub97c \ub370\uc774\ud130\ub85c \ub4e4\uace0 \uc774\ub3d9 \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
                tokenBorder.PointerPressed += async (s, e) =>
                {
                    if (!e.GetCurrentPoint(tokenBorder).Properties.IsLeftButtonPressed) return;
                    var data = new DataObject();
                    data.Set("formula_token_idx", idx.ToString());
                    await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                };

                // \u2500\u2500 \ub4dc\ub86d \ud0c0\uac9f: \ub2e4\ub978 \ud1a0\ud070 \ub4dc\ub86d \uc2dc \uadf8 \uc704\uce58 \uc55e\uc73c\ub85c \uc0bd\uc785 \u2500\u2500\u2500\u2500\u2500
                DragDrop.SetAllowDrop(tokenBorder, true);
                tokenBorder.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
                {
                    if (e.Data.Contains("formula_token_idx"))
                    {
                        e.DragEffects = DragDropEffects.Move;
                        tokenBorder.BorderBrush = new SolidColorBrush(Color.Parse("#4caf50"));
                    }
                    else e.DragEffects = DragDropEffects.None;
                    e.Handled = true;
                });
                tokenBorder.AddHandler(DragDrop.DragLeaveEvent, (object? s, RoutedEventArgs e) =>
                {
                    tokenBorder.BorderBrush = Brushes.Transparent;
                });
                tokenBorder.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
                {
                    tokenBorder.BorderBrush = Brushes.Transparent;
                    if (e.Data.Contains("formula_token_idx")
                        && e.Data.Get("formula_token_idx") is string srcStr
                        && int.TryParse(srcStr, out var src)
                        && src >= 0 && src < _formulaTokens.Count
                        && src != idx)
                    {
                        var t = _formulaTokens[src];
                        _formulaTokens.RemoveAt(src);
                        int insertAt = src < idx ? idx - 1 : idx;
                        if (insertAt < 0) insertAt = 0;
                        if (insertAt > _formulaTokens.Count) insertAt = _formulaTokens.Count;
                        _formulaTokens.Insert(insertAt, t);
                        RefreshFormulaDisplay(panel, analyte);
                    }
                    e.Handled = true;
                });
                panel.Children.Add(tokenBorder);
            }
            // 유효성 검사 아이콘 — 괄호 균형 + 연산자 연속 여부
            bool valid = IsFormulaValid(_formulaTokens);
            var validIcon = new TextBlock
            {
                Text = valid ? "✓" : "✗",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                FontWeight = FontWeight.Bold,
                Foreground = valid
                    ? new SolidColorBrush(Color.Parse("#66cc88"))
                    : new SolidColorBrush(Color.Parse("#ff6666")),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 4, 0),
            };
            panel.Children.Add(new TextBlock { Text = "= 농도", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 2) });
            panel.Children.Add(validIcon);
        }

        ETA.Services.SERVICE3.AnalysisNoteService.SaveFormula(analyte, string.Join(" ", _formulaTokens));
    }

    // =========================================================================
    // 탭 3 — 분석장비 관리 Show1
    // =========================================================================
    private static bool IsFormulaValid(List<string> tokens)
    {
        if (tokens.Count == 0) return false;
        // 1. 괄호 균형
        int depth = 0;
        foreach (var t in tokens)
        {
            if (t == "(") depth++;
            else if (t == ")") { depth--; if (depth < 0) return false; }
        }
        if (depth != 0) return false;
        // 2. 연산자로 시작/끝나지 않을 것
        bool IsOp(string t) => t == "+" || t == "-" || t == "*" || t == "/";
        if (IsOp(tokens[0]) || IsOp(tokens[tokens.Count - 1])) return false;
        // 3. 연산자 연속 없을 것
        for (int i = 0; i < tokens.Count - 1; i++)
            if (IsOp(tokens[i]) && IsOp(tokens[i + 1])) return false;
        return true;
    }

    private Control BuildEquipmentPanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var scrapeStatusTb = new TextBlock { FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(8, 2, 8, 4), IsVisible = false };

        var hdrGrid = new Grid();
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        var hdrTitle = new TextBlock { Text = "분석장비 목록", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"), FontFamily = Font, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 6) };
        Grid.SetColumn(hdrTitle, 0);
        var scrapeBtn = new Button { Content = "📡 스크랩", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        scrapeBtn.Click += async (_, _) =>
        {
            scrapeBtn.IsEnabled = false;
            scrapeStatusTb.IsVisible = true;
            scrapeStatusTb.Text = "스크랩 중...";
            var (saved, err) = await MeasurerCdpService.ScrapeEquipmentsAsync(msg =>
                Dispatcher.UIThread.Post(() => { scrapeStatusTb.Text = msg; }));
            scrapeBtn.IsEnabled = true;
            if (err != null) { scrapeStatusTb.Text = "⚠ " + err; scrapeStatusTb.Foreground = Brushes.OrangeRed; }
            else { scrapeStatusTb.Text = "✓ " + saved + "개 저장됨"; scrapeStatusTb.Foreground = Brushes.LightGreen; LoadEquipmentList(); }
        };
        Grid.SetColumn(scrapeBtn, 1);
        var scrapeFieldBtn = new Button { Content = "🎯 현장장비 스크랩", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusAccentBg, Foreground = AppTheme.StatusAccentFg, BorderBrush = AppTheme.StatusAccentBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        Avalonia.Controls.ToolTip.SetTip(scrapeFieldBtn, "의뢰추가 모달의 edit_meas_equip_no 드롭다운에서 직접 스크랩 — 모달이 열려있어야 함");
        scrapeFieldBtn.Click += async (_, _) =>
        {
            scrapeFieldBtn.IsEnabled = false;
            scrapeStatusTb.IsVisible = true;
            scrapeStatusTb.Text = "현장장비 스크랩 중...";
            var (saved, err) = await MeasurerCdpService.ScrapeFieldMeasEquipmentsAsync(msg =>
                Dispatcher.UIThread.Post(() => { scrapeStatusTb.Text = msg; }));
            scrapeFieldBtn.IsEnabled = true;
            if (err != null) { scrapeStatusTb.Text = "⚠ " + err; scrapeStatusTb.Foreground = Brushes.OrangeRed; }
            else { scrapeStatusTb.Text = $"✓ 현장측정장비 {saved}개 저장됨"; scrapeStatusTb.Foreground = Brushes.LightGreen; }
        };
        Grid.SetColumn(scrapeFieldBtn, 2);
        var addBtn = new Button { Content = "+ 추가", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusOkBg, Foreground = AppTheme.StatusOkFg, BorderBrush = AppTheme.StatusOkBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        addBtn.Click += (_, _) => { _show2Content.Content = BuildEquipmentEditPanel(null, null, isNew: true); };
        Grid.SetColumn(addBtn, 3);
        hdrGrid.Children.Add(hdrTitle); hdrGrid.Children.Add(scrapeBtn); hdrGrid.Children.Add(scrapeFieldBtn); hdrGrid.Children.Add(addBtn);
        var hdrOuter = new StackPanel { Spacing = 0 };
        hdrOuter.Children.Add(hdrGrid);
        hdrOuter.Children.Add(scrapeStatusTb);
        var headerBorder = new Border { Background = AppRes("PanelBg"), Child = hdrOuter };
        Grid.SetRow(headerBorder, 0);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = _equipmentListPanel };
        Grid.SetRow(scroll, 1);
        grid.Children.Add(headerBorder); grid.Children.Add(scroll);
        return grid;
    }

    private void LoadEquipmentList()
    {
        _equipmentListPanel.Children.Clear();
        try
        {
            var list = MeasurerService.GetAllEquipments();
            if (list.Count == 0)
            {
                _equipmentListPanel.Children.Add(new TextBlock
                {
                    Text         = "등록된 장비가 없습니다.\n+ 장비 추가 버튼으로 직접 입력하세요.",
                    FontSize     = AppTheme.FontBase,
                    FontFamily   = Font,
                    Foreground   = AppRes("FgMuted"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin       = new Thickness(12, 16),
                });
                return;
            }
            foreach (var (name, code) in list)
                _equipmentListPanel.Children.Add(CreateEquipmentRow(name, code));
        }
        catch (Exception ex) { Console.WriteLine($"[LoadEquipmentList] {ex.Message}"); }
    }

    private Border CreateEquipmentRow(string name, string code)
    {
        try
        {
            System.IO.Directory.CreateDirectory("Logs");
            if (ETA.App.EnableLogging) System.IO.File.AppendAllText("Logs/MeasurerDebug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CreateEquipmentRow] name='{name}' code='{code}'\n");
        }
        catch { }
        var row = new Grid { Margin = new Thickness(4, 2) };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));  // 현장측정 체크박스
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));  // 이름/코드
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        // 현장측정 플래그 체크박스 (행 내 인라인 — 즉시 저장)
        bool cur = false;
        try { cur = MeasurerService.GetFieldMeasFlag(code); } catch { }
        var cbField = new CheckBox
        {
            IsChecked = cur,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 10, 0),
        };
        Avalonia.Controls.ToolTip.SetTip(cbField, "현장측정장비");
        // Click(부모 PointerPressed 우선권 회피) + IsCheckedChanged 둘 다 연결
        void CommitField()
        {
            try
            {
                bool on = cbField.IsChecked == true;
                MeasurerService.SetFieldMeasFlag(code, on);
                try
                {
                    System.IO.Directory.CreateDirectory("Logs");
                    if (ETA.App.EnableLogging) System.IO.File.AppendAllText("Logs/MeasurerDebug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UI:인라인체크] code='{code}' on={on}\n");
                }
                catch { }
            }
            catch (Exception ex) { Console.WriteLine($"[SetFieldMeasFlag] {ex.Message}"); }
        }
        cbField.IsCheckedChanged += (_, _) => CommitField();
        cbField.Click           += (_, _) => CommitField();
        Grid.SetColumn(cbField, 0);
        row.Children.Add(cbField);

        var nameBlock = new TextBlock { Text = name, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center };
        var codeBlock = new TextBlock { Text = code, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 2) };
        var info = new StackPanel { Spacing = 1 };
        info.Children.Add(nameBlock); info.Children.Add(codeBlock);
        Grid.SetColumn(info, 1);
        row.Children.Add(info);

        var border = new Border
        {
            Padding      = new Thickness(8, 6),
            CornerRadius = new CornerRadius(4),
            Background   = Brushes.Transparent,
            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child        = row,
        };
        // 진단: Border 전체 클릭 탐지 (이벤트가 어디까지 오는지 확인)
        border.PointerPressed += (s, e) =>
        {
            try
            {
                System.IO.Directory.CreateDirectory("Logs");
                var src = e.Source?.GetType().Name ?? "null";
                if (ETA.App.EnableLogging) System.IO.File.AppendAllText("Logs/MeasurerDebug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Border.PointerPressed] code='{code}' src={src}\n");
            }
            catch { }
        };
        // 행 클릭 → 편집 패널 (체크박스 클릭은 CheckBox가 Handled 처리)
        info.PointerPressed += (_, _) =>
        {
            _show2Content.Content = BuildEquipmentEditPanel(name, code, isNew: false);
        };
        border.PointerEntered += (_, _) => border.Background = AppTheme.BgSecondary;
        border.PointerExited  += (_, _) => border.Background = Brushes.Transparent;
        return border;
    }

    // =========================================================================
    // 탭 3 — 분석장비 관리 Show2
    // =========================================================================
    private Control BuildEquipmentEditPanel(string? name, string? code, bool isNew)
    {
        var origCode = code;

        TextBox MkTb(string val, string hint) => new TextBox
        {
            Text = val, Watermark = hint, FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(6, 4), Margin = new Thickness(0, 2),
            Background = AppRes("PanelBg"), BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
        };

        var tbName = MkTb(name ?? "", "장비명 (예: ICP-OES / Plasma Quant 9100)");
        var tbCode = MkTb(code ?? "", "코드값 (측정인.kr 드롭다운 ID)");

        Control Labeled(string lbl, TextBox tb)
        {
            var s = new StackPanel { Spacing = 0 };
            s.Children.Add(new TextBlock { Text = lbl, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 4, 0, 1) });
            s.Children.Add(tb);
            return s;
        }

        var form = new StackPanel { Spacing = 2, Margin = new Thickness(0, 6) };
        form.Children.Add(Labeled("장비명", tbName));
        form.Children.Add(Labeled("코드값 (측정인.kr ID)", tbCode));

        // 현장측정장비 플래그 — 측정인 전송 시 이 장비들이 자동 선택됨
        bool curField = false;
        if (!isNew && !string.IsNullOrWhiteSpace(origCode))
        {
            try { curField = MeasurerService.GetFieldMeasFlag(origCode!); } catch { }
        }
        var cbField = new CheckBox
        {
            Content    = "현장측정장비 (측정인 전송 시 자동 선택)",
            IsChecked  = curField,
            FontSize   = AppTheme.FontBase,
            FontFamily = Font,
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 8, 0, 0),
        };
        form.Children.Add(cbField);

        var btnSave = new Button { Content = isNew ? "✚ 저장" : "💾 저장", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        var statusTb = new TextBlock { FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), IsVisible = false, Margin = new Thickness(0, 4) };
        btnSave.Click += (_, _) =>
        {
            var n = tbName.Text?.Trim() ?? "";
            var c = tbCode.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(n)) { statusTb.Text = "장비명을 입력하세요."; statusTb.IsVisible = true; return; }
            if (string.IsNullOrWhiteSpace(c)) c = n; // 코드값 없으면 장비명으로 대체
            try
            {
                MeasurerService.SaveEquipment(n, c, isNew ? null : origCode);
                try { MeasurerService.SetFieldMeasFlag(c, cbField.IsChecked == true); }
                catch (Exception ex2) { Console.WriteLine($"[SetFieldMeasFlag] {ex2.Message}"); }
                LoadEquipmentList();
                _show2Content.Content = BuildEquipmentEditPanel(n, c, isNew: false);
            }
            catch (Exception ex) { statusTb.Text = $"오류: {ex.Message}"; statusTb.Foreground = Brushes.Red; statusTb.IsVisible = true; }
        };

        var btnDelete = new Button { Content = "🗑 삭제", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = AppTheme.StatusBadBg, Foreground = AppTheme.StatusBadFg, BorderBrush = AppTheme.StatusBadBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold, IsVisible = !isNew };
        btnDelete.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(origCode)) return;
            try { MeasurerService.DeleteEquipment(origCode); LoadEquipmentList(); _show2Content.Content = BuildEquipmentHint(); }
            catch (Exception ex) { statusTb.Text = $"오류: {ex.Message}"; statusTb.Foreground = Brushes.Red; statusTb.IsVisible = true; }
        };

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0, 10, 0, 4) };
        btnRow.Children.Add(btnSave);
        if (!isNew) btnRow.Children.Add(btnDelete);

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock { Text = isNew ? "✚ 새 분석장비 추가" : $"분석장비 편집 — {name}", FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 4) });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 6) });
        root.Children.Add(new TextBlock { Text = "장비명과 코드값을 입력하세요.\n코드값은 측정인.kr 드롭다운의 element ID입니다.\n모를 경우 장비명만 입력해도 됩니다.", FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
        root.Children.Add(form);
        root.Children.Add(btnRow);
        root.Children.Add(statusTb);
        return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = root };
    }

    // =========================================================================
    // 탭 2 — 처리시설 매칭 Show1
    // =========================================================================
    private Control BuildFacilityMatchPanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var hdrTitle = new TextBlock
        {
            Text              = "처리시설 마스터 (시설·시료)",
            FontSize          = AppTheme.FontBase,
            FontWeight        = FontWeight.SemiBold,
            Foreground        = AppRes("FgMuted"),
            FontFamily        = Font,
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(10, 6),
        };
        var headerBorder = new Border { Background = AppRes("PanelBg"), Child = hdrTitle };
        Grid.SetRow(headerBorder, 0);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = _facilityMatchListPanel };
        Grid.SetRow(scroll, 1);
        grid.Children.Add(headerBorder);
        grid.Children.Add(scroll);
        return grid;
    }

    private void LoadFacilityMatchList()
    {
        _facilityMatchListPanel.Children.Clear();
        try
        {
            var pairs = FacilityMappingService.GetAllMasterPairs();
            if (pairs.Count == 0)
            {
                _facilityMatchListPanel.Children.Add(new TextBlock
                {
                    Text         = "처리시설_마스터 에 등록된 (시설·시료) 조합이 없습니다.",
                    FontSize     = AppTheme.FontBase,
                    FontFamily   = Font,
                    Foreground   = AppRes("FgMuted"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin       = new Thickness(12, 16),
                });
                return;
            }

            string? curFac = null;
            foreach (var (fac, smp) in pairs)
            {
                if (fac != curFac)
                {
                    curFac = fac;
                    _facilityMatchListPanel.Children.Add(new TextBlock
                    {
                        Text       = string.IsNullOrWhiteSpace(fac) ? "(시설명 없음)" : fac,
                        FontSize   = AppTheme.FontSM,
                        FontWeight = FontWeight.SemiBold,
                        FontFamily = Font,
                        Foreground = AppRes("FgMuted"),
                        Margin     = new Thickness(10, 8, 10, 2),
                    });
                }
                _facilityMatchListPanel.Children.Add(CreateFacilityMatchRow(fac, smp));
            }
        }
        catch (Exception ex) { Console.WriteLine($"[LoadFacilityMatchList] {ex.Message}"); }
    }

    private Border CreateFacilityMatchRow(string 시설명, string 시료명)
    {
        int mapCount = 0;
        try { mapCount = FacilityMappingService.GetAllForMaster(시설명, 시료명).Count; } catch { }

        var row = new Grid { Margin = new Thickness(16, 2, 6, 2) };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var nameBlock = new TextBlock
        {
            Text              = string.IsNullOrWhiteSpace(시료명) ? "(시료명 없음)" : 시료명,
            FontSize          = AppTheme.FontBase,
            FontFamily        = Font,
            Foreground        = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(nameBlock, 0);
        row.Children.Add(nameBlock);

        var badge = new Border
        {
            Background      = mapCount > 0 ? AppTheme.StatusOkBg : AppTheme.StatusInfoBg,
            BorderBrush     = mapCount > 0 ? AppTheme.StatusOkBorder : AppTheme.StatusInfoBorder,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(999),
            Padding         = new Thickness(8, 1),
            Margin          = new Thickness(4, 0),
            Child = new TextBlock
            {
                Text       = mapCount > 0 ? $"{mapCount}개 매핑" : "미매핑",
                FontSize   = AppTheme.FontXS,
                FontFamily = Font,
                Foreground = mapCount > 0 ? AppTheme.StatusOkFg : AppTheme.StatusInfoFg,
                FontWeight = FontWeight.SemiBold,
            },
        };
        Grid.SetColumn(badge, 1);
        row.Children.Add(badge);

        var border = new Border
        {
            Padding      = new Thickness(6, 4),
            CornerRadius = new CornerRadius(4),
            Background   = Brushes.Transparent,
            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child        = row,
        };
        border.PointerPressed += (_, _) =>
        {
            _selectedFacilityPair = (시설명, 시료명);
            _show2Content.Content = BuildFacilityMatchEditPanel(시설명, 시료명);
        };
        border.PointerEntered += (_, _) => border.Background = AppTheme.BgSecondary;
        border.PointerExited  += (_, _) => border.Background = Brushes.Transparent;
        return border;
    }

    // =========================================================================
    // 탭 2 — 처리시설 매칭 Show2
    // =========================================================================
    private Control BuildFacilityMatchEditPanel(string 시설명, string 시료명)
    {
        TextBox MkTb(string val, string hint) => new TextBox
        {
            Text = val, Watermark = hint, FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(6, 4), Margin = new Thickness(0, 2),
            Background = AppRes("PanelBg"), BorderBrush = AppRes("InputBorder"),
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
        };

        Control Labeled(string lbl, Control c)
        {
            var s = new StackPanel { Spacing = 0 };
            s.Children.Add(new TextBlock { Text = lbl, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 4, 0, 1) });
            s.Children.Add(c);
            return s;
        }

        var cbSystem = new ComboBox
        {
            FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(6, 2), Margin = new Thickness(0, 2),
            Background = AppRes("PanelBg"), BorderBrush = AppRes("InputBorder"),
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
        };
        cbSystem.Items.Add("WAYBLE");
        cbSystem.Items.Add("측정인");
        cbSystem.SelectedIndex = 0;

        var tbSiteCd   = MkTb("", "외부 시설 코드 (WAYBLE siteCd / 측정인 처리시설코드)");
        var tbSiteNm   = MkTb("", "외부 시설명");
        var tbSampleCd = MkTb("", "외부 시료 코드 (WAYBLE sampleCategory)");
        var tbSampleNm = MkTb("", "외부 시료명");

        // WAYBLE 카탈로그 드롭다운 — 선택 시 아래 4개 필드 자동 채움
        var cbCatalog = new ComboBox
        {
            FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(6, 2), Margin = new Thickness(0, 2),
            Background = AppRes("PanelBg"), BorderBrush = AppRes("InputBorder"),
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
            PlaceholderText = "WAYBLE 카탈로그에서 선택...",
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        var catalogEntries = new List<ExternalSiteCatalogService.Entry>();
        void ReloadCatalog(int? preselectIndex = null)
        {
            cbCatalog.Items.Clear();
            catalogEntries.Clear();
            var all = ExternalSiteCatalogService.GetWayble();
            // 시료명 매칭 우선 정렬 (sampleCategoryNm 에 시료명 포함 / 시료명 에 sampleCategoryNm 포함)
            string Norm(string s) => (s ?? "").Replace(" ", "").Replace("_", "").ToLower();
            var targetSample = Norm(시료명);
            var targetFac    = Norm(시설명);
            int ScoreOf(ExternalSiteCatalogService.Entry e)
            {
                int score = 0;
                var ns = Norm(e.SampleCategoryNm);
                var nf = Norm(e.SiteNm);
                if (!string.IsNullOrEmpty(targetSample))
                {
                    if (ns == targetSample) score += 100;
                    else if (ns.Contains(targetSample) || targetSample.Contains(ns)) score += 50;
                }
                if (!string.IsNullOrEmpty(targetFac))
                {
                    if (nf.Contains(targetFac) || targetFac.Contains(nf)) score += 30;
                }
                return score;
            }
            var sorted = new List<ExternalSiteCatalogService.Entry>(all);
            sorted.Sort((a, b) => ScoreOf(b).CompareTo(ScoreOf(a)));
            foreach (var e in sorted)
            {
                catalogEntries.Add(e);
                var score = ScoreOf(e);
                var prefix = score >= 100 ? "⭐ " : score >= 50 ? "✨ " : score >= 30 ? "· " : "  ";
                cbCatalog.Items.Add($"{prefix}{e.OfficeCd} / {e.SiteCd} {e.SiteNm} — {e.SampleCategory} {e.SampleCategoryNm}");
            }
            if (preselectIndex is int pi && pi >= 0 && pi < cbCatalog.Items.Count)
                cbCatalog.SelectedIndex = pi;
        }
        ReloadCatalog();

        cbCatalog.SelectionChanged += (_, _) =>
        {
            var idx = cbCatalog.SelectedIndex;
            if (idx < 0 || idx >= catalogEntries.Count) return;
            var e = catalogEntries[idx];
            cbSystem.SelectedIndex = 0; // WAYBLE
            tbSiteCd.Text   = e.SiteCd;
            tbSiteNm.Text   = e.SiteNm;
            tbSampleCd.Text = e.SampleCategory;
            tbSampleNm.Text = e.SampleCategoryNm;
        };

        var btnRefreshCat = new Button
        {
            Content = "🔄", FontSize = AppTheme.FontSM, FontFamily = Font,
            Padding = new Thickness(10, 2), Margin = new Thickness(4, 2, 0, 2),
            Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg,
            BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(999),
        };
        Avalonia.Controls.ToolTip.SetTip(btnRefreshCat, "xlsx에서 카탈로그 재추출 (Python 스크립트)");

        var catalogRow = new Grid();
        catalogRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        catalogRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        Grid.SetColumn(cbCatalog, 0);
        Grid.SetColumn(btnRefreshCat, 1);
        catalogRow.Children.Add(cbCatalog);
        catalogRow.Children.Add(btnRefreshCat);

        var form = new StackPanel { Spacing = 2, Margin = new Thickness(0, 6) };
        form.Children.Add(Labeled("📂 WAYBLE 카탈로그 (xlsx)", catalogRow));
        form.Children.Add(Labeled("외부 시스템", cbSystem));
        form.Children.Add(Labeled("외부 시설 코드", tbSiteCd));
        form.Children.Add(Labeled("외부 시설명", tbSiteNm));
        form.Children.Add(Labeled("외부 시료 코드", tbSampleCd));
        form.Children.Add(Labeled("외부 시료명", tbSampleNm));

        var statusTb = new TextBlock { FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), IsVisible = false, Margin = new Thickness(0, 4) };

        btnRefreshCat.Click += (_, _) =>
        {
            btnRefreshCat.IsEnabled = false;
            statusTb.Text = "xlsx 재스캔 중...";
            statusTb.Foreground = AppRes("FgMuted");
            statusTb.IsVisible = true;
            var (count, err) = ExternalSiteCatalogService.RefreshFromXlsx();
            btnRefreshCat.IsEnabled = true;
            if (err != null)
            {
                statusTb.Text = $"⚠ 재스캔 실패: {err}";
                statusTb.Foreground = Brushes.OrangeRed;
            }
            else
            {
                ReloadCatalog();
                statusTb.Text = $"✓ {count}개 조합 갱신됨 ({ExternalSiteCatalogService.LastSource})";
                statusTb.Foreground = Brushes.LightGreen;
            }
        };

        var existingPanel = new StackPanel { Spacing = 1 };
        void ReloadExisting()
        {
            existingPanel.Children.Clear();
            List<FacilityMappingService.Row> rows;
            try { rows = FacilityMappingService.GetAllForMaster(시설명, 시료명); }
            catch (Exception ex) { Console.WriteLine($"[FacilityMapping.Reload] {ex.Message}"); return; }

            if (rows.Count == 0)
            {
                existingPanel.Children.Add(new TextBlock
                {
                    Text = "등록된 매핑이 없습니다.", FontSize = AppTheme.FontSM, FontFamily = Font,
                    Foreground = AppRes("FgMuted"), Margin = new Thickness(6, 4),
                });
                return;
            }
            foreach (var r in rows)
            {
                var rowGrid = new Grid { Margin = new Thickness(4, 2) };
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

                var sysBadge = new Border
                {
                    Background      = r.ExternalSystem == "WAYBLE" ? AppTheme.StatusInfoBg : AppTheme.StatusAccentBg,
                    BorderBrush     = r.ExternalSystem == "WAYBLE" ? AppTheme.StatusInfoBorder : AppTheme.StatusAccentBorder,
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(999),
                    Padding         = new Thickness(8, 1),
                    Margin          = new Thickness(0, 0, 6, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text       = r.ExternalSystem,
                        FontSize   = AppTheme.FontXS,
                        FontFamily = Font,
                        Foreground = r.ExternalSystem == "WAYBLE" ? AppTheme.StatusInfoFg : AppTheme.StatusAccentFg,
                        FontWeight = FontWeight.SemiBold,
                    },
                };
                Grid.SetColumn(sysBadge, 0);
                rowGrid.Children.Add(sysBadge);

                var info = new StackPanel { Spacing = 1, VerticalAlignment = VerticalAlignment.Center };
                info.Children.Add(new TextBlock
                {
                    Text       = $"{r.ExternalSiteCd} · {r.ExternalSiteNm}",
                    FontSize   = AppTheme.FontSM,
                    FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                });
                info.Children.Add(new TextBlock
                {
                    Text       = string.IsNullOrWhiteSpace(r.ExternalSampleCd)
                                     ? $"(시료코드 없음) {r.ExternalSampleNm}"
                                     : $"{r.ExternalSampleCd} · {r.ExternalSampleNm}",
                    FontSize   = AppTheme.FontXS,
                    FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                });
                Grid.SetColumn(info, 1);
                rowGrid.Children.Add(info);

                var btnRowPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
                var btnEdit = new Button
                {
                    Content = "✎", FontSize = AppTheme.FontSM, FontFamily = Font,
                    Padding = new Thickness(8, 2), Background = AppTheme.StatusInfoBg,
                    Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder,
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
                };
                btnEdit.Click += (_, _) =>
                {
                    cbSystem.SelectedIndex = r.ExternalSystem == "측정인" ? 1 : 0;
                    tbSiteCd.Text   = r.ExternalSiteCd;
                    tbSiteNm.Text   = r.ExternalSiteNm;
                    tbSampleCd.Text = r.ExternalSampleCd;
                    tbSampleNm.Text = r.ExternalSampleNm;
                    statusTb.Text = "↑ 폼에 로드됨 — 수정 후 저장하면 UPSERT";
                    statusTb.Foreground = AppRes("FgMuted");
                    statusTb.IsVisible = true;
                };
                var btnDel = new Button
                {
                    Content = "🗑", FontSize = AppTheme.FontSM, FontFamily = Font,
                    Padding = new Thickness(8, 2), Background = AppTheme.StatusBadBg,
                    Foreground = AppTheme.StatusBadFg, BorderBrush = AppTheme.StatusBadBorder,
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
                };
                btnDel.Click += (_, _) =>
                {
                    try { FacilityMappingService.Delete(r.Id); ReloadExisting(); LoadFacilityMatchList(); _show4Content.Content = BuildWaybleCatalogPanel(); }
                    catch (Exception ex) { statusTb.Text = $"오류: {ex.Message}"; statusTb.Foreground = Brushes.Red; statusTb.IsVisible = true; }
                };
                btnRowPanel.Children.Add(btnEdit);
                btnRowPanel.Children.Add(btnDel);
                Grid.SetColumn(btnRowPanel, 2);
                rowGrid.Children.Add(btnRowPanel);

                var border = new Border
                {
                    Padding         = new Thickness(6, 4),
                    CornerRadius    = new CornerRadius(4),
                    Background      = AppRes("PanelBg"),
                    BorderBrush     = AppTheme.BorderDefault,
                    BorderThickness = new Thickness(1),
                    Margin          = new Thickness(0, 2),
                    Child           = rowGrid,
                };
                existingPanel.Children.Add(border);
            }
        }
        ReloadExisting();

        var btnSave = new Button
        {
            Content = "💾 저장 (UPSERT)", FontSize = AppTheme.FontBase, FontFamily = Font,
            Padding = new Thickness(14, 6), Background = AppTheme.StatusInfoBg,
            Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder,
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
            FontWeight = FontWeight.SemiBold,
        };
        btnSave.Click += (_, _) =>
        {
            var sys    = cbSystem.SelectedItem?.ToString() ?? "WAYBLE";
            var sCd    = tbSiteCd.Text?.Trim() ?? "";
            var sNm    = tbSiteNm.Text?.Trim() ?? "";
            var pCd    = tbSampleCd.Text?.Trim() ?? "";
            var pNm    = tbSampleNm.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(sCd))
            {
                statusTb.Text = "외부 시설 코드를 입력하세요.";
                statusTb.Foreground = Brushes.Red;
                statusTb.IsVisible = true;
                return;
            }
            try
            {
                FacilityMappingService.Save(sys, sCd, sNm, pCd, pNm, 시설명, 시료명);
                statusTb.Text = $"✓ 저장됨 — {sys} / {sCd} / {pCd}";
                statusTb.Foreground = Brushes.LightGreen;
                statusTb.IsVisible = true;
                tbSiteCd.Text = ""; tbSiteNm.Text = "";
                tbSampleCd.Text = ""; tbSampleNm.Text = "";
                ReloadExisting();
                LoadFacilityMatchList();
                _show4Content.Content = BuildWaybleCatalogPanel();
            }
            catch (Exception ex)
            {
                statusTb.Text = $"오류: {ex.Message}";
                statusTb.Foreground = Brushes.Red;
                statusTb.IsVisible = true;
            }
        };

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock
        {
            Text       = $"🏭 {시설명} — {시료명}",
            FontSize   = AppTheme.FontXL,
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font,
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 4),
        });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 6) });
        root.Children.Add(new TextBlock
        {
            Text         = "ETA 처리시설_마스터 의 (시설명·시료명) 과 외부 시스템의 (사이트코드·시료코드) 를 연결합니다.\n" +
                           "WAYBLE 은 siteCd / sampleCategory, 측정인.kr 은 처리시설코드 를 사용하세요.\n" +
                           "중복 저장 시 (external_system + external_site_cd + external_sample_cd) 기준 UPSERT 됩니다.",
            FontSize     = AppTheme.FontSM,
            FontFamily   = Font,
            Foreground   = AppRes("FgMuted"),
            TextWrapping = TextWrapping.Wrap,
            Margin       = new Thickness(0, 0, 0, 8),
        });
        root.Children.Add(new TextBlock
        {
            Text       = "➕ 새 매핑 입력",
            FontSize   = AppTheme.FontBase,
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font,
            Foreground = AppRes("FgMuted"),
            Margin     = new Thickness(0, 4, 0, 2),
        });
        root.Children.Add(form);
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0, 10, 0, 4) };
        btnRow.Children.Add(btnSave);
        root.Children.Add(btnRow);
        root.Children.Add(statusTb);

        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 12, 0, 6) });
        root.Children.Add(new TextBlock
        {
            Text       = "📋 기존 매핑",
            FontSize   = AppTheme.FontBase,
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font,
            Foreground = AppRes("FgMuted"),
            Margin     = new Thickness(0, 0, 0, 4),
        });
        root.Children.Add(existingPanel);

        return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = root };
    }

    // =========================================================================
    // 탭 2 — 처리시설 매칭 Show4 (WAYBLE 카탈로그 뷰)
    // =========================================================================
    private Control BuildWaybleCatalogPanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var catalog = ExternalSiteCatalogService.GetWayble();
        var allMaps = FacilityMappingService.GetAllMappings();
        // 인덱스: (external_site_cd, external_sample_cd) → ETA (시설명, 시료명)
        var mapIndex = new Dictionary<(string, string), (string 시설명, string 시료명)>();
        foreach (var m in allMaps)
        {
            if (m.ExternalSystem != "WAYBLE") continue;
            mapIndex[(m.ExternalSiteCd, m.ExternalSampleCd)] = (m.EtaSiseol, m.EtaSiryo);
        }
        int matched = catalog.Count(e => mapIndex.ContainsKey((e.SiteCd, e.SampleCategory)));

        var hdrTb = new TextBlock
        {
            Text              = $"📂 WAYBLE 카탈로그  {matched} / {catalog.Count} 매칭됨",
            FontSize          = AppTheme.FontBase,
            FontWeight        = FontWeight.SemiBold,
            FontFamily        = Font,
            Foreground        = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(10, 6),
        };
        var openXlsxBtn = new Button
        {
            Content = "📊 xlsx 열기", FontSize = AppTheme.FontSM, FontFamily = Font,
            Padding = new Thickness(8, 4), Margin = new Thickness(4),
            Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg,
            BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold,
        };
        Avalonia.Controls.ToolTip.SetTip(openXlsxBtn, "최신 WAYBLE xlsx 를 DataGrid 뷰어로 미리보기");
        openXlsxBtn.Click += (_, _) =>
        {
            try
            {
                var dir = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Data", "Exports");
                if (!System.IO.Directory.Exists(dir)) return;
                var latest = System.IO.Directory.GetFiles(dir, "wayble_수질현황_*.xlsx")
                    .Where(f => !System.IO.Path.GetFileName(f).StartsWith("~$"))
                    .OrderByDescending(System.IO.File.GetLastWriteTime)
                    .FirstOrDefault();
                if (latest == null)
                {
                    Console.WriteLine("[WAYBLE xlsx] 파일 없음");
                    return;
                }
                new ETA.Views.XlsxPreviewWindow(latest).Show();
            }
            catch (Exception ex) { Console.WriteLine($"[openXlsx] {ex.Message}"); }
        };

        var waybleLoginBtn = new Button
        {
            Content = ETA.Services.SERVICE2.WaybleSession.Connected ? "🔑 Wayble 연결됨" : "🔑 Wayble 로그인",
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Padding = new Thickness(8, 4), Margin = new Thickness(4),
            Background = AppTheme.StatusAccentBg, Foreground = AppTheme.StatusAccentFg,
            BorderBrush = AppTheme.StatusAccentBorder, BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold,
        };
        Avalonia.Controls.ToolTip.SetTip(waybleLoginBtn, "rewater.wayble.eco 세션 수립 (CDP Edge 브라우저)");
        waybleLoginBtn.Click += async (_, _) =>
        {
            var parent = Avalonia.VisualTree.VisualExtensions
                .FindAncestorOfType<Window>(waybleLoginBtn);
            var w = new ETA.Views.WaybleLoginWindow();
            if (parent is not null) await w.ShowDialog(parent);
            else                    w.Show();
            if (w.LoginSucceeded)
                waybleLoginBtn.Content = "🔑 Wayble 연결됨";
        };

        var hdrGrid = new Grid();
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        Grid.SetColumn(hdrTb, 0);
        Grid.SetColumn(waybleLoginBtn, 1);
        Grid.SetColumn(openXlsxBtn, 2);
        hdrGrid.Children.Add(hdrTb);
        hdrGrid.Children.Add(waybleLoginBtn);
        hdrGrid.Children.Add(openXlsxBtn);
        var hdrBorder = new Border { Background = AppRes("PanelBg"), Child = hdrGrid };
        Grid.SetRow(hdrBorder, 0);

        var body = new StackPanel { Spacing = 1, Margin = new Thickness(0, 4, 0, 8) };
        var bySite = catalog
            .GroupBy(e => (e.OfficeCd, e.SiteCd, e.SiteNm))
            .OrderBy(g => g.Key.OfficeCd).ThenBy(g => g.Key.SiteCd);
        foreach (var grp in bySite)
        {
            var key = grp.Key;
            var groupMatched = grp.Count(e => mapIndex.ContainsKey((e.SiteCd, e.SampleCategory)));
            body.Children.Add(new TextBlock
            {
                Text       = $"{key.OfficeCd} / {key.SiteCd} · {key.SiteNm}   ({groupMatched}/{grp.Count()})",
                FontSize   = AppTheme.FontSM,
                FontWeight = FontWeight.SemiBold,
                FontFamily = Font,
                Foreground = AppTheme.FgInfo,
                Margin     = new Thickness(10, 8, 10, 2),
            });

            foreach (var e in grp.OrderBy(x => x.SampleCategory))
            {
                bool isMatched = mapIndex.TryGetValue((e.SiteCd, e.SampleCategory), out var eta);

                var rowGrid = new Grid { Margin = new Thickness(16, 1, 6, 1) };
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

                var cdTb = new TextBlock
                {
                    Text       = e.SampleCategory,
                    FontSize   = AppTheme.FontXS,
                    FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin     = new Thickness(0, 0, 6, 0),
                    MinWidth   = 50,
                };
                Grid.SetColumn(cdTb, 0);

                var nmPanel = new StackPanel { Spacing = 0, VerticalAlignment = VerticalAlignment.Center };
                nmPanel.Children.Add(new TextBlock
                {
                    Text       = e.SampleCategoryNm,
                    FontSize   = AppTheme.FontBase,
                    FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                });
                if (isMatched)
                {
                    nmPanel.Children.Add(new TextBlock
                    {
                        Text       = $"→ {eta.시설명} / {eta.시료명}",
                        FontSize   = AppTheme.FontXS,
                        FontFamily = Font,
                        Foreground = AppTheme.StatusOkFg,
                    });
                }
                Grid.SetColumn(nmPanel, 1);

                var badge = new Border
                {
                    Background      = isMatched ? AppTheme.StatusOkBg : AppTheme.StatusBadBg,
                    BorderBrush     = isMatched ? AppTheme.StatusOkBorder : AppTheme.StatusBadBorder,
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(999),
                    Padding         = new Thickness(6, 1),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text       = isMatched ? "✓" : "✗",
                        FontSize   = AppTheme.FontXS,
                        FontFamily = Font,
                        Foreground = isMatched ? AppTheme.StatusOkFg : AppTheme.StatusBadFg,
                        FontWeight = FontWeight.SemiBold,
                    },
                };
                Grid.SetColumn(badge, 2);

                rowGrid.Children.Add(cdTb);
                rowGrid.Children.Add(nmPanel);
                rowGrid.Children.Add(badge);

                var border = new Border
                {
                    Padding      = new Thickness(6, 3),
                    CornerRadius = new CornerRadius(4),
                    Background   = Brushes.Transparent,
                    Cursor       = isMatched ? new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand) : null,
                    Child        = rowGrid,
                };
                if (isMatched)
                {
                    var etaLocal = eta;
                    border.PointerPressed += (_, _) =>
                    {
                        _selectedFacilityPair = (etaLocal.시설명, etaLocal.시료명);
                        _show2Content.Content = BuildFacilityMatchEditPanel(etaLocal.시설명, etaLocal.시료명);
                    };
                    border.PointerEntered += (_, _) => border.Background = AppTheme.BgSecondary;
                    border.PointerExited  += (_, _) => border.Background = Brushes.Transparent;
                }
                body.Children.Add(border);
            }
        }

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = body };
        Grid.SetRow(scroll, 1);
        grid.Children.Add(hdrBorder);
        grid.Children.Add(scroll);
        return grid;
    }

    // =========================================================================
    // DB 조작 — 분석항목
    // =========================================================================
    private static void SaveAnalyteToDB(AnalysisItem item, string? originalAnalyte)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            bool exists = false;
            if (!string.IsNullOrWhiteSpace(originalAnalyte))
            {
                using var chk = conn.CreateCommand();
                chk.CommandText = "SELECT COUNT(*) FROM `분석정보` WHERE `Analyte` = @a";
                chk.Parameters.AddWithValue("@a", originalAnalyte);
                exists = Convert.ToInt32(chk.ExecuteScalar()) > 0;
            }
            if (exists)
            {
                if (!string.Equals(originalAnalyte, item.Analyte, StringComparison.Ordinal))
                    RenameTestRecordTable(conn, originalAnalyte!, item.Analyte);
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "UPDATE `분석정보` SET Category=@cat, Analyte=@ana, `약칭`=@alias, `AliasX`=@aliasx, Parts=@parts, DecimalPlaces=@dp, unit=@unit, ES=@es, Method=@meth, instrument=@inst, `정량한계`=@loq WHERE Analyte=@orig";
                cmd.Parameters.AddWithValue("@cat",  item.Category);    cmd.Parameters.AddWithValue("@ana",  item.Analyte);
                cmd.Parameters.AddWithValue("@alias",item.약칭);          cmd.Parameters.AddWithValue("@aliasx", item.AliasX);
                cmd.Parameters.AddWithValue("@parts",item.Parts);        cmd.Parameters.AddWithValue("@dp",   item.DecimalPlaces);
                cmd.Parameters.AddWithValue("@unit", item.unit);         cmd.Parameters.AddWithValue("@es",   item.ES);
                cmd.Parameters.AddWithValue("@meth", item.Method);       cmd.Parameters.AddWithValue("@inst", item.instrument);
                cmd.Parameters.AddWithValue("@loq",  (object?)item.정량한계 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@orig", originalAnalyte);
                cmd.ExecuteNonQuery();
            }
            else
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO `분석정보` (Category, Analyte, `약칭`, `AliasX`, Parts, DecimalPlaces, unit, ES, Method, instrument, `정량한계`) VALUES (@cat, @ana, @alias, @aliasx, @parts, @dp, @unit, @es, @meth, @inst, @loq)";
                cmd.Parameters.AddWithValue("@cat",  item.Category);    cmd.Parameters.AddWithValue("@ana",  item.Analyte);
                cmd.Parameters.AddWithValue("@alias",item.약칭);          cmd.Parameters.AddWithValue("@aliasx", item.AliasX);
                cmd.Parameters.AddWithValue("@parts",item.Parts);        cmd.Parameters.AddWithValue("@dp",   item.DecimalPlaces);
                cmd.Parameters.AddWithValue("@unit", item.unit);         cmd.Parameters.AddWithValue("@es",   item.ES);
                cmd.Parameters.AddWithValue("@meth", item.Method);       cmd.Parameters.AddWithValue("@inst", item.instrument);
                cmd.Parameters.AddWithValue("@loq",  (object?)item.정량한계 ?? DBNull.Value);
                cmd.ExecuteNonQuery();
            }
            AnalysisService.SyncColumnsToAssignmentTable();
            AnalysisService.SyncColumnsToRequestTable();
            AnalysisService.SyncColumnsToContractTable();
            AnalysisService.SyncColumnsToDischargeTable();
            AnalysisService.SyncColumnsToQuotationTable();
            AnalysisService.SyncColumnsToStatementTable();
        }
        catch (Exception ex) { Console.WriteLine($"[SaveAnalyteToDB] {ex.Message}"); }
    }

    private static void RenameTestRecordTable(System.Data.Common.DbConnection conn, string oldAnalyte, string newAnalyte)
    {
        try
        {
            var oldTable = $"{WaterCenterDbMigration.SafeName(oldAnalyte)}_시험기록부";
            var newTable = $"{WaterCenterDbMigration.SafeName(newAnalyte)}_시험기록부";
            if (!DbConnectionFactory.TableExists(conn, oldTable) || DbConnectionFactory.TableExists(conn, newTable)) return;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"RENAME TABLE `{oldTable}` TO `{newTable}`";
            cmd.ExecuteNonQuery();
            Console.WriteLine($"[RenameTestRecordTable] {oldTable} → {newTable}");
        }
        catch (Exception ex) { Console.WriteLine($"[RenameTestRecordTable] {ex.Message}"); }
    }

    private static void DeleteAnalyteFromDB(string analyte)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `분석정보` WHERE Analyte = @a";
            cmd.Parameters.AddWithValue("@a", analyte);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Console.WriteLine($"[DeleteAnalyteFromDB] {ex.Message}"); }
    }

    // =========================================================================
    // 공통 헬퍼
    // =========================================================================
    private static Control MakePhotoCircle(Agent agent, double size)
    {
        if (!string.IsNullOrEmpty(agent.PhotoPath))
        {
            AgentService.EnsurePhotoLocal(agent.사번, agent.PhotoPath);
            var fullPath = System.IO.Path.IsPathRooted(agent.PhotoPath) ? agent.PhotoPath : System.IO.Path.Combine(AgentService.GetPhotoDirectory(), agent.PhotoPath);
            if (System.IO.File.Exists(fullPath))
            {
                try { using var stream = System.IO.File.OpenRead(fullPath); var bmp = new Bitmap(stream); return new Ellipse { Width = size, Height = size, Fill = new ImageBrush(bmp) { Stretch = Stretch.UniformToFill } }; }
                catch { }
            }
        }
        var initial = agent.성명.Length > 0 ? agent.성명[0].ToString() : "?";
        var key = !string.IsNullOrWhiteSpace(agent.직급) ? agent.직급 : agent.성명;
        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(key);
        return new Border { Width = size, Height = size, CornerRadius = new CornerRadius(size / 2), Background = new SolidColorBrush(Color.Parse(ibg)), ClipToBounds = true, Child = new TextBlock { Text = initial, FontSize = size * 0.42, Foreground = new SolidColorBrush(Color.Parse(ifg)), FontFamily = Font, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center } };
    }

    private static TextBlock MakeHint(string text) => new()
    {
        Text = text, FontSize = AppTheme.FontLG, FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        Foreground = new SolidColorBrush(Color.Parse("#666677")), HorizontalAlignment = HorizontalAlignment.Center,
        VerticalAlignment = VerticalAlignment.Center, TextAlignment = Avalonia.Media.TextAlignment.Center, Margin = new Thickness(20),
    };

    private static Control BuildEmptyHint()    => MakeHint("← 직원을 선택하면\n메뉴 접근 권한을 설정할 수 있습니다.");
    private static Control BuildAnalyteHint()  => MakeHint("← 항목을 클릭하면 편집할 수 있습니다.\n+ 항목 추가로 새 분석항목을 등록하세요.");
    private static Control BuildEquipmentHint()=> MakeHint("← 장비를 클릭하면 편집할 수 있습니다.\n+ 장비 추가로 새 분석장비를 등록하세요.");
    private static Control BuildMeasurerContractHint() => MakeHint("📡 스크랩 버튼으로\n측정인.kr 의뢰추가 모달의\n계약·측정대상 사업장·측정시설 목록을 가져옵니다.\n\n※ 스크랩 전 의뢰추가 모달을 먼저 열어두세요.");
    private static Control BuildFacilityMatchHint()    => MakeHint("← 처리시설 목록에서 시설·시료를 선택하면\n외부 시스템(WAYBLE / 측정인.kr) 과의\n매핑을 등록할 수 있습니다.");

    // =========================================================================
    // 탭 4 — 측정인계약 Show1
    // =========================================================================
    private Control BuildMeasurerContractPanel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        var scrapeStatusTb = new TextBlock { FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(8, 2, 8, 4), IsVisible = false, TextWrapping = TextWrapping.Wrap };

        var hdrGrid = new Grid();
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        var hdrTitle = new TextBlock { Text = "측정인 계약·측정대상 사업장", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"), FontFamily = Font, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 6) };
        Grid.SetColumn(hdrTitle, 0);

        var importBtn = new Button { Content = "📥 xlsx Import", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusOkBg, Foreground = AppTheme.StatusOkFg, BorderBrush = AppTheme.StatusOkBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        importBtn.Click += async (_, _) => await RunFacilityImportAsync(importBtn, scrapeStatusTb);
        Grid.SetColumn(importBtn, 1);

        var scrapeBtn = new Button { Content = "📡 스크랩", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = AppTheme.StatusInfoBg, Foreground = AppTheme.StatusInfoFg, BorderBrush = AppTheme.StatusInfoBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold };
        scrapeBtn.Click += async (_, _) =>
        {
            scrapeBtn.IsEnabled = false;
            scrapeStatusTb.IsVisible = true;
            scrapeStatusTb.Foreground = AppRes("FgMuted");
            scrapeStatusTb.Text = "스크랩 중...";
            var (contracts, places, facilities, err) = await MeasurerCdpService.ScrapeContractsAsync(msg =>
                Dispatcher.UIThread.Post(() => { scrapeStatusTb.Text = msg; }));
            scrapeBtn.IsEnabled = true;
            if (err != null) { scrapeStatusTb.Text = "⚠ " + err; scrapeStatusTb.Foreground = Brushes.OrangeRed; }
            else { scrapeStatusTb.Text = $"✓ 계약 {contracts}건 / 측정대상 사업장 {places}건 / 측정시설 {facilities}건 저장"; scrapeStatusTb.Foreground = Brushes.LightGreen; LoadMeasurerContractList(); }
        };
        Grid.SetColumn(scrapeBtn, 2);

        hdrGrid.Children.Add(hdrTitle);
        hdrGrid.Children.Add(importBtn);
        hdrGrid.Children.Add(scrapeBtn);

        var hdrOuter = new StackPanel { Spacing = 0 };
        hdrOuter.Children.Add(hdrGrid);
        hdrOuter.Children.Add(scrapeStatusTb);
        var headerBorder = new Border { Background = AppRes("PanelBg"), Child = hdrOuter };
        Grid.SetRow(headerBorder, 0);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = _measurerContractListPanel };
        Grid.SetRow(scroll, 1);

        grid.Children.Add(headerBorder);
        grid.Children.Add(scroll);
        return grid;
    }

    // ── 측정시설(지역) xlsx Import ───────────────────────────────────────────
    private async System.Threading.Tasks.Task RunFacilityImportAsync(Button btn, TextBlock statusTb)
    {
        try
        {
            btn.IsEnabled = false;
            statusTb.IsVisible = true;
            statusTb.Foreground = AppRes("FgMuted");

            // 파일 피커 — 시작 폴더는 Data/
            var top = TopLevel.GetTopLevel(btn);
            if (top == null)
            {
                statusTb.Text = "⚠ 파일 피커를 열 수 없습니다.";
                statusTb.Foreground = Brushes.OrangeRed;
                btn.IsEnabled = true; return;
            }

            Avalonia.Platform.Storage.IStorageFolder? startFolder = null;
            try
            {
                string dataDir = System.IO.Path.Combine(AppContext.BaseDirectory, "Data");
                if (!Directory.Exists(dataDir)) dataDir = "Data";
                if (Directory.Exists(dataDir))
                    startFolder = await top.StorageProvider.TryGetFolderFromPathAsync(new Uri(System.IO.Path.GetFullPath(dataDir)));
            }
            catch { }

            var files = await top.StorageProvider.OpenFilePickerAsync(new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title = "측정시설(지역)관리 xlsx 선택",
                AllowMultiple = false,
                SuggestedStartLocation = startFolder,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx" } }
                }
            });
            if (files.Count == 0) { statusTb.Text = "취소됨"; btn.IsEnabled = true; return; }
            string path = files[0].Path.LocalPath;
            if (string.IsNullOrWhiteSpace(path))
            {
                statusTb.Text = "⚠ 파일 경로를 읽을 수 없습니다.";
                statusTb.Foreground = Brushes.OrangeRed;
                btn.IsEnabled = true; return;
            }

            statusTb.Text = $"Import 중: {System.IO.Path.GetFileName(path)} …";

            // 3) 백그라운드 실행
            var result = await System.Threading.Tasks.Task.Run(() => MeasurerService.ImportFacilitiesFromXlsx(path));

            if (result.Error != null)
            {
                statusTb.Text = "⚠ " + result.Error;
                statusTb.Foreground = Brushes.OrangeRed;
            }
            else
            {
                statusTb.Text = $"✓ {System.IO.Path.GetFileName(path)} — 행 {result.TotalRows} · 계약 {result.ContractsUpsert} · 채취지점 {result.PointsUpsert}" +
                                (result.Skipped > 0 ? $" · skip {result.Skipped}" : "");
                statusTb.Foreground = Brushes.LightGreen;
                LoadMeasurerContractList();
            }
        }
        catch (Exception ex)
        {
            statusTb.Text = "⚠ Import 실패: " + ex.Message;
            statusTb.Foreground = Brushes.OrangeRed;
        }
        finally { btn.IsEnabled = true; }
    }

    private void LoadMeasurerContractList()
    {
        _measurerContractListPanel.Children.Clear();
        try
        {
            var list = MeasurerService.GetAllMeasurerContracts();
            if (list.Count == 0)
            {
                _measurerContractListPanel.Children.Add(new TextBlock
                {
                    Text         = "스크랩된 계약이 없습니다.\n📡 스크랩 버튼을 눌러\n측정인.kr 드롭다운을 가져오세요.",
                    FontSize     = AppTheme.FontBase,
                    FontFamily   = Font,
                    Foreground   = AppRes("FgMuted"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin       = new Thickness(12, 16),
                });
                return;
            }
            foreach (var c in list)
            {
                var placeCount = MeasurerService.GetMeasurerPlaces(c.계약번호).Count;
                _measurerContractListPanel.Children.Add(CreateMeasurerContractRow(c, placeCount));
            }
        }
        catch (Exception ex) { Console.WriteLine($"[LoadMeasurerContractList] {ex.Message}"); }
    }

    private Border CreateMeasurerContractRow(MeasurerService.MeasurerContract c, int placeCount)
    {
        var row = new Grid { Margin = new Thickness(4, 2) };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var titleBlock = new TextBlock { Text = c.계약번호, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center };
        var subBlock   = new TextBlock { Text = string.IsNullOrWhiteSpace(c.업체명) ? "(업체명 없음)" : c.업체명, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.NoWrap, TextTrimming = TextTrimming.CharacterEllipsis, Margin = new Thickness(0, 2) };
        var info = new StackPanel { Spacing = 1 };
        info.Children.Add(titleBlock); info.Children.Add(subBlock);
        Grid.SetColumn(info, 0);
        row.Children.Add(info);

        var countBadge = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#0b4f6c")),
            CornerRadius = new CornerRadius(8),
            Padding      = new Thickness(6, 2),
            VerticalAlignment = VerticalAlignment.Center,
            Child        = new TextBlock { Text = placeCount + "곳", FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brushes.White },
        };
        Grid.SetColumn(countBadge, 1);
        row.Children.Add(countBadge);

        var border = new Border
        {
            Padding      = new Thickness(8, 6),
            CornerRadius = new CornerRadius(4),
            Background   = Brushes.Transparent,
            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child        = row,
        };
        border.PointerPressed += (_, _) =>
        {
            _show2Content.Content = BuildMeasurerContractDetailPanel(c);
        };
        border.PointerEntered += (_, _) => border.Background = AppTheme.BgSecondary;
        border.PointerExited  += (_, _) => border.Background = Brushes.Transparent;
        return border;
    }

    // =========================================================================
    // 탭 4 — 측정인계약 Show2 (상세: 처리시설 목록)
    // =========================================================================
    private Control BuildMeasurerContractDetailPanel(MeasurerService.MeasurerContract c)
    {
        var places = MeasurerService.GetMeasurerPlaces(c.계약번호);

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock { Text = $"🧾 {c.계약번호}", FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 2) });
        root.Children.Add(new TextBlock { Text = c.업체명, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 6), TextWrapping = TextWrapping.Wrap });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 8) });

        void AddField(string label, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            var s = new StackPanel { Spacing = 0, Margin = new Thickness(0, 2) };
            s.Children.Add(new TextBlock { Text = label, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted") });
            s.Children.Add(new TextBlock { Text = value, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), TextWrapping = TextWrapping.Wrap });
            root.Children.Add(s);
        }
        AddField("대표사업장", c.대표사업장);
        AddField("계약기간", c.계약기간);

        root.Children.Add(new TextBlock { Text = $"측정대상 사업장 ({places.Count}곳)", FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppRes("AppFg"), Margin = new Thickness(0, 12, 0, 4) });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderDefault, Margin = new Thickness(0, 0, 0, 4) });

        if (places.Count == 0)
        {
            root.Children.Add(new TextBlock { Text = "(스크랩된 측정대상 사업장 없음)", FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 4) });
        }
        else
        {
            foreach (var p in places)
            {
                var headerLine = new Grid { Margin = new Thickness(0, 2) };
                headerLine.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
                headerLine.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
                var nameTb = new TextBlock { Text = p.처리시설명, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };
                var codeTb = new TextBlock { Text = p.처리시설코드, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center };
                Grid.SetColumn(nameTb, 0); Grid.SetColumn(codeTb, 1);
                headerLine.Children.Add(nameTb); headerLine.Children.Add(codeTb);

                var placeBlock = new StackPanel { Spacing = 2, Margin = new Thickness(0, 4, 0, 6) };
                placeBlock.Children.Add(new Border { Padding = new Thickness(6, 4), Child = headerLine });

                // 하위 측정시설 목록
                try
                {
                    var facilities = MeasurerService.GetMeasurerFacilities(c.계약번호, p.처리시설코드);
                    if (facilities.Count == 0)
                    {
                        placeBlock.Children.Add(new TextBlock { Text = "    (측정시설 없음)", FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 2) });
                    }
                    else
                    {
                        foreach (var f in facilities)
                        {
                            placeBlock.Children.Add(new TextBlock
                            {
                                Text       = "    · " + f.측정시설명,
                                FontSize   = AppTheme.FontSM,
                                FontFamily = Font,
                                Foreground = AppRes("AppFg"),
                                Margin     = new Thickness(0, 0, 0, 1),
                            });
                        }
                    }
                }
                catch (Exception ex) { Console.WriteLine($"[MeasurerFacilities] {ex.Message}"); }

                root.Children.Add(placeBlock);
            }
        }

        var btnDelete = new Button { Content = "🗑 이 계약 삭제", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(12, 6), Margin = new Thickness(0, 14, 0, 4), Background = AppTheme.StatusBadBg, Foreground = AppTheme.StatusBadFg, BorderBrush = AppTheme.StatusBadBorder, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Left };
        btnDelete.Click += (_, _) =>
        {
            try
            {
                MeasurerService.DeleteMeasurerContract(c.계약번호);
                LoadMeasurerContractList();
                _show2Content.Content = BuildMeasurerContractHint();
            }
            catch (Exception ex) { Console.WriteLine($"[DeleteMeasurerContract] {ex.Message}"); }
        };
        root.Children.Add(btnDelete);

        return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = root };
    }
}
