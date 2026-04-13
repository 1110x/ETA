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

namespace ETA.Views.Pages.Common;

/// <summary>
/// 관리자 페이지
///   탭 1 — 권한관리: 직원별 메뉴 접근 설정
///   탭 2 — 분석항목 관리: 분석정보 CRUD + 시험기록부 테이블 생성
///   탭 3 — 분석장비 관리: 측정인_분석장비 테이블 수동 CRUD
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

    // ── 공통 상태 ─────────────────────────────────────────────────────────────
    private readonly ContentControl _show2Content;
    private readonly ContentControl _show1Switcher;
    private string _mode = "access"; // "access" | "analyte" | "equipment"

    // ── 탭 버튼 ───────────────────────────────────────────────────────────────
    private Border? _tabAccess;
    private Border? _tabAnalyte;
    private Border? _tabEquipment;

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

    // ── 패널 캐시 ─────────────────────────────────────────────────────────────
    private readonly Control _employeePanelRef;
    private readonly Control _analytePanelRef;
    private readonly Control _equipmentPanelRef;

    // =========================================================================
    public AccessPage()
    {
        AccessService.EnsureTable();
        MeasurerService.EnsureEquipmentTable();

        _treeView          = BuildEmployeeTreeView();
        _analyteTree       = BuildAnalyteTreeView();
        _equipmentListPanel = new StackPanel { Spacing = 1 };

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

        _employeePanelRef  = BuildEmployeePanel();
        _analytePanelRef   = BuildAnalytePanel();
        _equipmentPanelRef = BuildEquipmentPanel();

        Show1 = BuildShow1Shell();
        Show2 = new Border { Padding = new Thickness(6), Child = _show2Content };

        SwitchMode("access");
        LoadEmployeeTree();
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
        _tabAccess    = MakeTabBtn("🔐 권한관리",    active: true,  () => SwitchMode("access"));
        _tabAnalyte   = MakeTabBtn("🧪 분석항목",    active: false, () => SwitchMode("analyte"));
        _tabEquipment = MakeTabBtn("🔧 분석장비",    active: false, () => SwitchMode("equipment"));
        stack.Children.Add(_tabAccess);
        stack.Children.Add(_tabAnalyte);
        stack.Children.Add(_tabEquipment);
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
        UpdateTabStyle(_tabAccess,    mode == "access");
        UpdateTabStyle(_tabAnalyte,   mode == "analyte");
        UpdateTabStyle(_tabEquipment, mode == "equipment");

        switch (mode)
        {
            case "analyte":
                _show1Switcher.Content = _analytePanelRef;
                _show2Content.Content  = BuildAnalyteHint();
                LoadAnalyteTree();
                break;
            case "equipment":
                _show1Switcher.Content = _equipmentPanelRef;
                _show2Content.Content  = BuildEquipmentHint();
                LoadEquipmentList();
                break;
            default:
                _show1Switcher.Content = _employeePanelRef;
                _show2Content.Content  = BuildEmptyHint();
                break;
        }
    }

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
        var addBtn = new Button { Content = "+ 항목 추가", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = new SolidColorBrush(Color.Parse("#1a6e2e")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
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
            var groups = items.GroupBy(a => string.IsNullOrWhiteSpace(a.Category) ? "기타" : a.Category).OrderBy(g => g.Key);
            foreach (var grp in groups)
            {
                var groupNode = new TreeViewItem
                {
                    Header = new TextBlock { Text = grp.Key, FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold, FontFamily = Font, Foreground = AppTheme.FgInfo, Margin = new Thickness(2, 2) },
                    IsExpanded = true,
                };
                foreach (var item in grp.OrderBy(i => i.ES))
                    groupNode.Items.Add(CreateAnalyteTreeItem(item));
                _analyteTree.Items.Add(groupNode);
            }
        }
        catch (Exception ex) { Console.WriteLine($"[LoadAnalyteTree] {ex.Message}"); }
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
        var tbCategory = MkTb(item?.Category ?? "",                    "카테고리 (예: 유기물류)");
        var tbUnit     = MkTb(item?.unit ?? "",                        "단위 (예: mg/L)");
        var tbDP       = MkTb((item?.DecimalPlaces ?? 0).ToString(),   "소수점 자리 (예: 1)");
        var tbES       = MkTb(item?.ES ?? "",                          "정렬번호 (예: 01)");
        var tbParts    = MkTb(item?.Parts ?? "",                       "구성 (예: 1)");
        var tbMethod   = MkTb(item?.Method ?? "",                      "시험방법 (예: 자외선/가시선분광법)");
        var tbInst     = MkTb(item?.instrument ?? "",                  "기기 (예: ICP-OES)");

        // ─ 기기 선택 드롭다운 ─────────────────────────────────────────
        var equipments = new List<(string 장비명, string 코드값)>();
        try { equipments = MeasurerService.GetAllEquipments(); } catch { }

        Control instField;
        if (equipments.Count > 0)
        {
            var combo = new ComboBox
            {
                PlaceholderText = "등록 장비에서 선택 →",
                FontSize        = AppTheme.FontSM,
                FontFamily      = Font,
                Height          = 26,
                Margin          = new Thickness(0, 2, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Background      = AppRes("PanelBg"),
                BorderBrush     = AppRes("InputBorder"),
                BorderThickness = new Thickness(1),
            };
            foreach (var (name, _) in equipments)
                combo.Items.Add(name);

            // 현재 값과 일치하는 항목 선택
            if (!string.IsNullOrWhiteSpace(item?.instrument))
            {
                var match = equipments.FirstOrDefault(e => item.instrument.Contains(e.장비명));
                if (match.장비명 != null)
                    combo.SelectedItem = match.장비명;
            }
            combo.SelectionChanged += (_, _) =>
            {
                if (combo.SelectedItem is string name && !string.IsNullOrWhiteSpace(name))
                    tbInst.Text = name;
            };

            instField = combo;
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
        form.Children.Add(TwoCol(Labeled("단위", tbUnit), Labeled("소수점 자리", tbDP)));
        form.Children.Add(TwoCol(Labeled("정렬번호 (ES)", tbES), Labeled("구성 (Parts)", tbParts)));
        form.Children.Add(Labeled("시험방법", tbMethod));
        form.Children.Add(Labeled("기기 (instrument)", instField));

        // ─ 버튼 행 ───────────────────────────────────────────────────
        var btnSave = new Button { Content = isNew ? "✚ 신규 저장" : "💾 저장", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = new SolidColorBrush(Color.Parse("#1565c0")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
        btnSave.Click += (_, _) =>
        {
            var updated = new AnalysisItem
            {
                Category      = tbCategory.Text?.Trim() ?? "",
                Analyte       = tbAnalyte.Text?.Trim() ?? "",
                약칭           = tbAlias.Text?.Trim() ?? "",
                Parts         = tbParts.Text?.Trim() ?? "",
                DecimalPlaces = int.TryParse(tbDP.Text?.Trim(), out var dp) ? dp : 0,
                unit          = tbUnit.Text?.Trim() ?? "",
                ES            = tbES.Text?.Trim() ?? "",
                Method        = tbMethod.Text?.Trim() ?? "",
                instrument    = tbInst.Text?.Trim() ?? "",
            };
            if (string.IsNullOrWhiteSpace(updated.Analyte)) return;
            SaveAnalyteToDB(updated, isNew ? null : origAnalyte);
            LoadAnalyteTree();
            _selectedAnalyte = updated;
            _show2Content.Content = BuildAnalyteEditPanel(updated, isNew: false);
        };

        var btnDelete = new Button { Content = "🗑 삭제", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = new SolidColorBrush(Color.Parse("#c62828")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4), IsVisible = !isNew };
        btnDelete.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(origAnalyte)) return;
            DeleteAnalyteFromDB(origAnalyte);
            LoadAnalyteTree();
            _selectedAnalyte = null;
            _show2Content.Content = BuildAnalyteHint();
        };

        var btnCreateTable = new Button { Content = "🗄 시험기록부 테이블 생성", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = new SolidColorBrush(Color.Parse("#4a148c")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
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
            }
            else
                stack.Children.Add(new TextBlock { Text = "⚠ 시험기록부 테이블 없음 — 위 버튼으로 생성하세요", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#f57f17")), TextWrapping = TextWrapping.Wrap });
        }
        catch (Exception ex) { stack.Children.Add(new TextBlock { Text = $"오류: {ex.Message}", FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = Brushes.Red }); }
        return stack;
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
                del.PointerPressed += (_, _) => { if (idx < _formulaTokens.Count) _formulaTokens.RemoveAt(idx); RefreshFormulaDisplay(panel, analyte); };
                inner.Children.Add(del);
                var tokenBorder = new Border { Background = new SolidColorBrush(Color.Parse(bg)), CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 3), Margin = new Thickness(0, 2, 4, 2), Child = inner };
                tokenBorder.PointerEntered += (_, _) => del.IsVisible = true;
                tokenBorder.PointerExited  += (_, _) => del.IsVisible = false;
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
        var hdrTitle = new TextBlock { Text = "분석장비 목록", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"), FontFamily = Font, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 6) };
        Grid.SetColumn(hdrTitle, 0);
        var scrapeBtn = new Button { Content = "📡 스크랩", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = new SolidColorBrush(Color.Parse("#0b4f6c")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
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
        var addBtn = new Button { Content = "+ 추가", FontSize = AppTheme.FontSM, FontFamily = Font, Padding = new Thickness(8, 4), Margin = new Thickness(4), Background = new SolidColorBrush(Color.Parse("#1a6e2e")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
        addBtn.Click += (_, _) => { _selectedEquipment = null; _show2Content.Content = BuildEquipmentEditPanel(null, null, isNew: true); };
        Grid.SetColumn(addBtn, 2);
        hdrGrid.Children.Add(hdrTitle); hdrGrid.Children.Add(scrapeBtn); hdrGrid.Children.Add(addBtn);
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
        var row = new Grid { Margin = new Thickness(4, 2) };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var nameBlock = new TextBlock { Text = name, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center };
        var codeBlock = new TextBlock { Text = code, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 2) };
        var info = new StackPanel { Spacing = 1 };
        info.Children.Add(nameBlock); info.Children.Add(codeBlock);
        Grid.SetColumn(info, 0);
        row.Children.Add(info);

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
            _selectedEquipment = (name, code);
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

        var btnSave = new Button { Content = isNew ? "✚ 저장" : "💾 저장", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = new SolidColorBrush(Color.Parse("#1565c0")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4) };
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
                LoadEquipmentList();
                _selectedEquipment = (n, c);
                _show2Content.Content = BuildEquipmentEditPanel(n, c, isNew: false);
            }
            catch (Exception ex) { statusTb.Text = $"오류: {ex.Message}"; statusTb.Foreground = Brushes.Red; statusTb.IsVisible = true; }
        };

        var btnDelete = new Button { Content = "🗑 삭제", FontSize = AppTheme.FontBase, FontFamily = Font, Padding = new Thickness(14, 6), Background = new SolidColorBrush(Color.Parse("#c62828")), Foreground = Brushes.White, CornerRadius = new CornerRadius(4), IsVisible = !isNew };
        btnDelete.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(origCode)) return;
            try { MeasurerService.DeleteEquipment(origCode); LoadEquipmentList(); _selectedEquipment = null; _show2Content.Content = BuildEquipmentHint(); }
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
                cmd.CommandText = "UPDATE `분석정보` SET Category=@cat, Analyte=@ana, `약칭`=@alias, Parts=@parts, DecimalPlaces=@dp, unit=@unit, ES=@es, Method=@meth, instrument=@inst WHERE Analyte=@orig";
                cmd.Parameters.AddWithValue("@cat",  item.Category);    cmd.Parameters.AddWithValue("@ana",  item.Analyte);
                cmd.Parameters.AddWithValue("@alias",item.약칭);          cmd.Parameters.AddWithValue("@parts",item.Parts);
                cmd.Parameters.AddWithValue("@dp",   item.DecimalPlaces); cmd.Parameters.AddWithValue("@unit", item.unit);
                cmd.Parameters.AddWithValue("@es",   item.ES);           cmd.Parameters.AddWithValue("@meth", item.Method);
                cmd.Parameters.AddWithValue("@inst", item.instrument);   cmd.Parameters.AddWithValue("@orig", originalAnalyte);
                cmd.ExecuteNonQuery();
            }
            else
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO `분석정보` (Category, Analyte, `약칭`, Parts, DecimalPlaces, unit, ES, Method, instrument) VALUES (@cat, @ana, @alias, @parts, @dp, @unit, @es, @meth, @inst)";
                cmd.Parameters.AddWithValue("@cat",  item.Category);    cmd.Parameters.AddWithValue("@ana",  item.Analyte);
                cmd.Parameters.AddWithValue("@alias",item.약칭);          cmd.Parameters.AddWithValue("@parts",item.Parts);
                cmd.Parameters.AddWithValue("@dp",   item.DecimalPlaces); cmd.Parameters.AddWithValue("@unit", item.unit);
                cmd.Parameters.AddWithValue("@es",   item.ES);           cmd.Parameters.AddWithValue("@meth", item.Method);
                cmd.Parameters.AddWithValue("@inst", item.instrument);
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
}
