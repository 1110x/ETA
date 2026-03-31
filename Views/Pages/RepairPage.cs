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
using System.Linq;

namespace ETA.Views.Pages;

/// <summary>
/// 보수요청 페이지.
/// TreeControl  → Content1 (년/월 트리)
/// ListControl  → Content2 (월별 목록, 클릭 시 폼에 로드)
/// FormControl  → Content3 (신규/수정 폼)
/// </summary>
public class RepairPage
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── 컨트롤 ───────────────────────────────────────────────────────────
    public Control TreeControl { get; }
    public Control ListControl { get; }
    public Control FormControl { get; }

    // ── 내부 상태 ─────────────────────────────────────────────────────────
    private readonly TreeView  _tree    = new() { Background = Brushes.Transparent, BorderThickness = new Thickness(0) };
    private readonly StackPanel _listPanel = new() { Spacing = 2 };
    private readonly TextBlock  _listTitle = new() { FontSize = 11, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#777")), Margin = new Thickness(4, 0, 0, 6) };
    private Border? _selectedRowBorder;

    // 폼 필드
    private readonly ComboBox  _cmbCategory;
    private readonly TextBox   _txbDevice;
    private readonly TextBox   _txbSymptom;
    private readonly TextBox   _txbLocation;
    private readonly TextBox   _txbRequester;
    private readonly TextBox   _txbDate;
    private readonly ComboBox  _cmbStatus;
    private readonly TextBox   _txbProcess;
    private readonly TextBox   _txbNote;
    private readonly Button    _btnSave;
    private readonly Button    _btnNew;
    private readonly TextBlock _txbFormTitle;

    private RepairItem? _editingItem;  // null = 신규, 값 있음 = 수정

    public RepairPage()
    {
        // 폼 위젯 초기화
        _cmbCategory  = MakeCombo(new[] { "장비", "시설", "차량", "IT", "기타" });
        _txbDevice    = MakeTxb("장비/시설명");
        _txbSymptom   = MakeTxb("증상 설명");
        _txbLocation  = MakeTxb("위치");
        _txbRequester = MakeTxb("요청자");
        _txbDate      = MakeTxb(""); _txbDate.Text = DateTime.Today.ToString("yyyy-MM-dd");
        _cmbStatus    = MakeCombo(new[] { "대기", "진행중", "완료", "반려" });
        _txbProcess   = MakeTxb("처리 내용");
        _txbNote      = MakeTxb("비고");
        _txbFormTitle = new TextBlock { Text = "신규 보수요청", FontSize = 13, FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#e0e0e0")), Margin = new Thickness(0,0,0,8) };
        _btnSave = MakeBtn("💾 저장", "#2a4a2a", "#aef0ae");
        _btnNew  = MakeBtn("✚ 신규", "#2a2a3a", "#aaa");
        _btnSave.Click += OnSave;
        _btnNew.Click  += (_, _) => ClearForm();

        TreeControl = BuildTree();
        ListControl = BuildList();
        FormControl = BuildForm();

        Refresh();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  공개 API
    // ══════════════════════════════════════════════════════════════════════
    public void Refresh() => LoadTree();

    // ══════════════════════════════════════════════════════════════════════
    //  트리 빌드 (Content1)
    // ══════════════════════════════════════════════════════════════════════
    private Control BuildTree()
    {
        var header = new TextBlock { Text = "🔧  보수요청", FontSize = 13, FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#e0e0e0")), Margin = new Thickness(0,0,0,8) };
        var scroll = new ScrollViewer { Content = _tree, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
        var grid = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        Grid.SetRow(header, 0); Grid.SetRow(scroll, 1);
        grid.Children.Add(header); grid.Children.Add(scroll);
        return grid;
    }

    private void LoadTree()
    {
        _tree.Items.Clear();
        var summary = RepairService.GetMonthSummary();
        int thisYear  = DateTime.Today.Year;
        int thisMonth = DateTime.Today.Month;

        var byYear = summary.GroupBy(s => s.Year).OrderByDescending(g => g.Key);
        foreach (var yg in byYear)
        {
            var yn = new TreeViewItem
            {
                Header = Row("📁", $"{yg.Key}년", "#a0d4a0", 13, bold: true),
                IsExpanded = true,
            };
            foreach (var (y, m, cnt) in yg.OrderByDescending(x => x.Month))
            {
                bool isCur = y == thisYear && m == thisMonth;
                var mn = new TreeViewItem
                {
                    Header = Row(isCur ? "📂" : "📁",
                                 $"  {m:D2}월  ({cnt}건)", isCur ? "#d4d4a0" : "#999977", 12),
                    Tag = (y, m),
                    IsExpanded = isCur,
                };
                mn.PropertyChanged += (s, e) =>
                {
                    if (e.Property == TreeViewItem.IsSelectedProperty
                        && s is TreeViewItem tvi && tvi.IsSelected
                        && tvi.Tag is ValueTuple<int,int> tag)
                        LoadList(tag.Item1, tag.Item2);
                };
                yn.Items.Add(mn);
            }
            _tree.Items.Add(yn);
        }

        // 현재 월 자동 로드
        LoadList(thisYear, thisMonth);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  목록 빌드 (Content2)
    // ══════════════════════════════════════════════════════════════════════
    private Control BuildList()
    {
        var divider = new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#333")), Margin = new Thickness(0,4,0,6) };
        var scroll  = new ScrollViewer { Content = _listPanel, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
        var grid = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(_listTitle, 0); Grid.SetRow(divider, 1); Grid.SetRow(scroll, 2);
        grid.Children.Add(_listTitle); grid.Children.Add(divider); grid.Children.Add(scroll);
        return grid;
    }

    private void LoadList(int year, int month)
    {
        _listPanel.Children.Clear();
        _selectedRowBorder = null;
        var items = RepairService.GetByMonth(year, month);
        _listTitle.Text = $"{year}년 {month:D2}월  —  {items.Count}건";

        if (items.Count == 0)
        {
            _listPanel.Children.Add(new TextBlock { Text = "보수요청 없음", FontSize = 11, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#555")), Margin = new Thickness(4,8) });
            return;
        }

        foreach (var item in items)
            _listPanel.Children.Add(MakeListRow(item));
    }

    private Border MakeListRow(RepairItem item)
    {
        var statusColor = item.상태 switch
        {
            "완료" => "#4a8a4a",
            "진행중" => "#4a6a8a",
            "반려" => "#8a4a4a",
            _ => "#555577",
        };

        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,60,50") };

        // 장비명 + 증상
        var info = new StackPanel { Orientation = Orientation.Vertical, Spacing = 2, [Grid.ColumnProperty] = 0 };
        info.Children.Add(new TextBlock { Text = item.장비명, FontSize = 12, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#ddd")), TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis });
        info.Children.Add(new TextBlock { Text = $"{item.구분}  {item.요청일:MM-dd}  {item.요청자}", FontSize = 10, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#666")) });
        if (!string.IsNullOrEmpty(item.증상))
            info.Children.Add(new TextBlock { Text = item.증상, FontSize = 10, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#888")), TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis });
        grid.Children.Add(info);

        // 상태 배지
        grid.Children.Add(new Border
        {
            Background = new SolidColorBrush(Color.Parse(statusColor)),
            CornerRadius = new CornerRadius(3), Padding = new Thickness(4,1),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock { Text = item.상태, FontSize = 10, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#eee")) },
            [Grid.ColumnProperty] = 2,
        });

        var border = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1e1e2e")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 7),
            Margin = new Thickness(0,0,0,2),
            Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child = grid,
        };

        border.PointerEntered += (_, _) =>
        {
            if (border != _selectedRowBorder)
                border.Background = new SolidColorBrush(Color.Parse("#252535"));
        };
        border.PointerExited += (_, _) =>
        {
            if (border != _selectedRowBorder)
                border.Background = new SolidColorBrush(Color.Parse("#1e1e2e"));
        };

        // ★ 클릭 → 폼에 내용 로드 (터널링: 자식이 이벤트를 소비해도 받음)
        border.AddHandler(
            Avalonia.Input.InputElement.PointerPressedEvent,
            (object? _, Avalonia.Input.PointerPressedEventArgs ev) =>
            {
                if (ev.GetCurrentPoint(border).Properties.IsLeftButtonPressed)
                {
                    // 이전 선택 해제
                    if (_selectedRowBorder != null)
                        _selectedRowBorder.Background = new SolidColorBrush(Color.Parse("#1e1e2e"));
                    _selectedRowBorder = border;
                    border.Background = new SolidColorBrush(Color.Parse("#2a3a5a"));
                    LoadToForm(item);
                }
            },
            Avalonia.Interactivity.RoutingStrategies.Tunnel);

        return border;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  폼 빌드 (Content3)
    // ══════════════════════════════════════════════════════════════════════
    private Control BuildForm()
    {
        var divider = new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#333")), Margin = new Thickness(0,4,0,8) };

        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("72,*"),
            RowDefinitions    = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,Auto,Auto,Auto,Auto"),
        };

        void Add(string label, Control ctrl, int row)
        {
            var lbl = new TextBlock { Text = label, FontSize = 11, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#888")), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0,0,0,4) };
            Grid.SetRow(lbl,  row); Grid.SetColumn(lbl,  0);
            Grid.SetRow(ctrl, row); Grid.SetColumn(ctrl, 1);
            grid.Children.Add(lbl); grid.Children.Add(ctrl);
        }

        Add("구분",   _cmbCategory,  0);
        Add("장비명", _txbDevice,    1);
        Add("증상",   _txbSymptom,   2);
        Add("위치",   _txbLocation,  3);
        Add("요청자", _txbRequester, 4);
        Add("요청일", _txbDate,      5);
        Add("상태",   _cmbStatus,    6);
        Add("처리",   _txbProcess,   7);
        Add("비고",   _txbNote,      8);

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0,10,0,0) };
        btnRow.Children.Add(_btnNew);
        btnRow.Children.Add(_btnSave);

        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Content = new StackPanel { Spacing = 0, Children = { grid, btnRow } },
        };

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(_txbFormTitle, 0); Grid.SetRow(divider, 1); Grid.SetRow(scroll, 2);
        root.Children.Add(_txbFormTitle); root.Children.Add(divider); root.Children.Add(scroll);
        return root;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  폼 로직
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>목록 아이템 클릭 → 폼에 기존 데이터 로드 (수정 모드)</summary>
    private void LoadToForm(RepairItem item)
    {
        _editingItem = item;
        _txbFormTitle.Text = $"✏  수정 — {item.장비명}";

        SelectCombo(_cmbCategory,  item.구분);
        _txbDevice.Text    = item.장비명;
        _txbSymptom.Text   = item.증상;
        _txbLocation.Text  = item.위치;
        _txbRequester.Text = item.요청자;
        _txbDate.Text      = item.요청일.ToString("yyyy-MM-dd");
        SelectCombo(_cmbStatus,    item.상태);
        _txbProcess.Text   = item.처리내용;
        _txbNote.Text      = item.비고;
    }

    /// <summary>폼 초기화 (신규 모드)</summary>
    private void ClearForm()
    {
        _editingItem = null;
        _txbFormTitle.Text = "신규 보수요청";
        _cmbCategory.SelectedIndex  = 0;
        _txbDevice.Text    = "";
        _txbSymptom.Text   = "";
        _txbLocation.Text  = "";
        _txbRequester.Text = "";
        _txbDate.Text      = DateTime.Today.ToString("yyyy-MM-dd");
        _cmbStatus.SelectedIndex = 0;
        _txbProcess.Text   = "";
        _txbNote.Text      = "";
    }

    private void OnSave(object? sender, RoutedEventArgs e)
    {
        var item = _editingItem ?? new RepairItem();
        item.구분     = (_cmbCategory.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
        item.장비명   = _txbDevice.Text    ?? "";
        item.증상     = _txbSymptom.Text   ?? "";
        item.위치     = _txbLocation.Text  ?? "";
        item.요청자   = _txbRequester.Text ?? "";
        item.상태     = (_cmbStatus.SelectedItem  as ComboBoxItem)?.Content?.ToString() ?? "대기";
        item.처리내용 = _txbProcess.Text   ?? "";
        item.비고     = _txbNote.Text      ?? "";
        if (DateTime.TryParse(_txbDate.Text, out var dt)) item.요청일 = dt;

        bool ok = _editingItem == null
            ? RepairService.Insert(item)
            : RepairService.Update(item);

        Debug.WriteLine(ok
            ? $"[Repair] {(_editingItem == null ? "신규" : "수정")} 완료 → {item.장비명}"
            : "[Repair] 저장 실패");

        if (ok) { ClearForm(); Refresh(); }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  서브메뉴 버튼 연동
    // ══════════════════════════════════════════════════════════════════════
    public void ApproveSelected() => UpdateSelectedStatus("진행중");
    public void RejectSelected()  => UpdateSelectedStatus("반려");
    public void CompleteSelected() => UpdateSelectedStatus("완료");
    public void DeleteSelected()
    {
        if (_editingItem == null) return;
        if (RepairService.Delete(_editingItem.Id))
        { ClearForm(); Refresh(); }
    }

    private void UpdateSelectedStatus(string status)
    {
        if (_editingItem == null) return;
        if (RepairService.UpdateStatus(_editingItem.Id, status))
        {
            _editingItem.상태 = status;
            SelectCombo(_cmbStatus, status);
            Refresh();
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static object Row(string icon, string text, string color, double size, bool bold = false)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        sp.Children.Add(new TextBlock { Text = icon, FontSize = size, VerticalAlignment = VerticalAlignment.Center });
        sp.Children.Add(new TextBlock { Text = text, FontSize = size, FontFamily = Font, FontWeight = bold ? FontWeight.Bold : FontWeight.Normal, Foreground = new SolidColorBrush(Color.Parse(color)), VerticalAlignment = VerticalAlignment.Center });
        return sp;
    }

    private static TextBox MakeTxb(string watermark) => new()
    {
        Watermark = watermark, Height = 26, FontSize = 12, FontFamily = Font,
        Background = new SolidColorBrush(Color.Parse("#2a2a35")),
        Foreground = new SolidColorBrush(Color.Parse("#ddd")),
        BorderBrush = new SolidColorBrush(Color.Parse("#444")),
        Padding = new Thickness(6,0), CornerRadius = new CornerRadius(4),
        Margin = new Thickness(0,0,0,4),
    };

    private static ComboBox MakeCombo(string[] items)
    {
        var cb = new ComboBox
        {
            Height = 26, FontSize = 12, FontFamily = Font,
            Background  = new SolidColorBrush(Color.Parse("#2a2a35")),
            Foreground  = new SolidColorBrush(Color.Parse("#ddd")),
            BorderBrush = new SolidColorBrush(Color.Parse("#444")),
            Padding = new Thickness(6,0), CornerRadius = new CornerRadius(4),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Margin = new Thickness(0,0,0,4),
        };
        foreach (var i in items) cb.Items.Add(new ComboBoxItem { Content = i });
        cb.SelectedIndex = 0;
        return cb;
    }

    private static Button MakeBtn(string text, string bg, string fg) => new()
    {
        Content = text, Height = 28, FontSize = 12, FontFamily = Font,
        Background = new SolidColorBrush(Color.Parse(bg)),
        Foreground = new SolidColorBrush(Color.Parse(fg)),
        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        Padding = new Thickness(14,0),
    };

    private static void SelectCombo(ComboBox cb, string value)
    {
        for (int i = 0; i < cb.Items.Count; i++)
            if ((cb.Items[i] as ComboBoxItem)?.Content?.ToString() == value)
            { cb.SelectedIndex = i; return; }
        cb.SelectedIndex = 0;
    }
}
