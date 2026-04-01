using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Models;
using ETA.Services.SERVICE1;
using ETA.Services.Common;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 출장 / 일정 관리 페이지
///   TreeControl     → Show1 (직원 / 계약업체 트리)
///   CalendarControl → Show2 (월별 큰 캘린더)
///   FormControl     → Show3 (일정 작성 폼)
/// </summary>
public class SchedulePage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    // 트리 노드 태그
    private record TreeTag(string Type, string Name, string Id);

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public Control TreeControl     { get; }
    public Control CalendarControl { get; }
    public Control FormControl     { get; }

    // ── 상태 ────────────────────────────────────────────────────────────────
    private DateTime _month    = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    private string?  _selDate;
    private string   _selCat   = "출장";
    private string   _selSite  = "여수";
    private bool     _showCon  = false;      // false=직원, true=계약업체
    private string   _agtName  = "";
    private string   _agtId    = "";

    // ── UI 참조 ──────────────────────────────────────────────────────────────
    private TreeView?   _tree;
    private TextBox?    _searchBox;
    private TextBlock?  _monthLbl;
    private Grid?       _calGrid;
    private TextBlock?  _dateLbl;
    private TextBlock?  _agentLbl;
    private StackPanel? _entryList;
    private TextBox?    _titleBox;
    private TextBox?    _contentBox;
    private TextBox?    _startBox;
    private TextBox?    _endBox;

    private Button? _btnAgents;
    private Button? _btnContracts;
    private readonly Dictionary<string, Button> _catBtns  = new();
    private readonly Dictionary<string, Button> _siteBtns = new();

    // ── 분류별 색상 + 아이콘 ─────────────────────────────────────────────────
    private static readonly Dictionary<string, (string Bg, string Fg, string Bd, string Icon)> CatStyle = new()
    {
        ["출장"]    = ("#1a2a4a", "#88aaee", "#336699", "✈"),
        ["휴일근무"] = ("#3a2a1a", "#ddaa66", "#996633", "🏢"),
        ["연차"]    = ("#1a3a1a", "#88cc88", "#336633", "🌴"),
        ["반차"]    = ("#1a4a2a", "#66bb66", "#2a6633", "🌿"),
        ["공가"]    = ("#3a1a4a", "#cc88ee", "#663399", "🎖"),
        ["기타"]    = ("#2a2a2a", "#aaaaaa", "#555555", "📌"),
    };

    // ── 사이트별 색상 ────────────────────────────────────────────────────────
    private static readonly Dictionary<string, (string Bg, string Fg, string Bd)> SiteStyle = new()
    {
        ["여수"] = ("#1e3a5a", "#88aacc", "#336699"),
        ["율촌"] = ("#1a3a1a", "#aaccaa", "#336633"),
        ["세풍"] = ("#3a2a1a", "#ccaa88", "#996633"),
    };

    // ── 기본 분류 색상 (없는 분류 폴백) ─────────────────────────────────────
    private static (string Bg, string Fg, string Bd, string Icon) GetCat(string 분류)
        => CatStyle.TryGetValue(분류, out var c) ? c : ("#2a2a2a", "#aaaaaa", "#555555", "📌");

    // ────────────────────────────────────────────────────────────────────────
    public SchedulePage()
    {
        TreeControl     = BuildTree();
        CalendarControl = BuildCalendar();
        FormControl     = BuildForm();
        LoadData();
    }

    // =========================================================================
    // Show1 — 직원 / 계약업체 트리
    // =========================================================================
    private Control BuildTree()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };

        // 헤더: 제목 + 탭 버튼
        var hdr = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto"),
            Margin = new Thickness(0, 0, 0, 6),
        };
        hdr.Children.Add(new TextBlock
        {
            Text = "구성원", FontSize = 13, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        });

        _btnAgents = MakeTabBtn("👤 직원", active: true);
        _btnAgents.Margin = new Thickness(0, 0, 2, 0);
        _btnAgents.Click += (_, _) => { _showCon = false; ApplyTabStyles(); LoadTree(); };
        Grid.SetColumn(_btnAgents, 1);

        _btnContracts = MakeTabBtn("🏢 계약업체", active: false);
        _btnContracts.Click += (_, _) => { _showCon = true; ApplyTabStyles(); LoadTree(); };
        Grid.SetColumn(_btnContracts, 2);

        hdr.Children.Add(_btnAgents);
        hdr.Children.Add(_btnContracts);
        Grid.SetRow(hdr, 0);
        root.Children.Add(hdr);

        // 구분선
        var sep = new Border { Height = 1, Background = AppRes("InputBorder"), Margin = new Thickness(0, 0, 0, 4) };
        Grid.SetRow(sep, 1);
        root.Children.Add(sep);

        // 검색창
        _searchBox = new TextBox
        {
            Watermark = "🔍 이름 / 업체명",
            FontFamily = Font, Height = 28, Margin = new Thickness(0, 0, 0, 4),
        };
        _searchBox.TextChanged += (_, _) => LoadTree();
        Grid.SetRow(_searchBox, 2);
        root.Children.Add(_searchBox);

        // 트리뷰
        _tree = new TreeView();
        _tree.SelectionChanged += OnTreeSelection;
        Grid.SetRow(_tree, 3);
        root.Children.Add(_tree);

        return root;
    }

    private static Button MakeTabBtn(string text, bool active) => new()
    {
        Content = text, FontFamily = Font, FontSize = 10,
        Padding = new Thickness(8, 3), CornerRadius = new CornerRadius(10),
        BorderThickness = new Thickness(1),
        Background  = active ? Brush.Parse("#1a2a4a") : Brush.Parse("#222222"),
        Foreground  = active ? Brush.Parse("#88aacc") : Brush.Parse("#666666"),
        BorderBrush = active ? Brush.Parse("#336699") : Brush.Parse("#444444"),
    };

    private void ApplyTabStyles()
    {
        SetTab(_btnAgents,   !_showCon);
        SetTab(_btnContracts, _showCon);
    }

    private static void SetTab(Button? btn, bool active)
    {
        if (btn == null) return;
        btn.Background  = active ? Brush.Parse("#1a2a4a") : Brush.Parse("#222222");
        btn.Foreground  = active ? Brush.Parse("#88aacc") : Brush.Parse("#666666");
        btn.BorderBrush = active ? Brush.Parse("#336699") : Brush.Parse("#444444");
    }

    public void LoadTree()
    {
        if (_tree == null) return;
        _tree.Items.Clear();
        var q = _searchBox?.Text?.Trim() ?? "";

        if (!_showCon)
        {
            // 직원 트리
            var agents = AgentService.GetAllItems()
                .Where(a => string.IsNullOrEmpty(q) || a.성명.Contains(q))
                .OrderBy(a => a.사번).ToList();

            foreach (var a in agents)
            {
                var (bg, fg) = BadgeColorHelper.GetBadgeColor(a.직급);
                var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };
                sp.Children.Add(new TextBlock
                {
                    Text = "👤", VerticalAlignment = VerticalAlignment.Center,
                }.BindSM());
                if (!string.IsNullOrWhiteSpace(a.직급))
                    sp.Children.Add(new Border
                    {
                        Background   = new SolidColorBrush(Color.Parse(bg)),
                        CornerRadius = new CornerRadius(3),
                        Padding      = new Thickness(4, 1),
                        VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = a.직급, FontSize = 10, FontFamily = Font,
                            Foreground = new SolidColorBrush(Color.Parse(fg)),
                        },
                    });
                sp.Children.Add(new TextBlock
                {
                    Text = a.성명, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                }.BindMD());
                _tree.Items.Add(new TreeViewItem { Tag = new TreeTag("agent", a.성명, a.사번), Header = sp });
            }
        }
        else
        {
            // 계약업체 트리
            var contracts = ContractService.GetAllContracts()
                .Where(c => string.IsNullOrEmpty(q) || c.C_CompanyName.Contains(q))
                .OrderBy(c => c.C_CompanyName).ToList();

            foreach (var c in contracts)
            {
                var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };
                sp.Children.Add(new TextBlock
                {
                    Text = "🏢", VerticalAlignment = VerticalAlignment.Center,
                }.BindSM());
                sp.Children.Add(new TextBlock
                {
                    Text = c.C_CompanyName, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                }.BindMD());
                if (!string.IsNullOrWhiteSpace(c.C_Abbreviation))
                    sp.Children.Add(new TextBlock
                    {
                        Text = c.C_Abbreviation, FontFamily = Font,
                        Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindXS());
                _tree.Items.Add(new TreeViewItem { Tag = new TreeTag("contract", c.C_CompanyName, ""), Header = sp });
            }
        }
    }

    private void OnTreeSelection(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0 || e.AddedItems[0] is not TreeViewItem tvi) return;
        if (tvi.Tag is not TreeTag tn) return;

        _agtName = tn.Name;
        _agtId   = tn.Id;
        if (_agentLbl != null)
        {
            _agentLbl.Text       = $"👤 {tn.Name}";
            _agentLbl.Foreground = AppRes("AppFg");
        }
        _tree!.SelectedItem = null;
    }

    // =========================================================================
    // Show2 — 월별 캘린더
    // =========================================================================
    private Control BuildCalendar()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,1,*") };

        // ── 네비게이션 바 ──
        var nav = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto"),
            Margin = new Thickness(0, 0, 0, 6),
        };

        var btnPrev = NavBtn("◀");
        btnPrev.Click += (_, _) => { _month = _month.AddMonths(-1); RefreshCalendar(); };
        Grid.SetColumn(btnPrev, 0);

        _monthLbl = new TextBlock
        {
            FontSize = 14, FontWeight = FontWeight.SemiBold, FontFamily = Font,
            Foreground = AppRes("AppFg"),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment   = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        };

        var btnNext = NavBtn("▶");
        btnNext.Margin = new Thickness(4, 0, 4, 0);
        btnNext.Click += (_, _) => { _month = _month.AddMonths(1); RefreshCalendar(); };
        Grid.SetColumn(btnNext, 2);

        var btnToday = new Button
        {
            Content = "오늘", Height = 28, Padding = new Thickness(10, 0),
            FontSize = 10, FontFamily = Font,
            Background = Brush.Parse("#1a2a4a"), Foreground = Brush.Parse("#88aacc"),
            BorderThickness = new Thickness(1), BorderBrush = Brush.Parse("#336699"),
            CornerRadius = new CornerRadius(4),
            [Grid.ColumnProperty] = 3,
        };
        btnToday.Click += (_, _) =>
        {
            _month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            RefreshCalendar();
        };

        nav.Children.Add(btnPrev); nav.Children.Add(_monthLbl);
        nav.Children.Add(btnNext); nav.Children.Add(btnToday);
        Grid.SetRow(nav, 0);
        root.Children.Add(nav);

        // 구분선
        var sep = new Border { Background = AppRes("InputBorder") };
        Grid.SetRow(sep, 1);
        root.Children.Add(sep);

        // ── 캘린더 본체 ──
        var calOuter = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,*"),
            Margin = new Thickness(0, 6, 0, 0),
        };

        // 요일 헤더
        var dowGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*,*,*,*,*,*") };
        var dowNames  = new[] { "월", "화", "수", "목", "금", "토", "일" };
        var dowColors = new[] { "#999", "#999", "#999", "#999", "#999", "#88aaee", "#ee8888" };
        for (int i = 0; i < 7; i++)
        {
            var tb = new TextBlock
            {
                Text = dowNames[i], FontSize = 11, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = Brush.Parse(dowColors[i]),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 4),
                [Grid.ColumnProperty] = i,
            };
            dowGrid.Children.Add(tb);
        }
        Grid.SetRow(dowGrid, 0);
        calOuter.Children.Add(dowGrid);

        // 날짜 셀 그리드 (6행 × 7열)
        _calGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,*,*,*,*,*,*"),
            RowDefinitions    = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,Auto"),
        };
        Grid.SetRow(_calGrid, 1);
        calOuter.Children.Add(_calGrid);

        var sv = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = calOuter };
        Grid.SetRow(sv, 2);
        root.Children.Add(sv);

        return root;
    }

    private static Button NavBtn(string txt) => new()
    {
        Content = txt, Width = 30, Height = 28, Padding = new Thickness(0),
        Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
        Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")),
        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        HorizontalContentAlignment = HorizontalAlignment.Center,
    };

    public void RefreshCalendar()
    {
        if (_monthLbl == null || _calGrid == null) return;
        _monthLbl.Text = $"{_month.Year}년  {_month.Month}월";
        _calGrid.Children.Clear();

        var entries  = ScheduleService.GetByMonth(_month.Year, _month.Month);
        int days     = DateTime.DaysInMonth(_month.Year, _month.Month);
        int startDow = ((int)_month.DayOfWeek + 6) % 7;  // 월요일=0

        for (int row = 0; row < 6; row++)
        {
            for (int col = 0; col < 7; col++)
            {
                int day  = row * 7 + col - startDow + 1;
                var cell = BuildCell(day, col, days, entries);
                Grid.SetColumn(cell, col);
                Grid.SetRow(cell, row);
                _calGrid.Children.Add(cell);
            }
        }
    }

    private Border BuildCell(int day, int col, int daysInMonth, List<ScheduleEntry> entries)
    {
        bool valid    = day >= 1 && day <= daysInMonth;
        bool isToday  = valid
            && _month.Year  == DateTime.Today.Year
            && _month.Month == DateTime.Today.Month
            && day == DateTime.Today.Day;
        string dateStr = valid
            ? $"{_month.Year:D4}-{_month.Month:D2}-{day:D2}"
            : "";
        bool selected = valid && dateStr == _selDate;

        var cell = new Border
        {
            Background = selected ? Brush.Parse("#182840")
                       : isToday  ? Brush.Parse("#0d1f0d")
                       : AppRes("PanelInnerBg"),
            BorderBrush = selected ? Brush.Parse("#4477cc")
                        : isToday  ? Brush.Parse("#226622")
                        : AppRes("InputBorder"),
            BorderThickness = new Thickness(selected ? 1.5 : 0.5),
            Margin    = new Thickness(1),
            MinHeight = 72,
            Cursor    = valid ? new Cursor(StandardCursorType.Hand) : Cursor.Default,
        };

        var sp = new StackPanel { Margin = new Thickness(3, 2), Spacing = 2 };

        if (valid)
        {
            var numFg = isToday  ? Brush.Parse("#66ee66")
                      : col == 5 ? Brush.Parse("#88aaee")  // 토
                      : col == 6 ? Brush.Parse("#ee8888")  // 일
                      : AppRes("AppFg");

            sp.Children.Add(new TextBlock
            {
                Text = day.ToString(), FontSize = 11, FontFamily = Font,
                FontWeight = isToday ? FontWeight.Bold : FontWeight.Normal,
                Foreground = numFg,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 2, 2),
            });

            // 일정 칩 (최대 3개)
            var dayEntries = entries.Where(e => e.날짜 == dateStr).ToList();
            foreach (var en in dayEntries.Take(3))
            {
                var cs = GetCat(en.분류);
                sp.Children.Add(new Border
                {
                    Background   = Brush.Parse(cs.Bg),
                    CornerRadius = new CornerRadius(3),
                    Padding      = new Thickness(3, 1),
                    Margin       = new Thickness(0, 0, 0, 1),
                    Child = new TextBlock
                    {
                        Text = $"{cs.Icon} {(string.IsNullOrEmpty(en.직원명) ? en.제목 : en.직원명)}",
                        FontSize = 9, FontFamily = Font,
                        Foreground = Brush.Parse(cs.Fg),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    },
                });
            }
            if (dayEntries.Count > 3)
                sp.Children.Add(new TextBlock
                {
                    Text = $"+{dayEntries.Count - 3}개 더",
                    FontSize = 9, FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                });

            // 클릭 → 날짜 선택
            var ds = dateStr;
            cell.PointerPressed += (_, _) =>
            {
                _selDate = ds;
                if (_dateLbl != null)
                {
                    DateTime.TryParse(ds, out var dt);
                    _dateLbl.Text       = $"📅 {dt:yyyy-MM-dd}  ({DowKr(dt)})";
                    _dateLbl.Foreground = AppRes("AppFg");
                }
                RefreshCalendar();
                RefreshEntries();
            };
        }

        cell.Child = sp;
        return cell;
    }

    // =========================================================================
    // Show3 — 작성 폼
    // =========================================================================
    private Control BuildForm()
    {
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(4) };

        root.Children.Add(new TextBlock
        {
            Text = "📋 일정 등록",
            FontSize = 13, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
        });
        root.Children.Add(new Border { Height = 1, Background = AppRes("InputBorder"), Margin = new Thickness(0, 0, 0, 4) });

        // ── 분류 토글 버튼 ──
        root.Children.Add(FormLbl("분류"));
        var catRow = new WrapPanel { Orientation = Orientation.Horizontal };
        foreach (var kv in CatStyle)
        {
            var btn = new Button
            {
                Content = $"{kv.Value.Icon} {kv.Key}",
                FontFamily = Font, FontSize = 10,
                Padding = new Thickness(10, 4), Margin = new Thickness(0, 0, 4, 4),
                CornerRadius = new CornerRadius(12), BorderThickness = new Thickness(1),
            };
            var k = kv.Key;
            btn.Click += (_, _) => { _selCat = k; ApplyCatStyles(); };
            _catBtns[kv.Key] = btn;
            catRow.Children.Add(btn);
        }
        ApplyCatStyles();
        root.Children.Add(catRow);

        // ── 사이트 토글 버튼 ──
        root.Children.Add(FormLbl("사이트"));
        var siteRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        foreach (var kv in SiteStyle)
        {
            var btn = new Button
            {
                Content = kv.Key,
                FontFamily = Font, FontSize = 10,
                Padding = new Thickness(10, 4),
                CornerRadius = new CornerRadius(10), BorderThickness = new Thickness(1),
            };
            var k = kv.Key;
            btn.Click += (_, _) => { _selSite = k; ApplySiteStyles(); };
            _siteBtns[kv.Key] = btn;
            siteRow.Children.Add(btn);
        }
        ApplySiteStyles();
        root.Children.Add(siteRow);

        // ── 날짜 (캘린더에서 선택) ──
        root.Children.Add(FormLbl("날짜"));
        _dateLbl = new TextBlock
        {
            Text = "← 캘린더에서 날짜를 선택하세요",
            FontSize = 11, FontFamily = Font, Foreground = AppRes("FgMuted"),
        };
        root.Children.Add(_dateLbl);

        // ── 구성원 (트리에서 선택) ──
        root.Children.Add(FormLbl("구성원  (왼쪽 트리에서 선택)"));
        _agentLbl = new TextBlock
        {
            Text = "👤 선택 안됨",
            FontSize = 11, FontFamily = Font, Foreground = AppRes("FgMuted"),
        };
        root.Children.Add(_agentLbl);

        // ── 시간 ──
        root.Children.Add(FormLbl("시간"));
        var timeRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        _startBox = new TextBox { Watermark = "09:00", Width = 70, Height = 28, FontFamily = Font };
        _endBox   = new TextBox { Watermark = "18:00", Width = 70, Height = 28, FontFamily = Font };
        timeRow.Children.Add(_startBox);
        timeRow.Children.Add(new TextBlock
        {
            Text = "~", FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        timeRow.Children.Add(_endBox);
        root.Children.Add(timeRow);

        // ── 제목 ──
        root.Children.Add(FormLbl("제목"));
        _titleBox = new TextBox { Watermark = "일정 제목", Height = 30, FontFamily = Font };
        root.Children.Add(_titleBox);

        // ── 내용 ──
        root.Children.Add(FormLbl("내용 (선택)"));
        _contentBox = new TextBox
        {
            Watermark    = "상세 내용",
            AcceptsReturn = true,
            TextWrapping  = TextWrapping.Wrap,
            MinHeight     = 60,
            FontFamily    = Font,
        };
        root.Children.Add(_contentBox);

        // ── 버튼 ──
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 4, 0, 0) };
        var btnSave = new Button
        {
            Content = "💾 저장", Height = 30, Padding = new Thickness(16, 0),
            FontFamily = Font, FontSize = 11,
            Background = Brush.Parse("#1a3a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnSave.Click += (_, _) => SaveEntry();
        var btnClear = new Button
        {
            Content = "↺ 초기화", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = Font, FontSize = 11,
            Background = AppRes("SubBtnBg"), Foreground = AppRes("FgMuted"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnClear.Click += (_, _) => ClearForm();
        btnRow.Children.Add(btnSave);
        btnRow.Children.Add(btnClear);
        root.Children.Add(btnRow);

        // ── 구분선 + 등록된 일정 ──
        root.Children.Add(new Border { Height = 1, Background = AppRes("InputBorder"), Margin = new Thickness(0, 8) });
        root.Children.Add(new TextBlock
        {
            Text = "📌 등록된 일정",
            FontSize = 12, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("FgMuted"),
        });
        _entryList = new StackPanel { Spacing = 4 };
        root.Children.Add(_entryList);

        return root;
    }

    // =========================================================================
    // 스타일 업데이트
    // =========================================================================
    private void ApplyCatStyles()
    {
        foreach (var (k, btn) in _catBtns)
        {
            bool a = k == _selCat;
            var  cs = GetCat(k);
            btn.Background  = a ? Brush.Parse(cs.Bg) : Brush.Parse("#222222");
            btn.Foreground  = a ? Brush.Parse(cs.Fg) : Brush.Parse("#666666");
            btn.BorderBrush = a ? Brush.Parse(cs.Bd) : Brush.Parse("#444444");
        }
    }

    private void ApplySiteStyles()
    {
        foreach (var (k, btn) in _siteBtns)
        {
            bool a = k == _selSite;
            SiteStyle.TryGetValue(k, out var ss);
            btn.Background  = a ? Brush.Parse(ss.Bg) : Brush.Parse("#222222");
            btn.Foreground  = a ? Brush.Parse(ss.Fg) : Brush.Parse("#666666");
            btn.BorderBrush = a ? Brush.Parse(ss.Bd) : Brush.Parse("#444444");
        }
    }

    // =========================================================================
    // CRUD
    // =========================================================================
    public void SaveEntry()
    {
        if (string.IsNullOrEmpty(_selDate)) return;

        ScheduleService.Insert(new ScheduleEntry
        {
            날짜     = _selDate,
            직원명   = _agtName,
            직원id   = _agtId,
            분류     = _selCat,
            사이트   = _selSite,
            제목     = _titleBox?.Text?.Trim() ?? "",
            내용     = _contentBox?.Text?.Trim() ?? "",
            시작시간 = _startBox?.Text?.Trim() ?? "",
            종료시간 = _endBox?.Text?.Trim() ?? "",
            첨부파일 = "",
            등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            등록자   = ETA.Views.MainPage.CurrentEmployeeId,
        });
        ClearForm();
        RefreshEntries();
        RefreshCalendar();
    }

    public void ClearForm()
    {
        if (_titleBox   != null) _titleBox.Text   = "";
        if (_contentBox != null) _contentBox.Text = "";
        if (_startBox   != null) _startBox.Text   = "";
        if (_endBox     != null) _endBox.Text     = "";
        _agtName = ""; _agtId = "";
        if (_agentLbl != null)
        {
            _agentLbl.Text       = "👤 선택 안됨";
            _agentLbl.Foreground = AppRes("FgMuted");
        }
    }

    private void RefreshEntries()
    {
        if (_entryList == null) return;
        _entryList.Children.Clear();
        if (string.IsNullOrEmpty(_selDate)) return;

        var list = ScheduleService.GetByDate(_selDate);
        if (list.Count == 0)
        {
            _entryList.Children.Add(new TextBlock
            {
                Text = "등록된 일정 없음",
                FontSize = 10, FontFamily = Font,
                Foreground = AppRes("FgMuted"),
                Margin = new Thickness(4),
            });
            return;
        }

        foreach (var en in list)
        {
            var cs = GetCat(en.분류);
            var row = new Border
            {
                Background      = Brush.Parse(cs.Bg),
                BorderBrush     = Brush.Parse(cs.Bd),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(8, 4),
                Margin          = new Thickness(0, 1),
            };
            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            var info = new StackPanel { Spacing = 2 };
            info.Children.Add(new TextBlock
            {
                Text = $"{cs.Icon} [{en.분류}] {en.제목}",
                FontSize = 11, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse(cs.Fg),
            });
            var sub = new List<string>();
            if (!string.IsNullOrEmpty(en.직원명)) sub.Add($"👤 {en.직원명}");
            if (!string.IsNullOrEmpty(en.사이트)) sub.Add($"🌐 {en.사이트}");
            if (!string.IsNullOrEmpty(en.시작시간)) sub.Add($"⏰ {en.시작시간}~{en.종료시간}");
            if (sub.Count > 0)
                info.Children.Add(new TextBlock
                {
                    Text = string.Join("  ", sub),
                    FontSize = 10, FontFamily = Font, Foreground = AppRes("FgMuted"),
                });
            if (!string.IsNullOrEmpty(en.내용))
                info.Children.Add(new TextBlock
                {
                    Text = en.내용,
                    FontSize = 10, FontFamily = Font, Foreground = AppRes("FgMuted"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
            grid.Children.Add(info);

            var id = en.Id;
            var del = new Button
            {
                Content = "✕", Width = 22, Height = 22, Padding = new Thickness(0),
                Background = Brushes.Transparent, Foreground = Brush.Parse("#883333"),
                BorderThickness = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Top,
                [Grid.ColumnProperty] = 1,
            };
            del.Click += (_, _) => { ScheduleService.Delete(id); RefreshEntries(); RefreshCalendar(); };
            grid.Children.Add(del);
            row.Child = grid;
            _entryList.Children.Add(row);
        }
    }

    // =========================================================================
    // Public API
    // =========================================================================
    public void LoadData()
    {
        LoadTree();
        RefreshCalendar();
        RefreshEntries();
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private static TextBlock FormLbl(string text) => new()
    {
        Text = text, FontSize = 10, FontFamily = Font,
        Foreground = new SolidColorBrush(Color.Parse("#888888")),
        Margin = new Thickness(0, 4, 0, 2),
    };

    private static string DowKr(DateTime d) => d.DayOfWeek switch
    {
        DayOfWeek.Monday    => "월",
        DayOfWeek.Tuesday   => "화",
        DayOfWeek.Wednesday => "수",
        DayOfWeek.Thursday  => "목",
        DayOfWeek.Friday    => "금",
        DayOfWeek.Saturday  => "토",
        _                   => "일",
    };
}
