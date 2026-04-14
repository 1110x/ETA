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
using ETA.Services.Common;
using ETA.Views;

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
        new("avares://ETA/Assets/Fonts#Pretendard");

    // 트리 노드 태그
    private record TreeTag(string Type, string Name, string Id, string Abbr = "", string Date = "");

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public Control TreeControl     { get; }
    public Control CalendarControl { get; }
    public Control FormControl     { get; }

    // ── 상태 ────────────────────────────────────────────────────────────────
    private DateTime _month     = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    private string?  _selDate;
    private string?  _focusDate;   // 의뢰목록 채수담당자 배정 날짜 포커스
    private string   _selCat   = "채수";
    private string   _selSite  = "여수";
    private bool     _showCon  = false;      // false=직원, true=의뢰목록
    private string   _agtName  = "";
    private string   _agtId    = "";
    private string   _selCompanyAbbr = "";
    private int      _editId   = -1;
    private bool     _lastSaved = false;     // 저장 완료 여부 (progress)
    // 채수 복수 의뢰 선택 (Id, 약칭)
    private readonly List<(int Id, string Abbr)> _selReqs = new();
    // 채수/다중 선택용
    private readonly List<(string Name, string Id)> _selectedAgents = new();

    // ── UI 참조 ──────────────────────────────────────────────────────────────
    private TreeView?   _tree;
    private TextBox?    _searchBox;
    private TextBlock?  _monthLbl;
    private Grid?       _calGrid;
    private TextBlock?  _dateLbl;
    private TextBlock?  _agentLbl;
    private WrapPanel?  _agentChipsPanel;
    private Border?     _companyBadge;
    private StackPanel? _entryList;
    private TextBlock?  _entryDateLbl;
    private TextBox?    _contentBox;
    private TextBlock?  _formTitle;
    private Button?     _btnSave;
    private Button?     _btnDelete;

    private Button?     _btnAgents;
    private Button?     _btnContracts;
    private StackPanel? _progressPanel;
    private readonly Dictionary<string, Button> _catBtns  = new();
    private readonly Dictionary<string, Button> _siteBtns = new();
    private bool _lastClickCtrl = false;

    // ── 분류별 색상 + 아이콘 ─────────────────────────────────────────────────
    // 출장: 사용자1 + 계약업체 → 즉시 등록
    // 휴일근무/반차/공가/기타: 사용자1 → 즉시 등록
    private static readonly Dictionary<string, (string Bg, string Fg, string Bd, string Icon)> CatStyle = new()
    {
        ["채수"]    = ("#1a2a4a", "#88aaee", "#336699", "🌊"),
        ["연차"]    = ("#1a3a3a", "#66cccc", "#226666", "🏖"),
        ["휴일근무"] = ("#3a2a1a", "#ddaa66", "#996633", "🏢"),
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
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,*") };

        // 헤더: 제목 + 탭 버튼
        var hdr = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto"),
            Margin = new Thickness(0, 0, 0, 6),
        };
        hdr.Children.Add(new TextBlock
        {
            Text = "구성원", FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        });

        _btnAgents = MakeTabBtn("👤 직원", active: true);
        _btnAgents.Margin = new Thickness(0, 0, 2, 0);
        _btnAgents.Click += (_, _) => { _showCon = false; ApplyTabStyles(); LoadTree(); };
        Grid.SetColumn(_btnAgents, 1);

        _btnContracts = MakeTabBtn("📋 의뢰목록", active: false);
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

        // ── 분류 토글 버튼 (Row 2) ──
        var catRow = new WrapPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 4) };
        foreach (var kv in CatStyle)
        {
            var btn = new Button
            {
                Content = $"{kv.Value.Icon} {kv.Key}",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Padding = new Thickness(8, 3), Margin = new Thickness(0, 0, 3, 3),
                CornerRadius = new CornerRadius(10), BorderThickness = new Thickness(1),
            };
            var k = kv.Key;
            btn.Click += (_, _) =>
            {
                _selCat = k;
                _lastSaved = false;
                ApplyCatStyles();
                RefreshProgress();
                if (_selCat == "채수")
                {
                    // 채수 선택 → 의뢰목록 탭으로 자동 전환
                    _showCon = true;
                    ApplyTabStyles();
                    LoadTree();
                }
            };
            _catBtns[kv.Key] = btn;
            catRow.Children.Add(btn);
        }
        ApplyCatStyles();
        Grid.SetRow(catRow, 2);
        root.Children.Add(catRow);

        // ── 진행 상태 표시 (Row 3) ──
        _progressPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 4,
            Margin = new Thickness(0, 2, 0, 4),
        };
        Grid.SetRow(_progressPanel, 3);
        root.Children.Add(_progressPanel);
        RefreshProgress();

        // 검색창
        _searchBox = new TextBox
        {
            Watermark = "🔍 이름 / 의뢰명",
            FontFamily = Font, Height = 28, Margin = new Thickness(0, 0, 0, 4),
        };
        _searchBox.TextChanged += (_, _) => LoadTree();
        Grid.SetRow(_searchBox, 4);
        root.Children.Add(_searchBox);

        // 트리뷰
        _tree = new TreeView();
        _tree.AddHandler(InputElement.PointerPressedEvent, (object? s, PointerPressedEventArgs ev) =>
        {
            _lastClickCtrl = ev.KeyModifiers.HasFlag(KeyModifiers.Control);
        }, handledEventsToo: true);
        _tree.SelectionChanged += OnTreeSelection;
        Grid.SetRow(_tree, 5);
        root.Children.Add(_tree);

        return root;
    }

    private static Button MakeTabBtn(string text, bool active) => new()
    {
        Content = text, FontFamily = Font, FontSize = AppTheme.FontSM,
        Padding = new Thickness(8, 3), CornerRadius = new CornerRadius(10),
        BorderThickness = new Thickness(1),
        Background  = active ? AppTheme.BgActiveBlue : AppTheme.BgSecondary,
        Foreground  = active ? AppTheme.FgInfo : AppTheme.FgDimmed,
        BorderBrush = active ? AppTheme.BorderInfo : AppTheme.BorderMuted,
    };

    private void ApplyTabStyles()
    {
        SetTab(_btnAgents,   !_showCon);
        SetTab(_btnContracts, _showCon);
    }

    private static void SetTab(Button? btn, bool active)
    {
        if (btn == null) return;
        btn.Background  = active ? AppTheme.BgActiveBlue : AppTheme.BgSecondary;
        btn.Foreground  = active ? AppTheme.FgInfo : AppTheme.FgDimmed;
        btn.BorderBrush = active ? AppTheme.BorderInfo : AppTheme.BorderMuted;
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
                            Text = a.직급, FontSize = AppTheme.FontSM, FontFamily = Font,
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
            // 의뢰목록 — 최근 분석의뢰 목록 (채수담당자 배정용)
            var records = AnalysisRequestService.GetAllRecords()
                .Where(r => string.IsNullOrEmpty(q)
                    || r.약칭.Contains(q) || r.시료명.Contains(q) || r.접수번호.Contains(q))
                .Take(60).ToList();

            foreach (var r in records)
            {
                var sp = new StackPanel { Orientation = Orientation.Vertical, Spacing = 1 };
                bool reqSelected = _selReqs.Any(x => x.Id == r.Id);

                var topRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };
                topRow.Children.Add(new TextBlock
                {
                    Text = "📋", VerticalAlignment = VerticalAlignment.Center,
                }.BindSM());

                // 약칭 배지
                if (!string.IsNullOrWhiteSpace(r.약칭))
                {
                    var (bg, fg) = BadgeColorHelper.GetBadgeColor(r.약칭);
                    topRow.Children.Add(new Border
                    {
                        Background = Brush.Parse(bg), CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(4, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock { Text = r.약칭, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = Brush.Parse(fg) },
                    });
                }

                topRow.Children.Add(new TextBlock
                {
                    Text = r.시료명, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 120,
                }.BindMD());

                // 정도보증 O 마크
                if (r.정도보증.Trim() == "O" || r.정도보증.Trim() == "o")
                    topRow.Children.Add(new Border
                    {
                        Background = Brush.Parse("#2a4a1a"), BorderBrush = AppTheme.BorderActive,
                        BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(3, 0), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock { Text = "O", FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brush.Parse("#88dd44"), FontWeight = FontWeight.Bold },
                    });

                // 채수담당자 배지
                if (!string.IsNullOrWhiteSpace(r.채수담당자))
                {
                    var (sbg, sfg) = BadgeColorHelper.GetBadgeColor(r.채수담당자);
                    topRow.Children.Add(new Border
                    {
                        Background = Brush.Parse(sbg), CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(4, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock { Text = r.채수담당자, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brush.Parse(sfg) },
                    });
                }

                if (reqSelected)
                    topRow.Children.Add(new TextBlock
                    {
                        Text = "✓", FontSize = AppTheme.FontSM, FontFamily = Font,
                        Foreground = AppTheme.FgSuccess, FontWeight = FontWeight.Bold,
                        VerticalAlignment = VerticalAlignment.Center,
                    });

                sp.Children.Add(topRow);

                // 채취일자 + 접수번호
                if (!string.IsNullOrEmpty(r.채취일자) || !string.IsNullOrEmpty(r.접수번호))
                    sp.Children.Add(new TextBlock
                    {
                        Text = $"  {r.채취일자}  {r.접수번호}".Trim(),
                        FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"),
                    });

                string displayAbbr = !string.IsNullOrWhiteSpace(r.약칭) ? r.약칭
                    : !string.IsNullOrWhiteSpace(r.접수번호) ? r.접수번호
                    : r.시료명;
                // Date: 채수담당자 배정된 경우에만 채취일자 저장 (캘린더 포커스용)
                string tagDate = !string.IsNullOrWhiteSpace(r.채수담당자) ? NormalizeDate(r.채취일자) : "";
                var tvi = new TreeViewItem
                {
                    Tag    = new TreeTag("request", displayAbbr, r.Id.ToString(), r.접수번호, tagDate),
                    Header = sp,
                    Background = reqSelected ? AppTheme.BgActiveGreen : null,
                };

                // 채수담당자 배정된 항목 → 우클릭 컨텍스트 메뉴
                if (!string.IsNullOrWhiteSpace(r.채수담당자))
                    tvi.ContextMenu = BuildSamplerContextMenu(r.Id, r.채수담당자);

                _tree.Items.Add(tvi);
            }
        }
    }

    // ── 의뢰목록 우클릭 컨텍스트 메뉴 (채수담당자 삭제/변경) ─────────────
    private ContextMenu BuildSamplerContextMenu(int reqId, string sampler)
    {
        var menu = new ContextMenu();
        var names = sampler.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

        // 개별 인원 삭제
        if (names.Count > 1)
        {
            foreach (var name in names)
            {
                var n = name;
                var mi = new MenuItem
                {
                    Header = $"❌ {n} 삭제",
                    FontFamily = Font, FontSize = AppTheme.FontMD,
                };
                mi.Click += (_, _) =>
                {
                    var remaining = names.Where(x => x != n).ToList();
                    AnalysisRequestService.UpdateSamplers(reqId, remaining);
                    LoadTree(); RefreshCalendar(); RefreshEntries();
                };
                menu.Items.Add(mi);
            }
            menu.Items.Add(new Separator());
        }

        // 전체 삭제
        var clearAll = new MenuItem
        {
            Header = "🗑 담당자 전체 삭제",
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };
        clearAll.Click += (_, _) =>
        {
            AnalysisRequestService.UpdateSamplers(reqId, Array.Empty<string>());
            LoadTree(); RefreshCalendar(); RefreshEntries();
        };
        menu.Items.Add(clearAll);

        // 담당자 변경 서브메뉴
        menu.Items.Add(new Separator());
        var changeMenu = new MenuItem
        {
            Header = "🔄 담당자 변경",
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };
        var agents = AgentService.GetAllItems().OrderBy(a => a.사번).ToList();
        foreach (var a in agents)
        {
            var agentName = a.성명;
            var mi = new MenuItem
            {
                Header = $"  {agentName} ({a.사번})",
                FontFamily = Font, FontSize = AppTheme.FontMD,
            };
            mi.Click += (_, _) =>
            {
                AnalysisRequestService.UpdateSamplers(reqId, new[] { agentName });
                LoadTree(); RefreshCalendar(); RefreshEntries();
            };
            changeMenu.Items.Add(mi);
        }
        menu.Items.Add(changeMenu);

        return menu;
    }

    private bool _suppressSelection;

    private void OnTreeSelection(object? sender, SelectionChangedEventArgs e)
    {
        if (_suppressSelection) return;
        if (e.AddedItems.Count == 0 || e.AddedItems[0] is not TreeViewItem tvi) return;
        if (tvi.Tag is not TreeTag tn) return;

        _suppressSelection = true;
        _tree!.SelectedItem = null;
        _suppressSelection = false;

        if (tn.Type == "request")
        {
            int reqId = int.TryParse(tn.Id, out var rid) ? rid : -1;
            string reqAbbr = tn.Name;
            string rawDate = tn.Date;
            bool hasSampler = !string.IsNullOrWhiteSpace(rawDate);
            if (hasSampler)
            {
                // ── 담당자 배정 완료: 시퀀스 진입 안함, 캘린더 포커스만 ──
                _selReqs.Clear();
                _selReqs.Add((reqId, reqAbbr));
                _selCompanyAbbr = reqAbbr;
                RefreshCompanyBadge();

                var normalized = NormalizeDate(rawDate);
                if (DateTime.TryParse(normalized, out var focusDt))
                {
                    var isoDate = focusDt.ToString("yyyy-MM-dd");
                    _focusDate = isoDate;
                    _selDate   = isoDate;
                    _month     = new DateTime(focusDt.Year, focusDt.Month, 1);
                    if (_dateLbl != null)
                    {
                        _dateLbl.Text       = $"📅 {focusDt:yyyy-MM-dd}  ({DowKr(focusDt)})";
                        _dateLbl.Foreground = AppRes("AppFg");
                    }
                    RefreshCalendar();
                    RefreshEntries();
                }
                return;
            }

            // ── 담당자 미배정: 채수 시퀀스 시작 ──
            _selCat = "채수";
            ApplyCatStyles();
            _focusDate = null;

            if (_lastClickCtrl)
            {
                int idx = _selReqs.FindIndex(r => r.Id == reqId);
                if (idx >= 0) _selReqs.RemoveAt(idx);
                else          _selReqs.Add((reqId, reqAbbr));
                _suppressSelection = true;
                LoadTree();
                _suppressSelection = false;
            }
            else
            {
                _selReqs.Clear();
                _selReqs.Add((reqId, reqAbbr));
                _showCon = false;
                ApplyTabStyles();
                _suppressSelection = true;
                LoadTree();
                _suppressSelection = false;
            }

            _selCompanyAbbr = string.Join(",", _selReqs.Select(r => r.Abbr).Distinct());
            RefreshCompanyBadge();
            RefreshProgress();
            return;
        }

        // 직원 노드
        if (_selCat == "채수")
        {
            // 채수: Ctrl+클릭 → 복수 선택, 일반 클릭 → 단일 선택 (자동저장 없음)
            int idx = _selectedAgents.FindIndex(a => a.Id == tn.Id);
            if (idx >= 0)
            {
                _selectedAgents.RemoveAt(idx);
            }
            else
            {
                if (!_lastClickCtrl)
                    _selectedAgents.Clear();
                _selectedAgents.Add((tn.Name, tn.Id));
            }
            RefreshAgentChips();
            RefreshProgress();
        }
        else
        {
            // 휴가류: 단일 선택 (자동저장 없음 — 완료 버튼으로 저장)
            _selectedAgents.Clear();
            _selectedAgents.Add((tn.Name, tn.Id));
            _agtName = tn.Name;
            _agtId   = tn.Id;
            RefreshAgentChips();
            RefreshProgress();
        }
    }

    private void RefreshAgentChips()
    {
        if (_agentChipsPanel == null) return;
        _agentChipsPanel.Children.Clear();

        if (_selectedAgents.Count == 0)
        {
            if (_agentLbl != null) { _agentLbl.Text = "👤 선택 안됨"; _agentLbl.Foreground = AppRes("FgMuted"); }
            return;
        }

        if (_agentLbl != null) { _agentLbl.Text = ""; }

        foreach (var (name, id) in _selectedAgents)
        {
            var chip = new Border
            {
                Background      = AppTheme.BgActiveBlue,
                BorderBrush     = AppTheme.BorderAccent,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(10),
                Padding         = new Thickness(8, 3),
                Margin          = new Thickness(0, 2, 4, 2),
            };
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            row.Children.Add(new TextBlock
            {
                Text = $"👤 {name}", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
            });
            // X 버튼
            var removeId = id;
            var btnX = new Button
            {
                Content = "×", FontSize = AppTheme.FontBase, Padding = new Thickness(2, 0),
                Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Foreground = new SolidColorBrush(Color.Parse("#886688")),
                VerticalAlignment = VerticalAlignment.Center,
            };
            btnX.Click += (_, _) =>
            {
                _selectedAgents.RemoveAll(a => a.Id == removeId);
                RefreshAgentChips();
            };
            row.Children.Add(btnX);
            chip.Child = row;
            _agentChipsPanel.Children.Add(chip);
        }
    }

    private void RefreshCompanyBadge()
    {
        if (_companyBadge == null) return;
        if (string.IsNullOrEmpty(_selCompanyAbbr))
        {
            _companyBadge.IsVisible = false;
            return;
        }
        _companyBadge.IsVisible = true;
        if (_companyBadge.Child is TextBlock tb)
            tb.Text = _selCat == "채수" ? $"📋 {_selCompanyAbbr}" : $"🏢 {_selCompanyAbbr}";
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
            FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold, FontFamily = Font,
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
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Background = AppTheme.BgActiveBlue, Foreground = AppTheme.FgInfo,
            BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderInfo,
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
        var dowNames  = new[] { "일", "월", "화", "수", "목", "금", "토" };
        var dowColors = new[] { "#ee8888", "#999", "#999", "#999", "#999", "#999", "#88aaee" };
        for (int i = 0; i < 7; i++)
        {
            var tb = new TextBlock
            {
                Text = dowNames[i], FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
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
        Background = AppTheme.BorderSubtle,
        Foreground = AppTheme.FgMuted,
        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        HorizontalContentAlignment = HorizontalAlignment.Center,
    };

    // ── 연속일 스팬 위치 ─────────────────────────────────────────────────
    private enum SpanPos { Single, Start, Middle, End }

    /// <summary>동일 직원+분류+업체약칭이 연속일에 걸쳐 있으면 스팬으로 묶는다.</summary>
    private static Dictionary<(string date, string key), SpanPos> ComputeSpans(List<ScheduleEntry> entries)
    {
        var result = new Dictionary<(string, string), SpanPos>();
        // 그룹키 = 직원id|분류|업체약칭
        var groups = entries
            .GroupBy(e => $"{e.직원id}|{e.분류}|{e.업체약칭}");

        foreach (var g in groups)
        {
            var dates = g.Select(e => e.날짜).Distinct()
                .Where(d => DateTime.TryParse(d, out _))
                .OrderBy(d => d).ToList();
            if (dates.Count < 2) { foreach (var d in dates) result[(d, g.Key)] = SpanPos.Single; continue; }

            // 연속 구간 찾기
            var runs = new List<List<string>> { new() { dates[0] } };
            for (int i = 1; i < dates.Count; i++)
            {
                var prev = DateTime.Parse(dates[i - 1]);
                var curr = DateTime.Parse(dates[i]);
                if ((curr - prev).TotalDays == 1)
                    runs[^1].Add(dates[i]);
                else
                    runs.Add(new List<string> { dates[i] });
            }

            foreach (var run in runs)
            {
                if (run.Count == 1) { result[(run[0], g.Key)] = SpanPos.Single; continue; }
                for (int i = 0; i < run.Count; i++)
                {
                    var pos = i == 0 ? SpanPos.Start : i == run.Count - 1 ? SpanPos.End : SpanPos.Middle;
                    result[(run[i], g.Key)] = pos;
                }
            }
        }
        return result;
    }

    /// <summary>분석의뢰 채수담당자 배정 → ScheduleEntry로 변환하여 캘린더에 표시</summary>
    private static List<ScheduleEntry> GetRequestEntriesForMonth(int year, int month)
    {
        var prefix = $"{year:D4}-{month:D2}";
        var list = new List<ScheduleEntry>();
        try
        {
            var records = AnalysisRequestService.GetAllRecords()
                .Where(r => !string.IsNullOrWhiteSpace(r.채수담당자))
                .ToList();

            foreach (var r in records)
            {
                var normalized = NormalizeDate(r.채취일자);
                if (!normalized.StartsWith(prefix)) continue;

                // 채수담당자가 쉼표로 구분된 복수 인원일 수 있음
                var names = r.채수담당자.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var name in names)
                {
                    list.Add(new ScheduleEntry
                    {
                        날짜     = normalized,
                        직원명   = name,
                        직원id   = name,  // id 없음 → 이름으로 대체
                        분류     = "채수",
                        업체약칭 = !string.IsNullOrWhiteSpace(r.약칭) ? r.약칭 : r.시료명,
                    });
                }
            }
        }
        catch { }
        return list;
    }

    public void RefreshCalendar()
    {
        if (_monthLbl == null || _calGrid == null) return;
        _monthLbl.Text = $"{_month.Year}년  {_month.Month}월";
        _calGrid.Children.Clear();

        // 일정 테이블 + 분석의뢰 채수담당자 배정 병합
        var entries = ScheduleService.GetByMonth(_month.Year, _month.Month);
        var reqEntries = GetRequestEntriesForMonth(_month.Year, _month.Month);
        // 중복 방지: 이미 일정에 같은 날짜+직원+약칭이 있으면 의뢰 엔트리 제외
        foreach (var re in reqEntries)
        {
            bool dup = entries.Any(e =>
                e.날짜 == re.날짜 && e.직원명 == re.직원명 && e.업체약칭 == re.업체약칭);
            if (!dup) entries.Add(re);
        }
        var spans       = ComputeSpans(entries);
        int totalAgents = AgentService.GetAllItems().Count;
        int days        = DateTime.DaysInMonth(_month.Year, _month.Month);
        int startDow    = (int)_month.DayOfWeek;  // 일요일=0

        for (int row = 0; row < 6; row++)
        {
            for (int col = 0; col < 7; col++)
            {
                int day  = row * 7 + col - startDow + 1;
                var cell = BuildCell(day, col, days, entries, totalAgents, spans);
                Grid.SetColumn(cell, col);
                Grid.SetRow(cell, row);
                _calGrid.Children.Add(cell);
            }
        }
    }

    private Border BuildCell(int day, int col, int daysInMonth, List<ScheduleEntry> entries, int totalAgents, Dictionary<(string date, string key), SpanPos> spans)
    {
        bool valid    = day >= 1 && day <= daysInMonth;
        bool isToday  = valid
            && _month.Year  == DateTime.Today.Year
            && _month.Month == DateTime.Today.Month
            && day == DateTime.Today.Day;
        string dateStr = valid
            ? $"{_month.Year:D4}-{_month.Month:D2}-{day:D2}"
            : "";
        bool selected  = valid && dateStr == _selDate;
        bool isFocused = valid && dateStr == _focusDate;

        var cell = new Border
        {
            Background = isFocused ? Brush.Parse("#2a1800")
                       : selected  ? Brush.Parse("#182840")
                       : isToday   ? Brush.Parse("#0d1f0d")
                       : AppRes("PanelInnerBg"),
            BorderBrush = isFocused ? Brush.Parse("#dd8800")
                        : selected  ? Brush.Parse("#4477cc")
                        : isToday   ? Brush.Parse("#226622")
                        : AppRes("InputBorder"),
            BorderThickness = new Thickness(isFocused || selected ? 2 : 0.5),
            Margin    = new Thickness(0),  // ← 제거 (grid spacing 불필요)
            Padding   = new Thickness(4, 2),  // ← 1(margin) + 3(기존 sp margin) = 4로 통합
            MinHeight = 94,
            Cursor    = valid ? new Cursor(StandardCursorType.Hand) : Cursor.Default,
        };

        var sp = new StackPanel { Margin = new Thickness(0), Spacing = 1 };  // Margin 0

        if (valid)
        {
            // col 0=일, 6=토
            var numFg = isToday  ? Brush.Parse("#66ee66")
                      : col == 0 ? AppTheme.FgDanger  // 일
                      : col == 6 ? AppTheme.FgInfo  // 토
                      : AppRes("AppFg");

            var dayEntries  = entries.Where(e => e.날짜 == dateStr).ToList();
            var tripEntries = dayEntries.Where(e => e.분류 == "채수").ToList();
            var leaveEntries= dayEntries.Where(e => e.분류 != "채수").ToList();

            // 가용인력 계산 (연차/반차/공가/채수 제외)
            int unavail  = dayEntries
                .Where(e => e.분류 is "연차" or "반차" or "공가" or "채수")
                .Select(e => e.직원id)
                .Distinct().Count();
            int avail = Math.Max(0, totalAgents - unavail);

            // ── 헤더: 가용인력 + 날짜 번호 ──
            var hdrRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            hdrRow.Children.Add(new TextBlock
            {
                Text = $"가용:{avail}명",
                FontFamily = Font,
                Foreground = totalAgents > 0 && avail < totalAgents * 0.6 ? AppTheme.FgDanger
                           : totalAgents > 0 && avail < totalAgents * 0.8 ? Brush.Parse("#eedd66")
                           : AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
            }.BindXS());
            hdrRow.Children.Add(new TextBlock
            {
                Text = day.ToString(), FontFamily = Font,
                FontWeight = isToday ? FontWeight.Bold : FontWeight.Normal,
                Foreground = numFg,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 2, 0),
                [Grid.ColumnProperty] = 1,
            }.BindSM());
            sp.Children.Add(hdrRow);

            var all = BuildEntryColumn(dayEntries, 4, dateStr, spans);
            sp.Children.Add(all);

            // 클릭 → 날짜 선택
            var ds = dateStr;
            cell.PointerPressed += (_, _) =>
            {
                _selDate   = ds;
                _focusDate = null;   // 직접 클릭하면 포커스 해제
                _lastSaved = false;
                if (_dateLbl != null)
                {
                    DateTime.TryParse(ds, out var dt);
                    _dateLbl.Text       = $"📅 {dt:yyyy-MM-dd}  ({DowKr(dt)})";
                    _dateLbl.Foreground = AppRes("AppFg");
                }
                RefreshProgress();
                RefreshCalendar();
                RefreshEntries();
            };
        }

        cell.Child = sp;
        return cell;
    }

    private StackPanel BuildEntryColumn(List<ScheduleEntry> list, int maxShow,
        string dateStr, Dictionary<(string date, string key), SpanPos> spans)
    {
        var col = new StackPanel { Spacing = 1 };

        // 같은 분류+업체약칭끼리 그룹핑 → 인원 합치기
        var grouped = list
            .GroupBy(e => $"{e.분류}|{e.업체약칭}")
            .Select(g => (
                First: g.First(),
                Names: string.Join(", ", g
                    .Select(e => !string.IsNullOrEmpty(e.직원명) ? e.직원명 : e.제목)
                    .Where(n => !string.IsNullOrEmpty(n))
                    .Distinct()),
                // 스팬은 그룹 내 첫 엔트리 기준 (같은 분류+약칭이면 스팬 동일)
                SpanKey: $"{g.First().직원id}|{g.First().분류}|{g.First().업체약칭}"
            ))
            .ToList();

        foreach (var (en, names, spanKey) in grouped.Take(maxShow))
        {
            var cs = GetCat(en.분류);

            // 다중 직원인 경우 스팬 키를 별도로 계산 (그룹 내 첫 번째만으로는 부족)
            // → 그룹 내 모든 직원 스팬 중 가장 유의미한 것 사용
            var groupEntries = list.Where(e => $"{e.분류}|{e.업체약칭}" == $"{en.분류}|{en.업체약칭}").ToList();
            var pos = SpanPos.Single;
            foreach (var ge in groupEntries)
            {
                var k = $"{ge.직원id}|{ge.분류}|{ge.업체약칭}";
                if (spans.TryGetValue((dateStr, k), out var p) && p != SpanPos.Single)
                { pos = p; break; }
            }

            string label = en.분류 == "채수" && !string.IsNullOrEmpty(en.업체약칭)
                ? en.업체약칭 : en.분류;

            // 스팬 위치별 텍스트
            string chipText;
            if (pos == SpanPos.Middle)
                chipText = "━━";
            else if (pos == SpanPos.End)
                chipText = string.IsNullOrEmpty(names) ? $"▸{label}" : $"▸{names}";
            else
                chipText = string.IsNullOrEmpty(names) ? $"{cs.Icon} {label}" : $"{cs.Icon} {label} {names}";

            var radius = pos switch
            {
                SpanPos.Start  => new CornerRadius(3, 0, 0, 3),
                SpanPos.Middle => new CornerRadius(0),
                SpanPos.End    => new CornerRadius(0, 3, 3, 0),
                _              => new CornerRadius(3),
            };
            var margin = pos switch
            {
                SpanPos.Start  => new Thickness(0, 0, -4, 0),  // Cell padding(4) 극복
                SpanPos.Middle => new Thickness(-4, 0, -4, 0),
                SpanPos.End    => new Thickness(-4, 0, 0, 0),
                _              => new Thickness(0),
            };

            col.Children.Add(new Border
            {
                Background      = Brush.Parse(cs.Bg),
                BorderBrush     = Brush.Parse(cs.Bd),
                BorderThickness = new Thickness(
                    pos is SpanPos.Middle or SpanPos.End ? 0 : 1,
                    1,
                    pos is SpanPos.Start or SpanPos.Middle ? 0 : 1,
                    1),
                CornerRadius    = radius,
                Padding         = new Thickness(3, 0),
                Margin          = margin,
                Child = new TextBlock
                {
                    Text = chipText, FontFamily = Font,
                    Foreground = Brush.Parse(cs.Fg),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }.BindXS(),
            });
        }
        if (grouped.Count > maxShow)
            col.Children.Add(new TextBlock
            {
                Text = $"+{grouped.Count - maxShow}",
                FontFamily = Font, Foreground = AppRes("FgMuted"),
            }.BindXS());
        return col;
    }

    // =========================================================================
    // Show3 — 작성 / 수정 폼
    // =========================================================================
    private Control BuildForm()
    {
        var outer = new Grid { RowDefinitions = new RowDefinitions("*,Auto") };

        // ═══════════════════════════════════════════════════════════════
        // Row 0: 등록된 일정 목록 (자체 ScrollViewer로 스크롤)
        // ═══════════════════════════════════════════════════════════════
        var entrySection = new StackPanel { Spacing = 4, Margin = new Thickness(4) };

        var entryHdrRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(0, 0, 0, 4) };
        entryHdrRow.Children.Add(new TextBlock
        {
            Text = "📌 등록된 일정",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        _entryDateLbl = new TextBlock
        {
            FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        };
        entryHdrRow.Children.Add(_entryDateLbl);
        entrySection.Children.Add(entryHdrRow);

        _entryList = new StackPanel { Spacing = 4 };
        entrySection.Children.Add(_entryList);

        var entrySv = new ScrollViewer
        {
            Content = entrySection,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        };
        Grid.SetRow(entrySv, 0);
        outer.Children.Add(entrySv);

        // ═══════════════════════════════════════════════════════════════
        // Row 1: 작성/수정 폼 (하단 고정)
        // ═══════════════════════════════════════════════════════════════
        var formPanel = new Border
        {
            BorderThickness = new Thickness(0, 1, 0, 0),
            BorderBrush = AppRes("InputBorder"),
            Padding = new Thickness(4, 6, 4, 4),
        };

        var twoCol = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("160,1,*"),
        };

        // ── 왼쪽: 날짜 / 업체 / 구성원 ──
        var left = new StackPanel { Spacing = 6, Margin = new Thickness(0, 0, 8, 0) };

        _dateLbl = new TextBlock
        {
            Text = "← 캘린더에서 날짜를 선택하세요",
            FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("FgMuted"),
            TextWrapping = TextWrapping.Wrap,
        };
        left.Children.Add(_dateLbl);

        _companyBadge = new Border
        {
            Background      = AppTheme.BgActiveGreen,
            BorderBrush     = AppTheme.BorderActive,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 3),
            IsVisible       = false,
            Child = new TextBlock { FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = Brush.Parse("#88ccaa") },
        };
        left.Children.Add(_companyBadge);

        left.Children.Add(FormLbl("구성원  (채수: Ctrl+클릭으로 복수 선택)"));
        _agentLbl = new TextBlock
        {
            Text = "👤 선택 안됨",
            FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppRes("FgMuted"),
        };
        _agentChipsPanel = new WrapPanel { Orientation = Orientation.Horizontal };
        left.Children.Add(_agentLbl);
        left.Children.Add(_agentChipsPanel);

        Grid.SetColumn(left, 0);
        twoCol.Children.Add(left);

        // 구분선
        var div = new Border { Width = 1, Background = AppRes("InputBorder") };
        Grid.SetColumn(div, 1);
        twoCol.Children.Add(div);

        // ── 오른쪽: 내용 + 버튼 ──
        var right = new StackPanel { Spacing = 6, Margin = new Thickness(8, 0, 0, 0) };

        _contentBox = new TextBox
        {
            Watermark     = "상세 내용",
            AcceptsReturn = true,
            TextWrapping  = TextWrapping.Wrap,
            MinHeight     = 60,
            FontFamily    = Font,
        };
        right.Children.Add(_contentBox);

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 4, 0, 0) };
        _btnSave = new Button
        {
            Content = "💾 저장", Height = 30, Padding = new Thickness(16, 0),
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        _btnSave.Click += (_, _) => SaveEntry();

        var btnClear = new Button
        {
            Content = "↺ 초기화", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Background = AppRes("SubBtnBg"), Foreground = AppRes("FgMuted"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnClear.Click += (_, _) => ClearForm();

        _btnDelete = new Button
        {
            Content = "🗑 삭제", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Background = AppTheme.BgDanger, Foreground = AppTheme.FgDanger,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            IsVisible = false,
        };
        _btnDelete.Click += (_, _) =>
        {
            if (_editId >= 0)
            {
                ScheduleService.Delete(_editId);
                ClearForm();
                RefreshEntries();
                RefreshCalendar();
            }
        };

        btnRow.Children.Add(_btnSave);
        btnRow.Children.Add(btnClear);
        btnRow.Children.Add(_btnDelete);
        right.Children.Add(btnRow);

        Grid.SetColumn(right, 2);
        twoCol.Children.Add(right);

        formPanel.Child = twoCol;
        Grid.SetRow(formPanel, 1);
        outer.Children.Add(formPanel);

        return outer;
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
            btn.Background  = a ? Brush.Parse(cs.Bg) : AppTheme.BgSecondary;
            btn.Foreground  = a ? Brush.Parse(cs.Fg) : AppTheme.FgDimmed;
            btn.BorderBrush = a ? Brush.Parse(cs.Bd) : AppTheme.BorderMuted;
        }
    }

    private void ApplySiteStyles()
    {
        foreach (var (k, btn) in _siteBtns)
        {
            bool a = k == _selSite;
            SiteStyle.TryGetValue(k, out var ss);
            btn.Background  = a ? Brush.Parse(ss.Bg) : AppTheme.BgSecondary;
            btn.Foreground  = a ? Brush.Parse(ss.Fg) : AppTheme.FgDimmed;
            btn.BorderBrush = a ? Brush.Parse(ss.Bd) : AppTheme.BorderMuted;
        }
    }

    // =========================================================================
    // CRUD
    // =========================================================================
    private (string Name, string Id) GetSingleAgentOrLogin()
    {
        if (_selectedAgents.Count > 0) return _selectedAgents[0];
        if (!string.IsNullOrEmpty(_agtName)) return (_agtName, _agtId);
        var id = ETA.Views.MainPage.CurrentEmployeeId;
        return (ETA.Services.SERVICE3.MyTaskService.GetEmployeeName(id), id);
    }

    private ScheduleEntry BuildEntry(string name, string id) => new()
    {
        날짜     = _selDate!,
        직원명   = name,
        직원id   = id,
        분류     = _selCat,
        사이트   = _selSite,
        업체약칭 = _selCompanyAbbr,
        제목     = "",
        내용     = _contentBox?.Text?.Trim() ?? "",
        시작시간 = "",
        종료시간 = "",
        첨부파일 = "",
        등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        등록자   = ETA.Views.MainPage.CurrentEmployeeId,
    };

    public void SaveEntry()
    {
        if (string.IsNullOrEmpty(_selDate)) return;
        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        if (_editId >= 0)
        {
            // 수정 모드
            var (name, id) = GetSingleAgentOrLogin();
            var upd = BuildEntry(name, id);
            upd.Id = _editId;
            upd.등록일시 = now;
            ScheduleService.Update(upd);
        }
        else if (_selCat == "채수" && _selectedAgents.Count > 0)
        {
            // 채수 다중 인원 등록 — 선택된 약칭 중복제거하여 업체약칭 설정
            _selCompanyAbbr = string.Join(",", _selReqs.Select(r => r.Abbr).Distinct());
            var names = new List<string>();
            foreach (var (name, id) in _selectedAgents)
            {
                var e = BuildEntry(name, id);
                e.등록일시 = now;
                ScheduleService.Insert(e);
                names.Add(name);
            }
            // 선택된 모든 의뢰에 채수담당자 등록
            foreach (var (reqId, _) in _selReqs)
                AnalysisRequestService.UpdateSamplers(reqId, names);
        }
        else
        {
            // 단일 등록
            var (name, id) = GetSingleAgentOrLogin();
            var e = BuildEntry(name, id);
            e.등록일시 = now;
            ScheduleService.Insert(e);
        }

        _lastSaved = true;
        RefreshProgress();
        ClearForm();
        RefreshEntries();
        RefreshCalendar();
    }

    public void ClearForm()
    {
        _editId         = -1;
        _selReqs.Clear();
        _selCompanyAbbr = "";
        _lastSaved      = false;
        if (_contentBox != null) _contentBox.Text = "";
        _agtName = ""; _agtId = "";
        _selectedAgents.Clear();
        RefreshAgentChips();
        RefreshCompanyBadge();
        RefreshProgress();
        if (_agentLbl  != null) { _agentLbl.Text = "👤 선택 안됨"; _agentLbl.Foreground = AppRes("FgMuted"); }
        if (_formTitle != null) _formTitle.Text  = "📋 일정 등록";
        if (_btnSave   != null) _btnSave.Content  = "💾 저장";
        if (_btnDelete != null) _btnDelete.IsVisible = false;
        if (_tree      != null) _tree.SelectedItem   = null;
    }

    private void RefreshEntries()
    {
        if (_entryList == null) return;
        _entryList.Children.Clear();

        if (string.IsNullOrEmpty(_selDate))
        {
            if (_entryDateLbl != null) _entryDateLbl.Text = "";
            return;
        }

        if (_entryDateLbl != null)
        {
            DateTime.TryParse(_selDate, out var dt);
            _entryDateLbl.Text = $"({dt:M/d})";
        }

        var list = ScheduleService.GetByDate(_selDate);
        if (list.Count == 0)
        {
            _entryList.Children.Add(new TextBlock
            {
                Text = "등록된 일정 없음",
                FontSize = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppRes("FgMuted"),
                Margin = new Thickness(4),
            });
            return;
        }

        foreach (var en in list)
        {
            var cs = GetCat(en.분류);
            var card = new Border
            {
                Background      = Brush.Parse(cs.Bg),
                BorderBrush     = Brush.Parse(cs.Bd),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 6),
                Margin          = new Thickness(0, 1),
            };
            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
            var info = new StackPanel { Spacing = 2 };

            // 첫 줄: 아이콘 + 분류 + 업체약칭/직원명
            var firstLine = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            firstLine.Children.Add(new TextBlock
            {
                Text = $"{cs.Icon} {en.분류}",
                FontSize = AppTheme.FontBase, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse(cs.Fg),
                VerticalAlignment = VerticalAlignment.Center,
            });
            var displayName = !string.IsNullOrEmpty(en.업체약칭) ? $"🏢 {en.업체약칭}" : "";
            if (!string.IsNullOrEmpty(en.직원명))
                displayName += (displayName.Length > 0 ? " · " : "") + $"👤 {en.직원명}";
            if (!string.IsNullOrEmpty(displayName))
                firstLine.Children.Add(new TextBlock
                {
                    Text = displayName, FontSize = AppTheme.FontBase, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
            info.Children.Add(firstLine);

            // 내용/제목
            if (!string.IsNullOrEmpty(en.제목))
                info.Children.Add(new TextBlock
                {
                    Text = en.제목, FontSize = AppTheme.FontSM, FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
            if (!string.IsNullOrEmpty(en.내용))
                info.Children.Add(new TextBlock
                {
                    Text = en.내용, FontSize = AppTheme.FontSM, FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 300,
                });

            grid.Children.Add(info);

            // 수정 버튼
            var capture = en;
            var edit = new Button
            {
                Content = "✏️", Width = 28, Height = 28, Padding = new Thickness(0),
                Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgInfo,
                BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderInfo,
                CornerRadius = new CornerRadius(4),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 2, 0),
                Cursor = new Cursor(StandardCursorType.Hand),
                [Grid.ColumnProperty] = 1,
            };
            edit.Click += (_, e2) =>
            {
                e2.Handled = true;
                LoadEntryToForm(capture);
            };
            grid.Children.Add(edit);

            // 삭제 버튼
            var entryId = en.Id;
            var del = new Button
            {
                Content = "🗑", Width = 28, Height = 28, Padding = new Thickness(0),
                Background = AppTheme.BgDanger, Foreground = AppTheme.FgDanger,
                BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderDanger,
                CornerRadius = new CornerRadius(4),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = new Cursor(StandardCursorType.Hand),
                [Grid.ColumnProperty] = 2,
            };
            del.Click += (_, e2) =>
            {
                e2.Handled = true;
                ScheduleService.Delete(entryId);
                RefreshEntries();
                RefreshCalendar();
            };
            grid.Children.Add(del);
            card.Child = grid;

            _entryList.Children.Add(card);
        }
    }

    private void LoadEntryToForm(ScheduleEntry en)
    {
        _editId         = en.Id;
        _selDate        = en.날짜;
        _selCat         = en.분류;
        _selSite        = en.사이트;
        _selCompanyAbbr = en.업체약칭;

        if (_dateLbl != null)
        {
            DateTime.TryParse(en.날짜, out var dt);
            _dateLbl.Text       = $"📅 {dt:yyyy-MM-dd}  ({DowKr(dt)})";
            _dateLbl.Foreground = AppRes("AppFg");
        }

        ApplyCatStyles();
        ApplySiteStyles();

        _selectedAgents.Clear();
        if (!string.IsNullOrEmpty(en.직원명))
            _selectedAgents.Add((en.직원명, en.직원id));
        RefreshAgentChips();
        RefreshCompanyBadge();

        if (_contentBox != null) _contentBox.Text = en.내용;
        if (_formTitle  != null) _formTitle.Text  = "✏️ 일정 수정";
        if (_btnSave   != null) _btnSave.Content  = "💾 수정";
        if (_btnDelete != null) _btnDelete.IsVisible = true;
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
    // =========================================================================
    // 진행 상태 표시
    // =========================================================================
    private void RefreshProgress()
    {
        if (_progressPanel == null) return;
        _progressPanel.Children.Clear();

        bool isSampl = _selCat == "채수";

        // 단계 정의
        // 채수: 날짜 → 의뢰(건수) → 직원 → 완료
        // 휴가: 날짜 → 직원 → 완료
        bool dateOk  = !string.IsNullOrEmpty(_selDate);
        bool reqOk   = _selReqs.Count > 0;
        bool agentOk = _selectedAgents.Count > 0;

        var stages = isSampl
            ? new (string Icon, string Label, bool Done, string? Extra)[]
              {
                ("📅", "날짜",                        dateOk,  null),
                ("📋", $"의뢰({_selReqs.Count}건)",   reqOk,   null),
                ("👤", $"직원({_selectedAgents.Count}명)", agentOk, null),
                ("✅", "완료",                         _lastSaved, null),
              }
            : new (string Icon, string Label, bool Done, string? Extra)[]
              {
                ("📅", "날짜",                        dateOk,  null),
                ("👤", $"직원({_selectedAgents.Count}명)", agentOk, null),
                ("✅", "완료",                         _lastSaved, null),
              };

        // 완료 버튼 활성화 조건
        bool canComplete = isSampl
            ? dateOk && reqOk && agentOk && !_lastSaved
            : dateOk && agentOk && !_lastSaved;

        for (int i = 0; i < stages.Length; i++)
        {
            if (i > 0)
                _progressPanel.Children.Add(new TextBlock
                {
                    Text = "→", FontSize = AppTheme.FontXS, FontFamily = Font,
                    Foreground = AppTheme.FgDimmed,
                    VerticalAlignment = VerticalAlignment.Center,
                });

            var (icon, label, done, _) = stages[i];
            bool isFinal  = i == stages.Length - 1;
            bool current  = !done && (i == 0 || stages[i - 1].Done);

            string bg = done    ? "#1a3a1a"
                      : current ? "#1a2a3a"
                      :           "#1a1a1a";
            string fg = done    ? "#88ee88"
                      : current ? "#88aaee"
                      :           "#444444";
            string bd = done    ? "#2a6a2a"
                      : current ? "#336699"
                      :           "#2a2a2a";

            if (isFinal && canComplete)
            {
                // 완료 단계: 클릭 가능한 버튼
                var btnDone = new Button
                {
                    Content = "✅ 완료",
                    FontSize = AppTheme.FontXS, FontFamily = Font,
                    Padding = new Thickness(6, 2),
                    Background      = AppTheme.BgActiveGreen,
                    Foreground      = AppTheme.FgSuccess,
                    BorderBrush     = AppTheme.BorderActive,
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(8),
                    Cursor          = new Cursor(StandardCursorType.Hand),
                };
                btnDone.Click += (_, _) => SaveEntry();
                _progressPanel.Children.Add(btnDone);
            }
            else
            {
                var stepText = done ? $"✓ {label}" : $"{icon} {label}";
                var step = new Border
                {
                    Background      = Brush.Parse(bg),
                    BorderBrush     = Brush.Parse(bd),
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(8),
                    Padding         = new Thickness(5, 2),
                };
                step.Child = new TextBlock
                {
                    Text       = stepText,
                    FontSize   = AppTheme.FontXS, FontFamily = Font,
                    Foreground = Brush.Parse(fg),
                    FontWeight = current ? FontWeight.SemiBold : FontWeight.Normal,
                };
                _progressPanel.Children.Add(step);
            }
        }
    }

    private static TextBlock FormLbl(string text) => new()
    {
        Text = text, FontSize = AppTheme.FontSM, FontFamily = Font,
        Foreground = AppTheme.FgMuted,
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

    /// <summary>"20240401" → "2024-04-01", "2024/04/01" → "2024-04-01", 이미 하이픈이면 그대로</summary>
    private static string NormalizeDate(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return raw;
        raw = raw.Trim().Replace("/", "-");
        if (raw.Length == 8 && !raw.Contains('-') && int.TryParse(raw, out _))
            return $"{raw[..4]}-{raw[4..6]}-{raw[6..8]}";
        return raw;
    }

}
