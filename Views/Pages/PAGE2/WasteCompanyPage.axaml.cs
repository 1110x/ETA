using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE2;

public partial class WasteCompanyPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }
    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    public event Action<Control?>?   DetailPanelChanged;
    public event Action<WasteCompany>? CompanySelected;   // 일반 선택 모드에서 업체 클릭 시
    public event Action?             OrderSaved;          // 의뢰 저장 완료 알림

    // ── 상태 ────────────────────────────────────────────────────────────────
    private WasteCompany? _selectedCompany;
    public  WasteCompany? SelectedCompany => _selectedCompany;
    private StackPanel?   _detailPanel;
    private bool          _isAddMode   = false;
    private string        _activeGroup = "여수";

    // ── 의뢰등록 모드 ────────────────────────────────────────────────────────
    private bool              _orderMode     = false;
    private string            _orderDate     = DateTime.Today.ToString("yyyy-MM-dd");
    private HashSet<string>   _orderExisting = new();   // 해당 날짜 이미 등록된 업체
    private List<WasteCompany> _orderSelected = new();  // 이번에 추가할 업체
    private StackPanel?       _orderListPanel;          // show2 내 선택목록 영역
    private TextBox?          _확인자Box;

    public WasteCompanyPage()
    {
        InitializeComponent();
    }

    // =========================================================================
    // 데이터 로드 (현재 _activeGroup 만 표시)
    // =========================================================================
    private List<WasteCompany> _allItems = new();

    public void LoadData()
    {
        WasteCompanyTreeView.Items.Clear();
        _selectedCompany = null;
        _isAddMode       = false;
        DetailPanelChanged?.Invoke(null);

        try
        {
            _allItems = WasteCompanyService.GetAllItems()
                            .Where(c => GetGroupLabel(c) == _activeGroup)
                            .OrderBy(c => c.관리번호)
                            .ToList();

            ApplyFilter(SearchBox?.Text ?? "");
            Log($"로드 완료 → {_activeGroup} {_allItems.Count}건");
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    private void ApplyFilter(string keyword)
    {
        WasteCompanyTreeView.Items.Clear();
        var q = string.IsNullOrWhiteSpace(keyword) ? _allItems
              : _allItems.Where(c =>
                    c.업체명.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                    c.약칭.Contains(keyword,   StringComparison.OrdinalIgnoreCase));
        foreach (var item in q)
            WasteCompanyTreeView.Items.Add(CreateCompanyNode(item));
    }

    private void SearchBox_TextChanged(object? sender, Avalonia.Controls.TextChangedEventArgs e)
        => ApplyFilter(SearchBox.Text ?? "");

    // =========================================================================
    // 의뢰등록 버튼 토글
    // =========================================================================
    public void OrderBtn_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        _orderMode = !_orderMode;

        if (_orderMode)
        {
            // 모드 진입 (amber 강조 — 테마 무관)
            BtnOrder.Background  = new SolidColorBrush(Color.Parse("#92400e"));
            BtnOrder.Foreground  = new SolidColorBrush(Color.Parse("#fbbf24"));
            BtnOrder.BorderBrush = new SolidColorBrush(Color.Parse("#d97706"));
            _orderDate     = DateTime.Today.ToString("yyyy-MM-dd");
            _orderSelected.Clear();
            _orderExisting = WasteSampleService.GetCompanyNamesForDate(_orderDate);
            LoadData();
            DetailPanelChanged?.Invoke(BuildOrderPanel());
        }
        else
        {
            // 모드 해제
            BtnOrder.Background  = AppRes("SubBtnBg");
            BtnOrder.Foreground  = AppRes("FgMuted");
            BtnOrder.BorderBrush = AppRes("InputBorder");
            _orderSelected.Clear();
            _orderExisting.Clear();
            LoadData();
            DetailPanelChanged?.Invoke(null);
        }
    }

    // =========================================================================
    // 의뢰등록 show2 패널
    // =========================================================================
    private StackPanel BuildOrderPanel()
    {
        var root = MakeRootPanel("📋  채수의뢰 등록");

        // 날짜 선택
        root.Children.Add(new TextBlock
        {
            Text = "채수일", FontSize = 11, FontFamily = Font,
            Foreground = AppRes("FgMuted"),
        });

        DateTime.TryParse(_orderDate, out var initDate);
        var dp = new CalendarDatePicker
        {
            SelectedDate = initDate == DateTime.MinValue ? DateTime.Today : initDate,
            FontFamily = Font, FontSize = 12,
            Margin = new Thickness(0, 0, 0, 8),
        };
        dp.SelectedDateChanged += (_, _) =>
        {
            if (dp.SelectedDate == null) return;
            _orderDate     = dp.SelectedDate.Value.ToString("yyyy-MM-dd");
            _orderSelected.Clear();
            _orderExisting = WasteSampleService.GetCompanyNamesForDate(_orderDate);
            LoadData();
            RefreshOrderList();
        };
        root.Children.Add(dp);

        // 확인자
        root.Children.Add(new TextBlock
        {
            Text = "확인자", FontSize = 11, FontFamily = Font,
            Foreground = AppRes("FgMuted"),
        });
        _확인자Box = new TextBox
        {
            FontFamily = Font, FontSize = 12,
            Background = AppRes("InputBg"),
            Foreground = AppRes("InputFg"),
            BorderBrush = AppRes("InputBorder"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 4),
            Margin = new Thickness(0, 0, 0, 8),
        };
        root.Children.Add(_확인자Box);

        // 선택 업체 목록
        root.Children.Add(new TextBlock
        {
            Text = "선택된 업체", FontSize = 11, FontFamily = Font,
            Foreground = AppRes("FgMuted"),
        });
        _orderListPanel = new StackPanel { Spacing = 3, Margin = new Thickness(0, 4, 0, 8) };
        root.Children.Add(_orderListPanel);

        // 저장 / 취소
        var btnSave = new Button
        {
            Content = "저장", Width = 80, Height = 28,
            Background = new SolidColorBrush(Color.Parse("#15803d")),
            Foreground = Brushes.White,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 12,
        };
        var btnCancel = new Button
        {
            Content = "취소", Width = 80, Height = 28,
            Background = AppRes("SubBtnBg"),
            Foreground = AppRes("FgMuted"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 12,
        };
        btnSave.Click += (_, _) =>
        {
            string 확인자 = _확인자Box?.Text?.Trim() ?? "";
            foreach (var c in _orderSelected)
                WasteSampleService.Insert(_orderDate, GetGroupLabel(c),
                    c.업체명, c.관리번호, 확인자);
            _orderExisting = WasteSampleService.GetCompanyNamesForDate(_orderDate);
            _orderSelected.Clear();
            LoadData();
            RefreshOrderList();
            OrderSaved?.Invoke();
        };
        btnCancel.Click += (_, _) =>
        {
            _orderMode = false;
            BtnOrder.Background  = AppRes("SubBtnBg");
            BtnOrder.Foreground  = AppRes("FgMuted");
            BtnOrder.BorderBrush = AppRes("InputBorder");
            _orderSelected.Clear();
            _orderExisting.Clear();
            LoadData();
            DetailPanelChanged?.Invoke(null);
        };

        root.Children.Add(new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8,
            Children = { btnSave, btnCancel },
        });

        return root;
    }

    private void RefreshOrderList()
    {
        if (_orderListPanel == null) return;
        _orderListPanel.Children.Clear();
        foreach (var c in _orderSelected)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            row.Children.Add(new TextBlock
            {
                Text = "✅", FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
            });
            row.Children.Add(new TextBlock
            {
                Text = c.업체명, FontSize = 11, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse("#16a34a")),
                VerticalAlignment = VerticalAlignment.Center,
            });
            _orderListPanel.Children.Add(row);
        }
        if (_orderSelected.Count == 0)
            _orderListPanel.Children.Add(new TextBlock
            {
                Text = "왼쪽 트리에서 업체를 선택하세요",
                FontSize = 10, FontFamily = Font,
                Foreground = AppRes("FgMuted"),
            });
    }

    // =========================================================================
    // 그룹 탭 토글
    // =========================================================================
    public void GroupTab_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        _activeGroup = btn.Tag as string ?? "여수";
        UpdateTabStyle();
        LoadData();
    }

    private static readonly Dictionary<string, (string bg, string fg, string border)> TabColors = new()
    {
        ["여수"] = ("#1e3a5a", "#88aacc", "#336699"),
        ["율촌"] = ("#1a3a1a", "#aaccaa", "#336633"),
        ["세풍"] = ("#3a2a1a", "#ccaa88", "#996633"),
    };

    private void UpdateTabStyle()
    {
        foreach (var btn in new[] { BtnYeosoo, BtnYulchon, BtnSepung })
        {
            var group  = btn.Tag as string ?? "";
            bool active = group == _activeGroup;
            if (active && TabColors.TryGetValue(group, out var c))
            {
                btn.Background  = new SolidColorBrush(Color.Parse(c.bg));
                btn.Foreground  = new SolidColorBrush(Color.Parse(c.fg));
                btn.BorderBrush = new SolidColorBrush(Color.Parse(c.border));
            }
            else
            {
                btn.Background  = AppRes("SubBtnBg");
                btn.Foreground  = AppRes("FgMuted");
                btn.BorderBrush = AppRes("InputBorder");
            }
        }
    }

    // =========================================================================
    // 그룹 라벨
    // =========================================================================
    private static string GetGroupLabel(WasteCompany c)
    {
        var pname = c.프로젝트명 ?? "";
        if (pname.Contains("여수")) return "여수";
        if (pname.Contains("세풍")) return "세풍";
        if (pname.Contains("율촌")) return "율촌";
        return "기타";
    }

    // =========================================================================
    // 그룹 노드 (여수 / 율촌)
    // =========================================================================
    private static TreeViewItem CreateGroupNode(string groupName, int count)
    {
        var (icon, color) = groupName switch
        {
            "여수" => ("🌊", "#88aacc"),
            "율촌" => ("🏗", "#aaccaa"),
            _      => ("📁", "#aaaaaa"),
        };

        return new TreeViewItem
        {
            IsExpanded = true,
            Header = new StackPanel
            {
                Orientation       = Orientation.Horizontal,
                Spacing           = 6,
                VerticalAlignment = VerticalAlignment.Center,
                Children =
                {
                    new TextBlock
                    {
                        Text = icon,
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindLG(),
                    new TextBlock
                    {
                        Text = groupName, FontWeight = FontWeight.SemiBold,
                        FontFamily = Font,
                        Foreground = new SolidColorBrush(Color.Parse(color)),
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindLG(),
                    new Border
                    {
                        Background   = new SolidColorBrush(Color.Parse("#2a2a3a")),
                        CornerRadius = new CornerRadius(8),
                        Padding      = new Thickness(5, 1),
                        VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = count.ToString(), FontFamily = Font,
                            Foreground = new SolidColorBrush(Color.Parse("#888888")),
                        }.BindXS()
                    }
                }
            }
        };
    }

    // =========================================================================
    // 업체 노드 — 의뢰등록 모드에서 🔴(이미등록) / ✅(이번선택) 표시
    // =========================================================================
    private TreeViewItem CreateCompanyNode(WasteCompany company)
    {
        bool alreadyOrdered = _orderMode && _orderExisting.Contains(company.업체명);
        bool selectedNow    = _orderMode && _orderSelected.Any(c => c.업체명 == company.업체명);

        string indicator = selectedNow    ? "✅"
                         : alreadyOrdered ? "🔴"
                         : "";

        var sp = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 4,
            VerticalAlignment = VerticalAlignment.Center,
        };

        if (_orderMode)
            sp.Children.Add(new TextBlock
            {
                Text = indicator.Length > 0 ? indicator : "⬜",
                VerticalAlignment = VerticalAlignment.Center,
                Width = 18,
            }.BindBase());

        if (!string.IsNullOrWhiteSpace(company.약칭))
        {
            var (bg, fg, bd) = GetChosungBadgeColor(company.약칭);
            sp.Children.Add(new Border
            {
                Background        = new SolidColorBrush(Color.Parse(bg)),
                BorderBrush       = new SolidColorBrush(Color.Parse(bd)),
                BorderThickness   = new Thickness(1),
                CornerRadius      = new CornerRadius(8),
                Padding           = new Thickness(5, 1),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = company.약칭, FontFamily = Font,
                    Foreground = new SolidColorBrush(Color.Parse(fg)),
                }.BindXS(),
            });
        }
        sp.Children.Add(new TextBlock
        {
            Text = company.업체명, FontFamily = Font,
            Foreground = selectedNow    ? new SolidColorBrush(Color.Parse("#aaffaa"))
                       : alreadyOrdered ? new SolidColorBrush(Color.Parse("#ff8888"))
                       : AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        }.BindMD());
        sp.Children.Add(new TextBlock
        {
            Text = company.관리번호, FontFamily = Font,
            Foreground = new SolidColorBrush(Color.Parse("#666666")),
            VerticalAlignment = VerticalAlignment.Center,
        }.BindXS());

        return new TreeViewItem { Tag = company, Header = sp };
    }

    // =========================================================================
    // 트리 선택 → 수정 패널
    // =========================================================================
    public void WasteCompanyTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        WasteCompany? company = null;

        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is WasteCompany c1) company = c1;
            else if (e.AddedItems[0] is WasteCompany c2) company = c2;
        }

        if (company == null) return;

        // ── 의뢰등록 모드 ──────────────────────────────────────────────────
        if (_orderMode)
        {
            // 이미 선택됐으면 해제, 아니면 추가
            var existing = _orderSelected.FirstOrDefault(c => c.업체명 == company.업체명);
            if (existing != null)
                _orderSelected.Remove(existing);
            else
                _orderSelected.Add(company);
            LoadData();   // 트리 재렌더 (✅ 표시 갱신)
            RefreshOrderList();
            return;
        }

        // ── 일반 편집 모드 ─────────────────────────────────────────────────
        _selectedCompany = company;
        _isAddMode       = false;
        _detailPanel     = BuildEditPanel(company);
        DetailPanelChanged?.Invoke(_detailPanel);   // → Show4 (MainPage에서 라우팅)
        CompanySelected?.Invoke(company);            // → Show2/Show3 데이터 로드
        Log($"선택: {company.업체명}");
    }

    // =========================================================================
    // 저장 (MainPage BT1)
    // =========================================================================
    public void SaveSelected()
    {
        if (_selectedCompany == null || _detailPanel == null)
        {
            Log("저장 스킵: 선택 없음"); return;
        }

        SyncPanelToCompany(_detailPanel, _selectedCompany);

        bool ok = WasteCompanyService.Update(_selectedCompany);
        Log(ok ? $"✅ 수정 저장: {_selectedCompany.업체명}"
               : $"❌ 수정 실패: {_selectedCompany.업체명}");

        if (ok)
            LoadData();
    }

    // =========================================================================
    // 패널 — 수정 모드
    // =========================================================================
    private StackPanel BuildEditPanel(WasteCompany c)
    {
        var root = MakeRootPanel($"🏭  {c.업체명} — 업소 정보");
        var grid = MakeTwoColumnGrid(6);

        AddGridRow(grid, 0, "관리번호",  c.관리번호,   isReadOnly: true, isLocked: true);
        AddGridRow(grid, 1, "업체명",    c.업체명);
        AddAbbrevRow(grid, 2, c.약칭);
        AddGridRow(grid, 3, "프로젝트",  c.프로젝트);
        AddGridRow(grid, 4, "프로젝트명", c.프로젝트명);
        AddGridRow(grid, 5, "사업자번호", c.사업자번호);

        root.Children.Add(grid);
        return root;
    }

    // =========================================================================
    // UI → WasteCompany 동기화
    // =========================================================================
    private static void SyncPanelToCompany(StackPanel panel, WasteCompany c)
    {
        var grid = panel.Children.OfType<Grid>().FirstOrDefault();
        if (grid == null) return;

        foreach (var child in grid.Children.OfType<StackPanel>())
        {
            if (child.Children.Count < 2) continue;

            var labelBlock = child.Children[0] as TextBlock;
            var label = (labelBlock?.Text ?? "")
                .Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();

            var tb = child.Children[1] as TextBox;
            if (tb == null || tb.IsReadOnly) continue;

            switch (label)
            {
                case "업체명":     c.업체명     = tb.Text ?? ""; break;
                case "약칭":       c.약칭       = tb.Text ?? ""; break;
                case "프로젝트":   c.프로젝트   = tb.Text ?? ""; break;
                case "프로젝트명": c.프로젝트명 = tb.Text ?? ""; break;
                case "사업자번호": c.사업자번호 = tb.Text ?? ""; break;
            }
        }
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    private static StackPanel MakeRootPanel(string title)
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock
        {
            Text       = title,
            FontSize   = 15,
            FontFamily = Font,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = new SolidColorBrush(Color.Parse("#555555")),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        return root;
    }

    private static Grid MakeTwoColumnGrid(int rowCount)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*"),
            RowSpacing        = 6,
        };
        for (int i = 0; i < rowCount; i++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        return grid;
    }

    private static void AddGridRow(Grid grid, int row, string label, string value,
                                   bool isReadOnly = false, bool isLocked = false)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

        panel.Children.Add(new TextBlock
        {
            Text              = (isLocked ? "🔒 " : "    ") + label + " :",
            Width             = 130,
            FontSize          = 12,
            FontFamily        = Font,
            Foreground        = isLocked
                                    ? new SolidColorBrush(Color.Parse("#888888"))
                                    : AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = 12,
            FontFamily      = Font,
            IsReadOnly      = isReadOnly,
            Background      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#252525"))
                                  : new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#666666"))
                                  : AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#333333"))
                                  : new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4),
        });

        Grid.SetColumn(panel, 0);
        Grid.SetRow(panel, row);
        grid.Children.Add(panel);
    }

    // ── 초성 추출 ──────────────────────────────────────────────────────────────
    private static readonly char[] Chosungs =
        { 'ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ' };

    // 약칭 첫 글자의 초성으로 배지 색 결정
    // ㄱ/ㄲ=파랑  ㄴ=청록  ㄷ/ㄸ=하늘  ㄹ=민트  ㅁ=초록  ㅂ/ㅃ=연두
    // ㅅ/ㅆ=노랑  ㅇ=주황  ㅈ/ㅉ=분홍  ㅊ=빨강  ㅋ=자주  ㅌ=보라  ㅍ=남보라  ㅎ=회파랑
    public static (string Bg, string Fg, string Bd) GetChosungBadgeColorPublic(string 약칭)
        => GetChosungBadgeColor(약칭);

    private static (string Bg, string Fg, string Bd) GetChosungBadgeColor(string 약칭)
    {
        if (string.IsNullOrEmpty(약칭)) return ("#2a2a3a", "#888888", "#444444");

        char first = 약칭[0];
        char cho = first;
        if (first >= 0xAC00 && first <= 0xD7A3)
            cho = Chosungs[(first - 0xAC00) / (21 * 28)];

        return cho switch
        {
            'ㄱ' or 'ㄲ' => ("#1a2a4a", "#88aaee", "#2255bb"),
            'ㄴ'         => ("#0f3030", "#55ddcc", "#1a8877"),
            'ㄷ' or 'ㄸ' => ("#0f2a3a", "#66ccee", "#1a6688"),
            'ㄹ'         => ("#0f3a2a", "#66ddaa", "#1a8855"),
            'ㅁ'         => ("#1a3a1a", "#88cc77", "#336633"),
            'ㅂ' or 'ㅃ' => ("#2a3a10", "#aabb55", "#556622"),
            'ㅅ' or 'ㅆ' => ("#3a3010", "#ddcc55", "#886622"),
            'ㅇ'         => ("#3a2010", "#ddaa55", "#885522"),
            'ㅈ' or 'ㅉ' => ("#3a1a2a", "#dd88aa", "#882255"),
            'ㅊ'         => ("#3a1010", "#ee7777", "#882222"),
            'ㅋ'         => ("#3a1a3a", "#cc77cc", "#772277"),
            'ㅌ'         => ("#2a1a3a", "#aa88dd", "#553388"),
            'ㅍ'         => ("#1a1a3a", "#8899dd", "#222288"),
            'ㅎ'         => ("#1a2a3a", "#7799bb", "#224466"),
            _            => ("#2a2a3a", "#999999", "#555555"),
        };
    }

    // 약칭 전용 행 — 초성 자동입력 버튼 포함
    private static void AddAbbrevRow(Grid grid, int row, string value)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

        panel.Children.Add(new TextBlock
        {
            Text              = "    약칭 :",
            Width             = 130,
            FontSize          = 12,
            FontFamily        = Font,
            Foreground        = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = 12,
            FontFamily      = Font,
            Background      = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4),
        });

        Grid.SetColumn(panel, 0);
        Grid.SetRow(panel, row);
        grid.Children.Add(panel);
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [WasteCompany] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("Logs/WasteCompanyDebug.log", line + Environment.NewLine); } catch { }
    }
}
