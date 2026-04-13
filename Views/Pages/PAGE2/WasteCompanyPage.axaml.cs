using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Views;

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

    // ── 상태 ────────────────────────────────────────────────────────────────
    private WasteCompany? _selectedCompany;
    public  WasteCompany? SelectedCompany => _selectedCompany;
    private StackPanel?   _detailPanel;
    private bool          _isAddMode   = false;
    private string        _activeGroup = "여수";


    public WasteCompanyPage()
    {
        InitializeComponent();
    }

    // =========================================================================
    // 데이터 로드 (현재 _activeGroup 만 표시)
    // =========================================================================
    private List<WasteCompany> _allItems = new();

    public async Task LoadDataAsync()
    {
        WasteCompanyTreeView.Items.Clear();
        _selectedCompany = null;
        _isAddMode       = false;
        DetailPanelChanged?.Invoke(null);

        try
        {
            var activeGroup = _activeGroup;
            var items = await Task.Run(() =>
                WasteCompanyService.GetAllItems()
                    .Where(c => GetGroupLabel(c) == activeGroup)
                    .OrderBy(c => c.관리번호)
                    .ToList());

            _allItems = items;
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
    // 그룹 탭 토글
    // =========================================================================
    public void GroupTab_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        _activeGroup = btn.Tag as string ?? "여수";
        UpdateTabStyle();
        _ = LoadDataAsync();
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

        var tvi = new TreeViewItem
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
                        Background   = AppTheme.BorderSubtle,
                        CornerRadius = new CornerRadius(8),
                        Padding      = new Thickness(5, 1),
                        VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = count.ToString(), FontFamily = Font,
                            Foreground = AppTheme.FgMuted,
                        }.BindXS()
                    }
                }
            }
        };
        TextShimmer.AttachHover(tvi);
        return tvi;
    }

    // =========================================================================
    // 업체 노드 — 의뢰등록 모드에서 🔴(이미등록) / ✅(이번선택) 표시
    // =========================================================================
    private TreeViewItem CreateCompanyNode(WasteCompany company)
    {
        var sp = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 4,
            VerticalAlignment = VerticalAlignment.Center,
        };

        if (false) // unused
            sp.Children.Add(new TextBlock
            {
                Text = "",
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
            Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        }.BindMD());
        sp.Children.Add(new TextBlock
        {
            Text = company.관리번호, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
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


        // ── 일반 편집 모드 ─────────────────────────────────────────────────
        _selectedCompany = company;
        _isAddMode       = false;
        _detailPanel     = BuildEditPanel(company);
        DetailPanelChanged?.Invoke(_detailPanel);   // → Show4 (MainPage에서 라우팅)
        CompanySelected?.Invoke(company);            // → Show2/Show3 데이터 로드
        Log($"선택: {company.업체명}");
    }

    // =========================================================================
    // 업소 등록 (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
        _isAddMode       = true;
        _selectedCompany = new WasteCompany
        {
            프로젝트  = _activeGroup,
            프로젝트명 = _activeGroup,
        };
        _detailPanel = BuildEditPanel(_selectedCompany, isNew: true);
        DetailPanelChanged?.Invoke(_detailPanel);
        Log($"신규 등록 모드 진입 ({_activeGroup})");
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

        if (_isAddMode)
        {
            _selectedCompany.Original업체명 = _selectedCompany.업체명;
            bool ok = WasteCompanyService.Insert(_selectedCompany);
            Log(ok ? $"✅ 신규 등록: {_selectedCompany.업체명}"
                   : $"❌ 등록 실패: {_selectedCompany.업체명}");
            if (ok)
            {
                _isAddMode = false;
                _ = LoadDataAsync();
                DetailPanelChanged?.Invoke(null);
            }
        }
        else
        {
            bool ok = WasteCompanyService.Update(_selectedCompany);
            Log(ok ? $"✅ 수정 저장: {_selectedCompany.업체명}"
                   : $"❌ 수정 실패: {_selectedCompany.업체명}");
            if (ok)
                _ = LoadDataAsync();
        }
    }

    // =========================================================================
    // 패널 — 수정 / 신규 모드
    // =========================================================================
    private StackPanel BuildEditPanel(WasteCompany c, bool isNew = false)
    {
        var title = isNew ? "🆕  신규 업소 등록" : $"🏭  {c.업체명} — 업소 정보";
        var root = MakeRootPanel(title);

        // ── 기본 정보 ──────────────────────────────────────────────────────
        var grid = MakeTwoColumnGrid(6);
        AddGridRow(grid, 0, "관리번호",  c.관리번호,   isReadOnly: !isNew, isLocked: !isNew);
        AddGridRow(grid, 1, "업체명",    c.업체명);
        AddAbbrevRow(grid, 2, c.약칭);
        AddGridRow(grid, 3, "프로젝트",  c.프로젝트);
        AddGridRow(grid, 4, "프로젝트명", c.프로젝트명);
        AddGridRow(grid, 5, "사업자번호", c.사업자번호);
        root.Children.Add(grid);

        // ── 허용기준 섹션 구분선 ────────────────────────────────────────────
        root.Children.Add(new TextBlock
        {
            Text = "허용기준",
            FontSize = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("FgMuted"),
            Margin = new Thickness(0, 10, 0, 2),
        });
        root.Children.Add(new Border
        {
            Height = 1, Background = AppTheme.BorderDefault,
            Margin = new Thickness(0, 0, 0, 6),
        });

        var grid2 = MakeTwoColumnGrid(9);
        AddGridRow(grid2, 0, "BOD",          c.BOD);
        AddGridRow(grid2, 1, "TOC",          c.TOC);
        AddGridRow(grid2, 2, "SS",           c.SS);
        AddGridRow(grid2, 3, "T-N",          c.TN);
        AddGridRow(grid2, 4, "T-P",          c.TP);
        AddGridRow(grid2, 5, "Phenols",      c.Phenols);
        AddGridRow(grid2, 6, "N-Hexan",      c.NHexan);
        AddGridRow(grid2, 7, "승인유량",     c.승인유량);
        AddGridRow(grid2, 8, "기타특이사항", c.기타특이사항);
        root.Children.Add(grid2);

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
                case "관리번호":   c.관리번호   = tb.Text ?? ""; break;
                case "업체명":     c.업체명     = tb.Text ?? ""; break;
                case "약칭":       c.약칭       = tb.Text ?? ""; break;
                case "프로젝트":   c.프로젝트   = tb.Text ?? ""; break;
                case "프로젝트명": c.프로젝트명 = tb.Text ?? ""; break;
                case "사업자번호": c.사업자번호 = tb.Text ?? ""; break;
                case "BOD":          c.BOD          = tb.Text ?? ""; break;
                case "TOC":          c.TOC          = tb.Text ?? ""; break;
                case "SS":           c.SS           = tb.Text ?? ""; break;
                case "T-N":          c.TN           = tb.Text ?? ""; break;
                case "T-P":          c.TP           = tb.Text ?? ""; break;
                case "Phenols":      c.Phenols      = tb.Text ?? ""; break;
                case "N-Hexan":      c.NHexan       = tb.Text ?? ""; break;
                case "승인유량":     c.승인유량     = tb.Text ?? ""; break;
                case "기타특이사항": c.기타특이사항 = tb.Text ?? ""; break;
            }
        }
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    private static StackPanel MakeRootPanel(string title)
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock
        {
            Text       = title,
            FontSize   = AppTheme.FontXL,
            FontFamily = Font,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderDefault,
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
            FontSize          = AppTheme.FontMD,
            FontFamily        = Font,
            Foreground        = isLocked
                                    ? AppTheme.FgMuted
                                    : AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            IsReadOnly      = isReadOnly,
            Background      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#252525"))
                                  : AppTheme.BorderSeparator,
            Foreground      = isReadOnly
                                  ? AppTheme.BorderMuted
                                  : AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? AppTheme.BorderSubtle
                                  : AppTheme.BorderDefault,
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
            FontSize          = AppTheme.FontMD,
            FontFamily        = Font,
            Foreground        = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            Background      = AppTheme.BorderSeparator,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderDefault,
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
        if (App.EnableLogging)
        {
            try { File.AppendAllText("Logs/WasteCompanyDebug.log", line + Environment.NewLine); } catch { }
        }
    }
}
