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
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE2;

public partial class WasteSampleListPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;

    private string? _selectedDate;
    private Panel?  _detailPanel;

    // ── Show4: 업체 트리 ─────────────────────────────────────────────────────
    private Control?   _companyTreePanel;
    private TreeView?  _companyTreeView;
    private TextBlock? _companyTreeStatus;
    private string     _companyActiveGroup = "여수";
    private Button?    _btnCo여수;
    private Button?    _btnCo율촌;
    private Button?    _btnCo세풍;
    private TextBox?   _companySearchBox;

    public Control CompanyTreePanel => _companyTreePanel ??= BuildCompanyTreePanel();

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    public WasteSampleListPage() => InitializeComponent();

    // =========================================================================
    // 데이터 로드 (날짜 트리뷰)
    // =========================================================================
    public void LoadData()
    {
        DateTreeView.Items.Clear();
        _selectedDate = null;
        DetailPanelChanged?.Invoke(null);

        try
        {
            var dates = WasteSampleService.GetDates();   // YYYY-MM-DD 역순

            // 연-월 그룹핑
            var byMonth = dates
                .GroupBy(d => d[..7])          // "YYYY-MM"
                .OrderByDescending(g => g.Key);

            foreach (var month in byMonth)
            {
                var monthNode = MakeMonthNode(month.Key, month.Count());
                foreach (var d in month.OrderByDescending(x => x))
                {
                    var samples   = WasteSampleService.GetByDate(d);
                    var dateNode  = MakeDateNode(d, samples);
                    monthNode.Items.Add(dateNode);
                }
                monthNode.IsExpanded = true;
                DateTreeView.Items.Add(monthNode);
            }

            Log($"날짜 로드 → {dates.Count}건");
        }
        catch (Exception ex) { Log("크래시: " + ex.Message); }
    }

    // =========================================================================
    // 날짜 트리 노드
    // =========================================================================
    private static TreeViewItem MakeMonthNode(string ym, int count)
    {
        if (!DateTime.TryParse(ym + "-01", out var d)) d = DateTime.MinValue;
        return new TreeViewItem
        {
            IsExpanded = true,
            Header = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock
                    {
                        Text = $"📅  {d.Year}년 {d.Month}월",
                        FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindMD(),
                    new Border
                    {
                        Background = Brush.Parse("#2a2a3a"), CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5,1),
                        Child = new TextBlock
                        {
                            Text = count.ToString(),
                            FontFamily = Font, Foreground = Brush.Parse("#666"),
                        }.BindXS()
                    }
                }
            }
        };
    }

    private static TreeViewItem MakeDateNode(string dateStr, List<WasteSample> samples)
    {
        DateTime.TryParse(dateStr, out var d);
        string dow   = d == DateTime.MinValue ? "" : $" ({DayOfWeekKr(d)})";
        string label = d == DateTime.MinValue ? dateStr : $"🗓  {d.Month}/{d.Day}{dow}";

        var node = new TreeViewItem
        {
            Tag        = dateStr,
            IsExpanded = true,
            Header     = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock
                    {
                        Text = label, FontFamily = Font,
                        Foreground = Brush.Parse("#dddddd"),
                    }.BindMD(),
                    new Border
                    {
                        Background = Brush.Parse("#2a2a3a"), CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = samples.Count.ToString(),
                            FontFamily = Font, Foreground = Brush.Parse("#666"),
                        }.BindXS(),
                    },
                }
            }
        };

        foreach (var s in samples)
            node.Items.Add(MakeSampleLeafNode(s, dateStr));

        return node;
    }

    private static readonly Dictionary<string, (string Bg, string Fg)> _groupColors = new()
    {
        ["여수"] = ("#1e3a5a", "#88aacc"),
        ["율촌"] = ("#1a3a1a", "#aaccaa"),
        ["세풍"] = ("#3a2a1a", "#ccaa88"),
    };

    private static TreeViewItem MakeSampleLeafNode(WasteSample s, string dateStr)
    {
        var (bg, fg) = _groupColors.TryGetValue(s.구분, out var gc) ? gc : ("#2a2a3a", "#999");

        var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };

        // 구분 배지
        sp.Children.Add(new Border
        {
            Background      = Brush.Parse(bg),
            CornerRadius    = new CornerRadius(8),
            Padding         = new Thickness(5, 1),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = s.구분, FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse(fg),
            },
        });

        // SN
        if (!string.IsNullOrEmpty(s.SN))
            sp.Children.Add(new TextBlock
            {
                Text = s.SN, FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse("#666"),
                VerticalAlignment = VerticalAlignment.Center,
            });

        // 업체명
        sp.Children.Add(new TextBlock
        {
            Text = s.업체명, FontSize = 11, FontFamily = Font,
            Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        return new TreeViewItem { Tag = dateStr, Header = sp };
    }

    private static string DayOfWeekKr(DateTime d) => d.DayOfWeek switch
    {
        DayOfWeek.Monday    => "월",
        DayOfWeek.Tuesday   => "화",
        DayOfWeek.Wednesday => "수",
        DayOfWeek.Thursday  => "목",
        DayOfWeek.Friday    => "금",
        DayOfWeek.Saturday  => "토",
        _                   => "일",
    };

    // =========================================================================
    // 날짜 선택 → 상세 패널
    // =========================================================================
    public void DateTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0) return;
        if (e.AddedItems[0] is not TreeViewItem tvi || tvi.Tag is not string date) return;

        _selectedDate = date;
        RefreshDetail();
    }

    // =========================================================================
    // + 날짜 추가 — Show1 인라인 캘린더 토글
    // =========================================================================
    public void AddNewDate() => BtnNewDate_Click(null, null!);

    private void BtnNewDate_Click(object? sender, RoutedEventArgs e)
    {
        bool nowVisible = !CalendarPickerBorder.IsVisible;
        CalendarPickerBorder.IsVisible = nowVisible;
        if (nowVisible)
        {
            InlineCalendar.SelectedDates.Clear();
            InlineCalendar.DisplayDate = DateTime.Today;
        }
    }

    private void InlineCalendar_DateChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (InlineCalendar?.SelectedDate == null) return;
        string date = InlineCalendar.SelectedDate.Value.ToString("yyyy-MM-dd");
        CalendarPickerBorder.IsVisible = false;
        _selectedDate = date;
        EnsureDateInTree(date);
        RefreshDetail();
    }

    // =========================================================================
    // 상세 패널 새로고침
    // =========================================================================
    private void RefreshDetail()
    {
        if (_selectedDate == null) return;
        _detailPanel = BuildSamplePanel(_selectedDate);
        DetailPanelChanged?.Invoke(_detailPanel);
    }

    // =========================================================================
    // 채수 의뢰 상세 패널 (날짜 선택 시 Show2)
    // =========================================================================
    private Panel BuildSamplePanel(string date)
    {
        var rows = WasteSampleService.GetByDate(date);

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(4) };

        // 헤더
        DateTime.TryParse(date, out var d);
        root.Children.Add(new TextBlock
        {
            Text = $"🗓  채수일  {d:yyyy-MM-dd} ({DayOfWeekKr(d)})",
            FontWeight = FontWeight.SemiBold, FontFamily = Font,
            Foreground = Brush.Parse("#e0e0e0"),
        }.BindLG());
        root.Children.Add(new Border { Height = 1, Background = Brush.Parse("#333"), Margin = new Thickness(0,0,0,4) });

        // 그룹별 섹션
        foreach (var groupKey in new[] { "여수", "율촌", "세풍" })
        {
            var groupRows = rows.Where(r => r.구분 == groupKey).ToList();

            // 그룹 헤더
            var (gColor, gIcon) = groupKey switch
            {
                "여수" => ("#88aacc", "🌊"),
                "율촌" => ("#aaccaa", "🏗"),
                "세풍" => ("#ccaa88", "🏭"),
                _      => ("#aaaaaa", "📁"),
            };

            var groupHeader = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Margin = new Thickness(0, 4, 0, 4),
                Children =
                {
                    new TextBlock { Text = gIcon, VerticalAlignment = VerticalAlignment.Center }.BindLG(),
                    new TextBlock
                    {
                        Text = groupKey, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                        Foreground = Brush.Parse(gColor), VerticalAlignment = VerticalAlignment.Center,
                    }.BindMD(),
                    new TextBlock
                    {
                        Text = $"({groupRows.Count}건)", FontFamily = Font,
                        Foreground = Brush.Parse("#666"), VerticalAlignment = VerticalAlignment.Center,
                    }.BindSM()
                }
            };
            root.Children.Add(groupHeader);

            // 행 목록
            var listPanel = new StackPanel { Spacing = 2 };
            foreach (var row in groupRows)
                listPanel.Children.Add(BuildSampleRow(row, date, groupKey));
            root.Children.Add(listPanel);

            // + 업체 추가 버튼
            var gk = groupKey;
            var addBtn = new Button
            {
                Content = $"+ {groupKey} 업체 추가",
                Height = 24, FontSize = 10, FontFamily = Font,
                Background = Brush.Parse("#1a2a1a"), Foreground = Brush.Parse("#88cc88"),
                BorderThickness = new Thickness(1), BorderBrush = Brush.Parse("#2a4a2a"),
                CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 0),
                Margin = new Thickness(0, 2, 0, 8),
            };
            addBtn.Click += (_, _) => ShowAddCompanyDialog(date, gk);
            root.Children.Add(addBtn);
        }

        return root;
    }

    // =========================================================================
    // 개별 행 (S/N + 업체명 + ↑↓ 삭제)
    // =========================================================================
    private Grid BuildSampleRow(WasteSample s, string date, string groupKey)
    {
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("90,*,26,26,26"),
            Background = Brush.Parse(s.순서 % 2 == 1 ? "#1a1a28" : "#1e1e30"),
            Margin = new Thickness(0, 1),
        };

        // S/N
        row.Children.Add(new TextBlock
        {
            Text = s.SN, FontFamily = Font,
            Foreground = Brush.Parse("#aaccff"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 4),
            [Grid.ColumnProperty] = 0,
        }.BindBase());

        // 업체명
        row.Children.Add(new TextBlock
        {
            Text = s.업체명, FontFamily = Font,
            Foreground = Brush.Parse("#e0e0e0"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 4),
            [Grid.ColumnProperty] = 1,
        }.BindMD());

        // ↑
        var btnUp = MakeIconBtn("▲", "#555");
        btnUp.SetValue(Grid.ColumnProperty, 2);
        btnUp.Click += (_, _) =>
        {
            WasteSampleService.MoveUp(s.Id, date, groupKey);
            RefreshDetail();
        };

        // ↓
        var btnDown = MakeIconBtn("▼", "#555");
        btnDown.SetValue(Grid.ColumnProperty, 3);
        btnDown.Click += (_, _) =>
        {
            WasteSampleService.MoveDown(s.Id, date, groupKey);
            RefreshDetail();
        };

        // 삭제
        var btnDel = MakeIconBtn("✕", "#883333");
        btnDel.SetValue(Grid.ColumnProperty, 4);
        btnDel.Click += (_, _) =>
        {
            WasteSampleService.Delete(s.Id);
            RefreshDetail();
        };

        row.Children.Add(btnUp);
        row.Children.Add(btnDown);
        row.Children.Add(btnDel);
        return row;
    }

    private static Button MakeIconBtn(string text, string fg) => new()
    {
        Content = text, FontSize = 9,
        Background = Brushes.Transparent, Foreground = Brush.Parse(fg),
        BorderThickness = new Thickness(0), Padding = new Thickness(2),
        VerticalAlignment = VerticalAlignment.Center,
        HorizontalAlignment = HorizontalAlignment.Center,
    };

    // =========================================================================
    // 업체 추가 다이얼로그 (AutoCompleteBox)
    // =========================================================================
    private void ShowAddCompanyDialog(string date, string groupKey)
    {
        var companies = WasteCompanyService.GetAllItems()
            .Where(c => GetGroupLabel(c) == (groupKey == "율촌" || groupKey == "세풍" ? "율촌" : "여수"))
            .Select(c => c.업체명)
            .OrderBy(n => n)
            .ToList();

        var dlg = new Window
        {
            Title = $"{groupKey} 업체 추가",
            Width = 400, Height = 250,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Brush.Parse("#2a2a3a"),
        };

        // 날짜 선택 (기본값: 현재 선택된 날짜)
        DateTime.TryParse(date, out var parsedDate);
        var dp = new CalendarDatePicker
        {
            SelectedDate = parsedDate == DateTime.MinValue
                ? DateTime.Today
                : parsedDate,
            FontFamily = Font, FontSize = 12,
        };

        var acb = new AutoCompleteBox
        {
            ItemsSource = companies,
            Watermark = "업체명 입력 또는 선택",
            FilterMode = AutoCompleteFilterMode.Contains,
            Height = 32, FontSize = 12, FontFamily = Font,
        };

        var btnOk = new Button
        {
            Content = "추가", Width = 80, Height = 30,
            Background = Brush.Parse("#1a3a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 12,
        };
        var btnCancel = new Button
        {
            Content = "취소", Width = 80, Height = 30,
            Background = Brush.Parse("#333"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 12,
        };

        btnOk.Click += (_, _) =>
        {
            var name = acb.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(name)) return;
            if (dp.SelectedDate == null) return;

            string selectedDate = dp.SelectedDate.Value.ToString("yyyy-MM-dd");

            var comp = WasteCompanyService.GetAllItems()
                .FirstOrDefault(c => c.업체명 == name);
            WasteSampleService.Insert(selectedDate, groupKey, name, comp?.관리번호.ToString() ?? "");

            EnsureDateInTree(selectedDate);
            _selectedDate = selectedDate;
            dlg.Close();
            RefreshDetail();
        };
        btnCancel.Click += (_, _) => dlg.Close();

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(16), Spacing = 10,
            Children =
            {
                new TextBlock { Text = "채수일", FontSize = 11, FontFamily = Font, Foreground = Brush.Parse("#888") },
                dp,
                new TextBlock { Text = "업체명", FontSize = 11, FontFamily = Font, Foreground = Brush.Parse("#888") },
                acb,
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8,
                    Children = { btnOk, btnCancel }
                }
            }
        };

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) dlg.ShowDialog(owner);
        else dlg.Show();
    }

    // =========================================================================
    // 신규 날짜 추가 패널
    // =========================================================================
    private Panel BuildNewDatePanel()
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };

        root.Children.Add(new TextBlock
        {
            Text = "채수 날짜 신규 추가",
            FontSize = 15, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#e0e0e0"),
        });
        root.Children.Add(new Border { Height = 1, Background = Brush.Parse("#333") });

        var dp = new CalendarDatePicker
        {
            SelectedDate = DateTime.Today,
            FontFamily = Font, FontSize = 12,
        };

        var btnCreate = new Button
        {
            Content = "날짜 생성",
            Height = 30, FontSize = 12, FontFamily = Font,
            Background = Brush.Parse("#1a3a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(16, 0),
        };
        btnCreate.Click += (_, _) =>
        {
            if (dp.SelectedDate == null) return;
            string date = dp.SelectedDate.Value.ToString("yyyy-MM-dd");
            _selectedDate = date;
            EnsureDateInTree(date);
            RefreshDetail();
        };

        root.Children.Add(new TextBlock { Text = "채수일 선택", FontSize = 11, FontFamily = Font, Foreground = Brush.Parse("#888") });
        root.Children.Add(dp);
        root.Children.Add(btnCreate);
        return root;
    }

    // =========================================================================
    // 트리뷰에 날짜 노드 추가 (없을 때만)
    // =========================================================================
    private void EnsureDateInTree(string date)
    {
        string ym = date[..7];

        // 월 노드 찾기
        TreeViewItem? monthNode = DateTreeView.Items
            .OfType<TreeViewItem>()
            .FirstOrDefault(m => m.Items.OfType<TreeViewItem>().Any(d => (string?)d.Tag == date ||
                                  (d.Tag == null && m.Items.OfType<TreeViewItem>().Any(x => (x.Tag as string)?[..7] == ym))));

        // 없으면 월 노드 새로 생성
        if (monthNode == null)
        {
            monthNode = MakeMonthNode(ym, 0);
            monthNode.IsExpanded = true;
            DateTreeView.Items.Insert(0, monthNode);
        }

        // 날짜 노드가 이미 있으면 스킵
        bool exists = monthNode.Items.OfType<TreeViewItem>().Any(d => (string?)d.Tag == date);
        if (exists) return;

        var newSamples = WasteSampleService.GetByDate(date);
        monthNode.Items.Insert(0, MakeDateNode(date, newSamples));
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private static string GetGroupLabel(WasteCompany c)
    {
        var pname = c.프로젝트명 ?? "";
        if (pname.Contains("여수")) return "여수";
        if (pname.Contains("율촌")) return "율촌";
        return "기타";
    }

    // =========================================================================
    // Show4: 폐수배출업소 트리 패널
    // =========================================================================
    private Control BuildCompanyTreePanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };

        // Row 0: 헤더 (타이틀 + 그룹 탭 버튼)
        var headerSp = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 4,
            VerticalAlignment = VerticalAlignment.Center,
        };
        headerSp.Children.Add(new TextBlock
        {
            Text = "폐수배출업소", FontSize = 11, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 6, 0),
        });
        _btnCo여수 = MakeCompanyGroupBtn("🌊 여수", "여수");
        _btnCo율촌 = MakeCompanyGroupBtn("🏗 율촌", "율촌");
        _btnCo세풍 = MakeCompanyGroupBtn("🏭 세풍", "세풍");
        headerSp.Children.Add(_btnCo여수);
        headerSp.Children.Add(_btnCo율촌);
        headerSp.Children.Add(_btnCo세풍);

        root.Children.Add(new Border
        {
            Background = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(8, 4),
            Child = headerSp,
        });
        UpdateCompanyGroupTabStyles();

        // Row 1: 검색창
        _companySearchBox = new TextBox
        {
            Watermark = "🔍 업체명 검색",
            FontFamily = Font, FontSize = 11,
            Height = 28,
            Background = Brush.Parse("#252535"),
            Foreground = Brush.Parse("#cccccc"),
            BorderBrush = Brush.Parse("#333355"),
            BorderThickness = new Thickness(0, 0, 0, 1),
            CornerRadius = new CornerRadius(0),
            Padding = new Thickness(8, 4),
        };
        _companySearchBox.TextChanged += (_, _) => LoadCompanyTree();
        Grid.SetRow(_companySearchBox, 1);
        root.Children.Add(_companySearchBox);

        // Row 2: 상태
        _companyTreeStatus = new TextBlock
        {
            FontSize = 10, FontFamily = Font,
            Foreground = Brush.Parse("#666666"),
            Margin = new Thickness(8, 3),
        };
        Grid.SetRow(_companyTreeStatus, 2);
        root.Children.Add(_companyTreeStatus);

        // Row 3: 트리뷰
        _companyTreeView = new TreeView { Margin = new Thickness(5) };
        _companyTreeView.SelectionChanged += CompanyTree_SelectionChanged;
        Grid.SetRow(_companyTreeView, 3);
        root.Children.Add(_companyTreeView);

        LoadCompanyTree();
        return root;
    }

    private static readonly Dictionary<string, (string Bg, string Fg, string Border)> _groupTabColors = new()
    {
        ["여수"] = ("#1e3a5a", "#88aacc", "#336699"),
        ["율촌"] = ("#1a3a1a", "#aaccaa", "#336633"),
        ["세풍"] = ("#3a2a1a", "#ccaa88", "#996633"),
    };

    private Button MakeCompanyGroupBtn(string content, string tag)
    {
        var btn = new Button
        {
            Content = content, Tag = tag,
            FontSize = 10, FontFamily = Font,
            Padding = new Thickness(8, 3),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
        };
        btn.Click += (_, _) =>
        {
            _companyActiveGroup = tag;
            UpdateCompanyGroupTabStyles();
            LoadCompanyTree();
        };
        return btn;
    }

    private void UpdateCompanyGroupTabStyles()
    {
        foreach (var (btn, tag) in new[] { (_btnCo여수, "여수"), (_btnCo율촌, "율촌"), (_btnCo세풍, "세풍") })
        {
            if (btn == null) continue;
            bool active = tag == _companyActiveGroup;
            if (active && _groupTabColors.TryGetValue(tag, out var c))
            {
                btn.Background  = Brush.Parse(c.Bg);
                btn.Foreground  = Brush.Parse(c.Fg);
                btn.BorderBrush = Brush.Parse(c.Border);
            }
            else
            {
                btn.Background  = Brush.Parse("#222222");
                btn.Foreground  = Brush.Parse("#666666");
                btn.BorderBrush = Brush.Parse("#444444");
            }
        }
    }

    public void LoadCompanyTree()
    {
        if (_companyTreeView == null) return;
        _companyTreeView.Items.Clear();

        var search = _companySearchBox?.Text?.Trim() ?? "";

        try
        {
            var companies = WasteCompanyService.GetAllItems()
                .Where(c => CompanyGroup(c) == _companyActiveGroup)
                .Where(c => string.IsNullOrEmpty(search) ||
                            c.업체명.Contains(search) || c.약칭.Contains(search))
                .OrderBy(c => c.관리번호)
                .ToList();

            foreach (var c in companies)
                _companyTreeView.Items.Add(MakeCompanyLeafNode(c));

            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = Brush.Parse("#666666");
                string statusText = $"{_companyActiveGroup} {companies.Count}개";
                if (!string.IsNullOrEmpty(search)) statusText += $"  (검색: \"{search}\")";
                _companyTreeStatus.Text = statusText + " — 클릭하면 목록에 추가";
            }
        }
        catch (Exception ex)
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = Brush.Parse("#ee4444");
                _companyTreeStatus.Text = $"오류: {ex.Message}";
            }
        }
    }

    private static string CompanyGroup(WasteCompany c)
    {
        var pname = c.프로젝트명 ?? "";
        if (pname.Contains("여수")) return "여수";
        if (pname.Contains("세풍")) return "세풍";
        if (pname.Contains("율촌")) return "율촌";
        return "기타";
    }

    private static TreeViewItem MakeCompanyGroupNode(string name, int count)
    {
        var (icon, color) = name switch
        {
            "여수" => ("🌊", "#88aacc"),
            "율촌" => ("🏗", "#aaccaa"),
            "세풍" => ("🏭", "#ccaa88"),
            _      => ("📁", "#aaaaaa"),
        };
        return new TreeViewItem
        {
            IsExpanded = true,
            Header = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock { Text = icon, FontSize = 13, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock
                    {
                        Text = name, FontSize = 13, FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = Brush.Parse(color),
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new Border
                    {
                        Background = Brush.Parse("#2a2a3a"), CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = count.ToString(), FontSize = 9,
                            FontFamily = Font, Foreground = Brush.Parse("#888"),
                        }
                    }
                }
            }
        };
    }

    private static TreeViewItem MakeCompanyLeafNode(WasteCompany c)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };

        if (!string.IsNullOrWhiteSpace(c.약칭))
        {
            var (bg, fg, bd) = GetChosungBadgeColor(c.약칭);
            sp.Children.Add(new Border
            {
                Background        = Brush.Parse(bg),
                BorderBrush       = Brush.Parse(bd),
                BorderThickness   = new Thickness(1),
                CornerRadius      = new CornerRadius(8),
                Padding           = new Thickness(5, 1),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = c.약칭, FontSize = 9, FontFamily = Font,
                    Foreground = Brush.Parse(fg),
                },
            });
        }

        sp.Children.Add(new TextBlock
        {
            Text = c.업체명, FontSize = 12, FontFamily = Font,
            Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text = c.관리번호, FontSize = 9, FontFamily = Font,
            Foreground = Brush.Parse("#666"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        return new TreeViewItem { Tag = c, Header = sp };
    }

    private static readonly char[] _chosungs =
        { 'ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ' };

    private static (string Bg, string Fg, string Bd) GetChosungBadgeColor(string 약칭)
    {
        if (string.IsNullOrEmpty(약칭)) return ("#2a2a3a", "#888888", "#444444");
        char first = 약칭[0];
        char cho = (first >= 0xAC00 && first <= 0xD7A3)
            ? _chosungs[(first - 0xAC00) / (21 * 28)]
            : first;
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

    private void CompanyTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0) return;
        if (e.AddedItems[0] is not TreeViewItem tvi || tvi.Tag is not WasteCompany company) return;

        // 날짜 미선택
        if (string.IsNullOrEmpty(_selectedDate))
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = Brush.Parse("#ee8844");
                _companyTreeStatus.Text = "왼쪽에서 날짜를 먼저 선택하세요";
            }
            _companyTreeView!.SelectedItem = null;
            return;
        }

        // 중복 체크
        var existing = WasteSampleService.GetByDate(_selectedDate);
        if (existing.Any(s => s.업체명 == company.업체명))
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = Brush.Parse("#ee8844");
                _companyTreeStatus.Text = $"이미 추가됨: {company.업체명}";
            }
            _companyTreeView!.SelectedItem = null;
            return;
        }

        // 삽입
        string groupKey = CompanyGroup(company);
        if (groupKey == "기타") groupKey = "여수";
        WasteSampleService.Insert(_selectedDate, groupKey, company.업체명, company.관리번호);
        EnsureDateInTree(_selectedDate);

        if (_companyTreeStatus != null)
        {
            _companyTreeStatus.Foreground = Brush.Parse("#88ee88");
            _companyTreeStatus.Text = $"추가됨: {company.업체명}";
        }

        _companyTreeView!.SelectedItem = null;
        RefreshDetail();
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [WasteSampleList] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("Logs/WasteSampleDebug.log", line + Environment.NewLine); } catch { }
    }
}
