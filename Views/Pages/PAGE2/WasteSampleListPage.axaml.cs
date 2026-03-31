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
    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;

    private string? _selectedDate;
    private Panel?  _detailPanel;

    // ── Show4: 업체 트리 ─────────────────────────────────────────────────────
    private Control?   _companyTreePanel;
    private TreeView?  _companyTreeView;
    private TextBlock? _companyTreeStatus;

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
                    monthNode.Items.Add(MakeDateNode(d));
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
                        Text = $"{d.Year}년 {d.Month}월",
                        FontSize = 12, FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new Border
                    {
                        Background = Brush.Parse("#2a2a3a"), CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5,1),
                        Child = new TextBlock
                        {
                            Text = count.ToString(), FontSize = 9,
                            FontFamily = Font, Foreground = Brush.Parse("#666"),
                        }
                    }
                }
            }
        };
    }

    private static TreeViewItem MakeDateNode(string dateStr)
    {
        DateTime.TryParse(dateStr, out var d);
        string label = d == DateTime.MinValue ? dateStr : $"{d.Month}/{d.Day} ({DayOfWeekKr(d)})";
        return new TreeViewItem
        {
            Tag    = dateStr,
            Header = new TextBlock
            {
                Text = label, FontSize = 12, FontFamily = Font,
                Foreground = Brush.Parse("#dddddd"),
            }
        };
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
    // + 날짜 추가
    // =========================================================================
    public void AddNewDate() => BtnNewDate_Click(null, null!);

    private void BtnNewDate_Click(object? sender, RoutedEventArgs e)
    {
        // 오늘 날짜로 신규 날짜 추가 패널 표시
        _selectedDate = null;
        DetailPanelChanged?.Invoke(BuildNewDatePanel());
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
            Text = $"채수일  {d:yyyy-MM-dd} ({DayOfWeekKr(d)})",
            FontSize = 15, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#e0e0e0"),
        });
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
                    new TextBlock { Text = gIcon, FontSize = 13, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock
                    {
                        Text = groupKey, FontSize = 12, FontFamily = Font,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = Brush.Parse(gColor),
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new TextBlock
                    {
                        Text = $"({groupRows.Count}건)", FontSize = 10, FontFamily = Font,
                        Foreground = Brush.Parse("#666"), VerticalAlignment = VerticalAlignment.Center,
                    }
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
            Text = s.SN, FontSize = 11, FontFamily = Font,
            Foreground = Brush.Parse("#aaccff"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 4),
            [Grid.ColumnProperty] = 0,
        });

        // 업체명
        row.Children.Add(new TextBlock
        {
            Text = s.업체명, FontSize = 12, FontFamily = Font,
            Foreground = Brush.Parse("#e0e0e0"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 4),
            [Grid.ColumnProperty] = 1,
        });

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
        var dp = new DatePicker
        {
            SelectedDate = parsedDate == DateTime.MinValue
                ? DateTimeOffset.Now
                : new DateTimeOffset(parsedDate),
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

            string selectedDate = dp.SelectedDate.Value.Date.ToString("yyyy-MM-dd");

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

        var dp = new DatePicker
        {
            SelectedDate = DateTimeOffset.Now,
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
            string date = dp.SelectedDate.Value.Date.ToString("yyyy-MM-dd");
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

        monthNode.Items.Insert(0, MakeDateNode(date));
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
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };

        root.Children.Add(new Border
        {
            Background = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text = "🏭  폐수배출업소",
                FontSize = 11, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
            }
        });

        _companyTreeStatus = new TextBlock
        {
            FontSize = 10, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = Brush.Parse("#666666"),
            Margin = new Thickness(8, 3),
        };
        Grid.SetRow(_companyTreeStatus, 1);
        root.Children.Add(_companyTreeStatus);

        _companyTreeView = new TreeView { Margin = new Thickness(5) };
        _companyTreeView.SelectionChanged += CompanyTree_SelectionChanged;
        Grid.SetRow(_companyTreeView, 2);
        root.Children.Add(_companyTreeView);

        LoadCompanyTree();
        return root;
    }

    public void LoadCompanyTree()
    {
        if (_companyTreeView == null) return;
        _companyTreeView.Items.Clear();

        try
        {
            var all = WasteCompanyService.GetAllItems();
            int total = 0;
            foreach (var g in new[] { "여수", "율촌", "세풍" })
            {
                var companies = all
                    .Where(c => CompanyGroup(c) == g)
                    .OrderBy(c => c.관리번호)
                    .ToList();
                if (companies.Count == 0) continue;

                var groupNode = MakeCompanyGroupNode(g, companies.Count);
                foreach (var c in companies)
                    groupNode.Items.Add(MakeCompanyLeafNode(c));
                groupNode.IsExpanded = true;
                _companyTreeView.Items.Add(groupNode);
                total += companies.Count;
            }
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = Brush.Parse("#666666");
                _companyTreeStatus.Text = $"전체 {total}개 — 클릭하면 목록에 추가";
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

    private static TreeViewItem MakeCompanyLeafNode(WasteCompany c) => new()
    {
        Tag = c,
        Header = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 4,
            Children =
            {
                new TextBlock
                {
                    Text = c.업체명, FontSize = 12, FontFamily = Font,
                    FontWeight = FontWeight.Regular,
                    Foreground = Brushes.WhiteSmoke,
                    VerticalAlignment = VerticalAlignment.Center,
                },
                new TextBlock
                {
                    Text = c.관리번호, FontSize = 9, FontFamily = Font,
                    FontWeight = FontWeight.Regular,
                    Foreground = Brush.Parse("#666"),
                    VerticalAlignment = VerticalAlignment.Center,
                },
            }
        }
    };

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
