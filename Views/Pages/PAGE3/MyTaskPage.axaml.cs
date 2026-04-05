using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE3;
using ETA.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE3;

public partial class MyTaskPage : UserControl
{
    // ── 외부 이벤트 ──────────────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;

    // ── 상태 ─────────────────────────────────────────────────────────────────
    private string _activeCategory = "수질분석센터";
    private DateTime _selectedDate   = DateTime.Today;
    private string   _employeeId     = "";
    private string   _employeeName   = "";
    private HashSet<DateTime> _analysisRequestDates = new();
    private HashSet<DateTime> _wasteRequestDates = new();
    private Calendar? _calendar;

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // 카테고리별 활성 버튼 스타일
    private static readonly Dictionary<string, (string Bg, string Fg, string Bd)> CatStyle = new()
    {
        ["수질분석센터"] = ("#1a2a4a", "#88aaee", "#336699"),
        ["비용부담금"]   = ("#2a2a1a", "#ccbb88", "#887733"),
        ["처리시설"]     = ("#2a1a3a", "#bb99cc", "#663399"),
        ["의뢰목록"]     = ("#1a3a2a", "#88ccaa", "#339966"),
    };

    private readonly Dictionary<string, Button> _catBtns = new();

    public MyTaskPage()
    {
        InitializeComponent();

        // 날짜 텍스트 초기화
        UpdateDateText();

        // 인라인 캘린더 직접 생성 (날짜별 색상 마킹)
        _calendar = new Calendar
        {
            SelectionMode = CalendarSelectionMode.SingleDate,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            DisplayDate = DateTime.Today,
            Focusable = true,
        };

        var calBorder = new Border
        {
            IsVisible = false,
            Background = AppTheme.BgInput,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
            Padding = new Avalonia.Thickness(4),
            Margin = new Avalonia.Thickness(0, 4, 0, 0),
            Focusable = true,
            Child = _calendar,
        };

        // 날짜 클릭 → 선택 처리
        _calendar.SelectedDatesChanged += (_, _) =>
        {
            if (_calendar.SelectedDate.HasValue)
            {
                var date = _calendar.SelectedDate.Value;
                calBorder.IsVisible = false;
                _selectedDate = date;
                UpdateDateText();
                RefreshTree();
                _calendar.SelectedDates.Clear();
            }
        };

        // ESC → expanded 상태면 날짜 뷰로 복귀, 아니면 캘린더 닫기
        _calendar.KeyDown += (_, e) =>
        {
            if (e.Key == Avalonia.Input.Key.Escape && calBorder.IsVisible)
            {
                // 월/년 선택 뷰(expanded)면 DisplayMode를 Month로 복귀
                if (_calendar.DisplayMode != CalendarMode.Month)
                {
                    _calendar.DisplayMode = CalendarMode.Month;
                    Dispatcher.UIThread.Post(StyleCalendarDayButtons, DispatcherPriority.Loaded);
                }
                else
                {
                    calBorder.IsVisible = false;
                }
                e.Handled = true;
            }
        };

        // 월 변경 시 마킹 갱신
        _calendar.DisplayDateChanged += (_, _) =>
            Dispatcher.UIThread.Post(StyleCalendarDayButtons, DispatcherPriority.Loaded);

        // 토글 버튼
        btnCalToggle.Click += (_, _) =>
        {
            calBorder.IsVisible = !calBorder.IsVisible;
            if (calBorder.IsVisible)
            {
                _calendar.DisplayDate = _selectedDate;
                LoadDateSets();
                _calendar.Focus();
            }
        };

        calendarHost.Child = calBorder;

        _catBtns["수질분석센터"] = BtnMeokdae;
        _catBtns["비용부담금"]   = BtnWasteCost;
        _catBtns["처리시설"]     = BtnFacility;
        _catBtns["의뢰목록"]     = BtnRequestList;
    }

    private void UpdateDateText()
    {
        string dayKr = _selectedDate.DayOfWeek switch
        {
            DayOfWeek.Monday    => "월",
            DayOfWeek.Tuesday   => "화",
            DayOfWeek.Wednesday => "수",
            DayOfWeek.Thursday  => "목",
            DayOfWeek.Friday    => "금",
            DayOfWeek.Saturday  => "토",
            _                   => "일",
        };
        txtDate.Text = $"{_selectedDate:yyyy-MM-dd} ({dayKr})";
    }

    private void LoadDateSets()
    {
        _analysisRequestDates = new HashSet<DateTime>(MyTaskService.GetAnalysisRequestDates());
        _wasteRequestDates = new HashSet<DateTime>(MyTaskService.GetWasteRequestDates());
        Dispatcher.UIThread.Post(StyleCalendarDayButtons, DispatcherPriority.Loaded);
    }

    // 파란색: 분석의뢰및결과, 붉은색: 폐수의뢰및결과, 보라색: 둘 다
    private static readonly IBrush BrushAnalysis = new SolidColorBrush(Color.Parse("#B366aaee")); // 파란 70%
    private static readonly IBrush BrushWaste    = new SolidColorBrush(Color.Parse("#B3ee6666")); // 붉은 70%
    private static readonly IBrush BrushBoth     = new SolidColorBrush(Color.Parse("#B3aa66ee")); // 보라 70%

    private void StyleCalendarDayButtons()
    {
        if (_calendar == null) return;

        foreach (var btn in _calendar.GetVisualDescendants().OfType<CalendarDayButton>())
        {
            if (btn.DataContext is DateTime dt)
            {
                var date = dt.Date;
                bool isAnalysis = _analysisRequestDates.Contains(date);
                bool isWaste = _wasteRequestDates.Contains(date);

                if (isAnalysis && isWaste)
                    btn.Background = BrushBoth;
                else if (isAnalysis)
                    btn.Background = BrushAnalysis;
                else if (isWaste)
                    btn.Background = BrushWaste;
                else
                    btn.Background = Brushes.Transparent;
            }
        }
    }

    // =========================================================================
    // 외부 호출 — MainPage에서 진입 시
    // =========================================================================
    public void LoadData()
    {
        _employeeId   = MainPage.CurrentEmployeeId;
        _employeeName = MyTaskService.GetEmployeeName(_employeeId);
        RefreshTree();
    }

    // =========================================================================
    // 날짜 변경 — InlineCalendarHelper 콜백으로 처리
    // =========================================================================

    // =========================================================================
    // 카테고리 버튼 클릭
    // =========================================================================
    private void CategoryBtn_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        _activeCategory = btn.Tag as string ?? "수질분석센터";
        UpdateCategoryButtons();
        RefreshTree();
        DetailPanelChanged?.Invoke(null);
    }

    private void UpdateCategoryButtons()
    {
        foreach (var (cat, btn) in _catBtns)
        {
            if (cat == _activeCategory && CatStyle.TryGetValue(cat, out var s))
            {
                btn.Background  = new SolidColorBrush(Color.Parse(s.Bg));
                btn.Foreground  = new SolidColorBrush(Color.Parse(s.Fg));
                btn.BorderBrush = new SolidColorBrush(Color.Parse(s.Bd));
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
    // 트리 새로고침
    // =========================================================================
    private void RefreshTree()
    {
        TaskTreeView.Items.Clear();

        if (string.IsNullOrEmpty(_employeeId)) return;

        string dateStr = _selectedDate.ToString("yyyy-MM-dd");

        switch (_activeCategory)
        {
            case "수질분석센터": LoadMeokdaeTree(dateStr);       break;
            case "비용부담금":   LoadAllWasteTree(dateStr);       break;
            case "처리시설":     LoadFacilityTree(dateStr);       break;
            case "의뢰목록":     LoadRequestListTree(dateStr);    break;
        }
    }

    // ── 수질분석센터 트리 ─────────────────────────────────────────────────────────────
    private void LoadMeokdaeTree(string dateStr)
    {
        var items = MyTaskService.GetMyAnalysisItems(_employeeId, _selectedDate);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("할당된 분석항목 없음"));
            return;
        }

        foreach (var (fullName, shortName) in items)
        {
            var node = MakeAnalyteNode(fullName, shortName);
            TaskTreeView.Items.Add(node);
        }
    }

    private TreeViewItem MakeAnalyteNode(string fullName, string shortName)
    {
        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock
        {
            Text = "🧪", FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = string.IsNullOrEmpty(shortName) ? fullName : shortName,
            FontFamily = Font, FontSize = AppTheme.FontLG,
            VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(shortName) && shortName != fullName)
        {
            header.Children.Add(new TextBlock
            {
                Text = $"({fullName})", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
        }

        var item = new TreeViewItem { Header = header };
        item.Tag = new AnalyteTag(fullName, shortName);
        return item;
    }

    private record AnalyteTag(string FullName, string ShortName);

    // ── 비용부담금 트리 (여수+율촌+세풍 통합) ─────────────────────────────────
    private void LoadAllWasteTree(string dateStr)
    {
        var allItems = new List<WasteRequestItem>();
        foreach (var 구분 in new[] { "여수", "율촌", "세풍" })
            allItems.AddRange(MyTaskService.GetWasteItems(dateStr, 구분));

        if (allItems.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("비용부담금 의뢰 없음"));
            return;
        }

        // 구분별 그룹 → 업체별 서브그룹
        var byGubun = allItems.GroupBy(i => i.구분)
            .OrderBy(g => g.Key == "여수" ? 0 : g.Key == "율촌" ? 1 : 2);

        foreach (var gubunGrp in byGubun)
        {
            var (icon, color) = gubunGrp.Key switch
            {
                "여수" => ("🌊", "#88aacc"),
                "율촌" => ("🏗", "#aaccaa"),
                _      => ("🏭", "#ccaa88"),
            };

            // 구분 헤더 노드
            var gubunHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            gubunHeader.Children.Add(new TextBlock { Text = icon, FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = gubunGrp.Key, FontFamily = Font, FontSize = AppTheme.FontLG,
                Foreground = new SolidColorBrush(Color.Parse(color)),
                FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
            });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = $"{gubunGrp.Count()}건", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
            });

            var gubunNode = new TreeViewItem { Header = gubunHeader, IsExpanded = true };

            // 업체별 그룹
            var byCompany = gubunGrp.GroupBy(i => i.업체명).OrderBy(g => g.Key);
            foreach (var compGrp in byCompany)
            {
                var companyNode = MakeCompanyNode(compGrp.Key, compGrp.Count(), gubunGrp.Key);
                foreach (var wi in compGrp)
                    companyNode.Items.Add(MakeWasteItemNode(wi));
                gubunNode.Items.Add(companyNode);
            }

            TaskTreeView.Items.Add(gubunNode);
        }
    }

    private void LoadWasteTree(string dateStr, string 구분)
    {
        var items = MyTaskService.GetWasteItems(dateStr, 구분);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode($"{구분} 의뢰 없음"));
            return;
        }

        // 업체별 그룹
        var groups = items.GroupBy(i => i.업체명).OrderBy(g => g.Key);
        foreach (var grp in groups)
        {
            var companyNode = MakeCompanyNode(grp.Key, grp.Count(), 구분);
            foreach (var wi in grp)
                companyNode.Items.Add(MakeWasteItemNode(wi));
            TaskTreeView.Items.Add(companyNode);
        }
    }

    private TreeViewItem MakeCompanyNode(string company, int count, string 구분)
    {
        var (icon, color) = 구분 switch
        {
            "여수" => ("🌊", "#88aacc"),
            "율촌" => ("🏗", "#aaccaa"),
            "세풍" => ("🏭", "#ccaa88"),
            _     => ("📍", "#aaaaaa"),
        };

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock
        {
            Text = icon, FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = company, FontFamily = Font, FontSize = AppTheme.FontLG,
            Foreground = new SolidColorBrush(Color.Parse(color)),
            VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = $"{count}건", FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
        });

        return new TreeViewItem { Header = header };
    }

    private TreeViewItem MakeWasteItemNode(WasteRequestItem wi)
    {
        var (statusColor, statusIcon) = wi.상태 switch
        {
            "완료" => ("#3388dd", "✅"),
            "담음" => ("#dd8833", "🔄"),
            _     => ("#dd3333", "⬜"),
        };

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock
        {
            Text = statusIcon, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = wi.항목, FontFamily = Font, FontSize = AppTheme.FontMD,
            Foreground = new SolidColorBrush(Color.Parse(statusColor)),
            VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(wi.배정자))
        {
            header.Children.Add(new TextBlock
            {
                Text = wi.배정자, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
        }

        var item = new TreeViewItem { Header = header };
        item.Tag = wi;
        return item;
    }

    // ── 의뢰목록 트리 ──────────────────────────────────────────────────────────
    private void LoadRequestListTree(string dateStr)
    {
        var items = MyTaskService.GetRequestListItems(dateStr);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("해당일 의뢰 없음"));
            return;
        }

        // 구분별 그룹
        var byGubun = items.GroupBy(i => i.구분)
            .OrderBy(g => g.Key == "여수" ? 0 : g.Key == "율촌" ? 1 : 2);

        foreach (var grp in byGubun)
        {
            var (icon, color) = grp.Key switch
            {
                "여수" => ("🌊", "#88aacc"),
                "율촌" => ("🏗", "#aaccaa"),
                _      => ("🏭", "#ccaa88"),
            };

            var gubunHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            gubunHeader.Children.Add(new TextBlock { Text = icon, FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = grp.Key, FontFamily = Font, FontSize = AppTheme.FontLG,
                Foreground = new SolidColorBrush(Color.Parse(color)),
                FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
            });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = $"{grp.Count()}건", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
            });

            var gubunNode = new TreeViewItem { Header = gubunHeader, IsExpanded = true };

            foreach (var ri in grp)
                gubunNode.Items.Add(MakeRequestListNode(ri));

            TaskTreeView.Items.Add(gubunNode);
        }
    }

    private TreeViewItem MakeRequestListNode(MyTaskService.RequestListItem ri)
    {
        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock
        {
            Text = ri.HasResult ? "🟢" : "🔴",
            FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = ri.업체명, FontFamily = Font, FontSize = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(ri.관리번호))
        {
            header.Children.Add(new TextBlock
            {
                Text = ri.관리번호, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
        }
        if (!string.IsNullOrEmpty(ri.확인자))
        {
            header.Children.Add(new TextBlock
            {
                Text = ri.확인자, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = new SolidColorBrush(Color.Parse("#88ccaa")),
                VerticalAlignment = VerticalAlignment.Center
            });
        }

        var item = new TreeViewItem { Header = header };
        item.Tag = ri;
        return item;
    }

    // ── 처리시설 트리 ──────────────────────────────────────────────────────────
    private void LoadFacilityTree(string dateStr)
    {
        var items = MyTaskService.GetFacilityItems(dateStr);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("처리시설 작업 없음"));
            return;
        }

        foreach (var fi in items)
        {
            var node = MakeFacilityNode(fi);
            TaskTreeView.Items.Add(node);
        }
    }

    private TreeViewItem MakeFacilityNode(FacilityWorkItem fi)
    {
        var (statusColor, statusIcon) = fi.상태 switch
        {
            "완료" => ("#3388dd", "✅"),
            "담음" => ("#dd8833", "🔄"),
            _     => ("#dd3333", "⬜"),
        };

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock
        {
            Text = statusIcon, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = fi.시설명, FontFamily = Font, FontSize = AppTheme.FontLG,
            VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = fi.시료명, FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
        });

        var item = new TreeViewItem { Header = header };
        item.Tag = fi;
        return item;
    }


    // ── 빈 노드 ───────────────────────────────────────────────────────────────
    private TreeViewItem MakeEmptyNode(string text)
    {
        return new TreeViewItem
        {
            Header = new TextBlock
            {
                Text = text, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppRes("FgMuted")
            }
        };
    }

    // =========================================================================
    // 트리 선택 → Show2 상세 패널
    // =========================================================================
    private void TaskTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (TaskTreeView.SelectedItem is not TreeViewItem sel) return;

        Control? panel = sel.Tag switch
        {
            AnalyteTag tag                    => BuildAnalyteDetail(tag),
            WasteRequestItem wi               => BuildWasteItemDetail(wi),
            FacilityWorkItem fi               => BuildFacilityDetail(fi),
            MyTaskService.RequestListItem ri  => BuildRequestListDetail(ri),
            _                                 => null,
        };

        DetailPanelChanged?.Invoke(panel);
    }

    // ── 분석항목 상세 ────────────────────────────────────────────────────────
    private Control BuildAnalyteDetail(AnalyteTag tag)
    {
        string dateStr = _selectedDate.ToString("yyyy-MM-dd");
        var samples = MyTaskService.GetSamplesForAnalyte(dateStr, tag.FullName);

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        // 제목
        root.Children.Add(new TextBlock
        {
            Text = $"🧪  {(string.IsNullOrEmpty(tag.ShortName) ? tag.FullName : tag.ShortName)}",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });
        root.Children.Add(new TextBlock
        {
            Text = tag.FullName, FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppRes("FgMuted")
        });
        root.Children.Add(new TextBlock
        {
            Text = $"채취일: {dateStr}  /  총 {samples.Count}건",
            FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted"),
            Margin = new Thickness(0, 4, 0, 4)
        });

        if (samples.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "해당 날짜에 의뢰된 시료가 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted")
            });
        }
        else
        {
            foreach (var s in samples)
            {
                var card = new Border
                {
                    Background = AppRes("PanelInnerBg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 6),
                    Margin = new Thickness(0, 2)
                };
                var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
                row.Children.Add(new TextBlock
                {
                    Text = s.Company, FontFamily = Font, FontSize = AppTheme.FontMD,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("AppFg"),
                    Width = 80
                });
                row.Children.Add(new TextBlock
                {
                    Text = s.Sample, FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppRes("AppFg"), Width = 120
                });
                if (!string.IsNullOrEmpty(s.접수번호))
                    row.Children.Add(new TextBlock
                    {
                        Text = s.접수번호, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppRes("FgMuted")
                    });
                card.Child = row;
                root.Children.Add(card);
            }
        }

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
    }

    // ── 폐수 항목 상세 ───────────────────────────────────────────────────────
    private Control BuildWasteItemDetail(WasteRequestItem wi)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"🌊  {wi.업체명} — {wi.항목}",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted"), Width = 70 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = color != null ? new SolidColorBrush(Color.Parse(color)) : AppRes("AppFg")
            });
            root.Children.Add(sp);
        }

        AddRow("의뢰번호", wi.의뢰번호);
        AddRow("채취일", wi.채취일자);
        AddRow("구분", wi.구분);

        var statusColor = wi.상태 switch
        {
            "완료" => "#3388dd",
            "담음" => "#dd8833",
            _     => "#dd3333",
        };
        AddRow("상태", wi.상태, statusColor);

        if (!string.IsNullOrEmpty(wi.배정자))  AddRow("배정자", wi.배정자);
        if (!string.IsNullOrEmpty(wi.배정일시)) AddRow("배정일시", wi.배정일시);
        if (!string.IsNullOrEmpty(wi.완료일시)) AddRow("완료일시", wi.완료일시);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
    }

    // ── 의뢰목록 상세 ────────────────────────────────────────────────────────
    private Control BuildRequestListDetail(MyTaskService.RequestListItem ri)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"📋  {ri.업체명}",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            if (string.IsNullOrEmpty(value)) return;
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted"), Width = 70 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = color != null ? new SolidColorBrush(Color.Parse(color)) : AppRes("AppFg")
            });
            root.Children.Add(sp);
        }

        AddRow("SN", ri.SN);
        AddRow("구분", ri.구분);
        AddRow("채수일", ri.채수일);
        AddRow("관리번호", ri.관리번호);
        AddRow("확인자", ri.확인자, "#88ccaa");
        AddRow("비고", ri.비고);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
    }

    // ── 처리시설 상세 ────────────────────────────────────────────────────────
    private Control BuildFacilityDetail(FacilityWorkItem fi)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"⚙  {fi.시설명}",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted"), Width = 70 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = color != null ? new SolidColorBrush(Color.Parse(color)) : AppRes("AppFg")
            });
            root.Children.Add(sp);
        }

        AddRow("시료명", fi.시료명);
        AddRow("채취일", fi.채취일자);
        AddRow("항목", fi.항목목록);

        var statusColor = fi.상태 switch
        {
            "완료" => "#3388dd",
            "담음" => "#dd8833",
            _     => "#dd3333",
        };
        AddRow("상태", fi.상태, statusColor);

        if (!string.IsNullOrEmpty(fi.배정자))  AddRow("배정자", fi.배정자);
        if (!string.IsNullOrEmpty(fi.배정일시)) AddRow("배정일시", fi.배정일시);
        if (!string.IsNullOrEmpty(fi.완료일시)) AddRow("완료일시", fi.완료일시);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }
}
