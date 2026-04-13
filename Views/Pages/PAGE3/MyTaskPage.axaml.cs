using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ETA.Models;
using ETA.Services.Common;
using ETA.Services.SERVICE3;
using ETA.Views;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE3;

public partial class MyTaskPage : UserControl
{
    // ── 외부 이벤트 ──────────────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;   // Show2
    public event Action<Control?>? EditPanelChanged;     // Show3
    public event Action<Control?>? StatsPanelChanged;    // Show4

    // ── 상태 ─────────────────────────────────────────────────────────────────
    private string _activeCategory = "수질분석센터";
    private DateTime _selectedDate   = DateTime.Today;
    private string   _employeeId     = "";
    private string   _employeeName   = "";

    // 수질분석센터/비용부담금 — 선택 바스켓 (Show2)
    private readonly Dictionary<string, string> _basket = new();  // "category:id" → 표시레이블
    private readonly List<(int Id, CheckBox Cb)> _treeCheckboxes = new();

    // 수질분석센터 — 대리 분석 추가 항목
    private HashSet<string> _myExtraItems = new();
    private HashSet<string> _removedBaseItems = new();  // 기본 담당항목 중 임시 제거된 것
    private MyTaskService.SampleRow? _selectedSample;
    private List<MyTaskService.AnalyteAssignment> _currentAssignments = new();

    // Show4 항상표시용
    private List<MyTaskService.AnalyteAssignment> _allAnalytes = new();
    private HashSet<string>? _markedColumns = null;  // null=선택 없음

    // Show3 처리시설용 — 사용자 분장항목
    private List<MyTaskService.UserAssignedItem> _userAssigned = new();

    // Show4→Show3 대리분석 — 클릭으로 빌린 항목
    private readonly HashSet<string> _borrowedAnalytes = new();

    // Show2 노트 — 선택된 분석항목 (Excel 생성용)
    private readonly HashSet<string> _selectedAnalytesForNote = new();

    // 트리 재빌드 중 SelectionChanged 억제 플래그
    private bool _isRefreshing = false;

    // 캘린더 툴팁 캐시
    private Dictionary<DateTime, string> _calendarTooltips = new();
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
                DetailPanelChanged?.Invoke(BuildSelectedListPanel());
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

        // 월 변경 시 마킹 + 툴팁 갱신
        _calendar.DisplayDateChanged += (_, _) =>
        {
            LoadCalendarTooltipsForMonth(_calendar.DisplayDate);
            Dispatcher.UIThread.Post(StyleCalendarDayButtons, DispatcherPriority.Loaded);
        };

        // 토글 버튼
        btnCalToggle.Click += (_, _) =>
        {
            calBorder.IsVisible = !calBorder.IsVisible;
            if (calBorder.IsVisible)
            {
                _calendar.DisplayDate = _selectedDate;
                LoadDateSets();
                LoadCalendarTooltipsForMonth(_selectedDate);
                Dispatcher.UIThread.Post(StyleCalendarDayButtons, DispatcherPriority.Loaded);
                _calendar.Focus();
            }
        };

        calendarHost.Child = calBorder;

        _catBtns["수질분석센터"] = BtnMeokdae;
        _catBtns["비용부담금"]   = BtnWasteCost;
        _catBtns["처리시설"]     = BtnFacility;

        BtnSelectAll.Click += (_, _) =>
        {
            foreach (var (id, cb) in _treeCheckboxes)
                if (!cb.IsChecked.GetValueOrDefault()) cb.IsChecked = true;
        };

        BtnSendToShow2.Click += (_, _) =>
            DetailPanelChanged?.Invoke(BuildSelectedListPanel());
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
    private static readonly IBrush BrushAnalysis = new SolidColorBrush(Color.Parse("#664d7fb3")); // 스틸블루 40%
    private static readonly IBrush BrushWaste    = new SolidColorBrush(Color.Parse("#66b35555")); // 로즈브라운 40%
    private static readonly IBrush BrushBoth     = new SolidColorBrush(Color.Parse("#66755aaa")); // 슬레이트바이올릿 40%

    private void LoadCalendarTooltipsForMonth(DateTime displayDate)
    {
        var start = new DateTime(displayDate.Year, displayDate.Month, 1).AddDays(-7);
        var end   = start.AddDays(42);
        string s = start.ToString("yyyy-MM-dd");
        string e = end.ToString("yyyy-MM-dd");

        var meokData  = MyTaskService.GetSamplesByMonth(s, e);
        var wasteData = MyTaskService.GetWasteRequestsByMonth(s, e);

        _calendarTooltips.Clear();
        foreach (var date in meokData.Keys.Union(wasteData.Keys))
        {
            var sb = new StringBuilder();

            if (meokData.TryGetValue(date, out var meokList) && meokList.Count > 0)
            {
                sb.AppendLine($"수질분석센터 ({meokList.Count}건)");
                foreach (var (yakchim, sample) in meokList.Take(8))
                {
                    string prefix = string.IsNullOrEmpty(yakchim) ? "" : $"{yakchim} · ";
                    sb.AppendLine($"  {prefix}{sample}");
                }
                if (meokList.Count > 8)
                    sb.AppendLine($"  ... 외 {meokList.Count - 8}건");
            }

            if (wasteData.TryGetValue(date, out var wasteList) && wasteList.Count > 0)
            {
                if (sb.Length > 0) sb.AppendLine();
                sb.AppendLine($"비용부담금 ({wasteList.Count}건)");
                var byGubun = wasteList.GroupBy(x => x.구분).OrderBy(g => g.Key == "여수" ? 0 : g.Key == "율촌" ? 1 : 2);
                foreach (var grp in byGubun)
                {
                    sb.AppendLine($"  [{grp.Key}]");
                    foreach (var item in grp.Take(10))
                        sb.AppendLine($"    {item.업체명}");
                    if (grp.Count() > 10)
                        sb.AppendLine($"    ... 외 {grp.Count() - 10}건");
                }
            }

            _calendarTooltips[date] = sb.ToString().TrimEnd();
        }
    }

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

                if (_calendarTooltips.TryGetValue(date, out var tip))
                    ToolTip.SetTip(btn, tip);
                else
                    ToolTip.SetTip(btn, null);
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
        // RefreshTree 내부에서 DetailPanelChanged 호출하므로 별도 호출 불필요
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
        DetailPanelChanged?.Invoke(BuildSelectedListPanel());
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
        _isRefreshing = true;
        TaskTreeView.Items.Clear();
        _isRefreshing = false;

        if (string.IsNullOrEmpty(_employeeId))
        {
            DetailPanelChanged?.Invoke(BuildSelectedListPanel());
            return;
        }

        string dateStr = _selectedDate.ToString("yyyy-MM-dd");

        _treeCheckboxes.Clear();
        SelectAllBar.IsVisible = true;
        BtnSendToShow2.IsVisible = _activeCategory == "처리시설";

        switch (_activeCategory)
        {
            case "수질분석센터": LoadMeokdaeTree(dateStr);       break;
            case "비용부담금":   LoadAllWasteTree(dateStr);       break;
            case "처리시설":     LoadFacilityTree(dateStr);       break;
        }

        LoadShow4Analytes();  // 모든 카테고리에서 Show4 항상 유지

        // 카테고리 전환 후 Show2를 항상 바스켓 목록으로 복원
        DetailPanelChanged?.Invoke(BuildSelectedListPanel());
    }

    // ── 수질분석센터 트리 ─────────────────────────────────────────────────────────────
    private void LoadMeokdaeTree(string dateStr)
    {
        _selectedSample = null;
        _myExtraItems.Clear();
        _removedBaseItems.Clear();
        _currentAssignments.Clear();

        var samples = MyTaskService.GetSamplesForDate(dateStr);

        if (samples.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("해당 날짜에 의뢰된 시료 없음"));
            return;
        }

        // 약칭별 그룹
        var byCompany = samples.GroupBy(s => string.IsNullOrEmpty(s.약칭) ? "기타" : s.약칭)
                               .OrderBy(g => g.Key);

        foreach (var grp in byCompany)
        {
            var compHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            compHeader.Children.Add(new TextBlock { Text = "🏢", FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center });
            compHeader.Children.Add(new TextBlock
            {
                Text = grp.Key, FontFamily = Font, FontSize = AppTheme.FontLG,
                Foreground = new SolidColorBrush(Color.Parse("#88aaee")),
                FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center
            });
            compHeader.Children.Add(new TextBlock
            {
                Text = $"{grp.Count()}건", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
            var compNode = new TreeViewItem { Header = compHeader, IsExpanded = true };

            foreach (var s in grp)
                compNode.Items.Add(MakeSampleNode(s));

            TaskTreeView.Items.Add(compNode);
        }
    }

    private TreeViewItem MakeSampleNode(MyTaskService.SampleRow s)
    {
        string label = (string.IsNullOrEmpty(s.시료명) ? "(시료명 없음)" : s.시료명)
                       + (string.IsNullOrEmpty(s.견적번호) ? "" : $"  {s.견적번호}");

        var cb = new CheckBox
        {
            IsChecked = _basket.ContainsKey($"meok:{s.RowId}"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 2, 0)
        };
        cb.IsCheckedChanged += (_, _) =>
        {
            if (cb.IsChecked.GetValueOrDefault()) _basket[$"meok:{s.RowId}"] = label;
            else _basket.Remove($"meok:{s.RowId}");
            DetailPanelChanged?.Invoke(BuildSelectedListPanel());
        };
        _treeCheckboxes.Add((s.RowId, cb));

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        header.Children.Add(cb);
        header.Children.Add(new TextBlock { Text = "🧫", FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center });
        header.Children.Add(new TextBlock
        {
            Text = string.IsNullOrEmpty(s.시료명) ? "(시료명 없음)" : s.시료명,
            FontFamily = Font, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(s.견적번호))
            header.Children.Add(new TextBlock
            {
                Text = s.견적번호, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });

        var item = new TreeViewItem { Header = header };
        item.Tag = s;
        return item;
    }

    // ── 비용부담금 트리 (여수+율촌+세풍 통합, 폐수의뢰및결과 기반) ──────────
    private void LoadAllWasteTree(string dateStr)
    {
        var items = MyTaskService.GetRequestListItems(dateStr);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode("비용부담금 의뢰 없음"));
            return;
        }

        // 구분별 그룹
        var byGubun = items.GroupBy(i => i.구분)
            .OrderBy(g => g.Key == "여수" ? 0 : g.Key == "율촌" ? 1 : 2);

        foreach (var gubunGrp in byGubun)
        {
            var (icon, color) = gubunGrp.Key switch
            {
                "여수" => ("🌊", "#88aacc"),
                "율촌" => ("🏗", "#aaccaa"),
                _      => ("🏭", "#ccaa88"),
            };

            var gubunHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            gubunHeader.Children.Add(new TextBlock { Text = icon, FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = gubunGrp.Key, FontFamily = Font, FontSize = AppTheme.FontLG,
                Foreground = new SolidColorBrush(Color.Parse(color)),
                FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center,
            });
            gubunHeader.Children.Add(new TextBlock
            {
                Text = $"{gubunGrp.Count()}건", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
            });

            var gubunNode = new TreeViewItem { Header = gubunHeader, IsExpanded = true };

            foreach (var ri in gubunGrp.OrderBy(i => i.업체명))
                gubunNode.Items.Add(MakeRequestListNode(ri));

            TaskTreeView.Items.Add(gubunNode);
        }

        // Show3 — 담당항목 체크리스트 초기화
        _userAssigned = MyTaskService.GetUserAssignedItems(dateStr, _employeeName);
        _selectedAnalytesForNote.Clear();
        foreach (var a in _userAssigned) _selectedAnalytesForNote.Add(a.FullName);
        foreach (var b in _borrowedAnalytes) _selectedAnalytesForNote.Add(b);
        EditPanelChanged?.Invoke(BuildUserAssignedPanel());
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
        string label = ri.업체명 + (string.IsNullOrEmpty(ri.관리번호) ? "" : $"  {ri.관리번호}");

        var cb = new CheckBox
        {
            IsChecked = _basket.ContainsKey($"waste:{ri.Id}"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 2, 0)
        };
        cb.IsCheckedChanged += (_, _) =>
        {
            if (cb.IsChecked.GetValueOrDefault()) _basket[$"waste:{ri.Id}"] = label;
            else _basket.Remove($"waste:{ri.Id}");
            DetailPanelChanged?.Invoke(BuildSelectedListPanel());
        };
        _treeCheckboxes.Add((ri.Id, cb));

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        header.Children.Add(cb);
        header.Children.Add(new TextBlock
        {
            Text = ri.HasResult ? "🟢" : "🔴",
            FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = ri.업체명, FontFamily = Font, FontSize = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(ri.관리번호))
            header.Children.Add(new TextBlock
            {
                Text = ri.관리번호, FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });

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

        // 시설명별 그룹화
        var byFacility = items.GroupBy(i => i.시설명).OrderBy(g => g.Key);

        foreach (var grp in byFacility)
        {
            var grpHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            grpHeader.Children.Add(new TextBlock { Text = "⚙", FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center });
            grpHeader.Children.Add(new TextBlock
            {
                Text = grp.Key, FontFamily = Font, FontSize = AppTheme.FontLG,
                Foreground = new SolidColorBrush(Color.Parse("#bb99cc")),
                FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center
            });
            grpHeader.Children.Add(new TextBlock
            {
                Text = $"{grp.Count()}건", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
            var grpNode = new TreeViewItem { Header = grpHeader, IsExpanded = true };

            foreach (var fi in grp)
                grpNode.Items.Add(MakeFacilityNode(fi));

            TaskTreeView.Items.Add(grpNode);
        }

        // Show3 — 담당항목 체크리스트 초기화
        _userAssigned = MyTaskService.GetUserAssignedItems(dateStr, _employeeName);
        _selectedAnalytesForNote.Clear();
        foreach (var a in _userAssigned) _selectedAnalytesForNote.Add(a.FullName);
        foreach (var b in _borrowedAnalytes) _selectedAnalytesForNote.Add(b);
        EditPanelChanged?.Invoke(BuildUserAssignedPanel());
    }

    private TreeViewItem MakeFacilityNode(FacilityWorkItem fi)
    {
        string statusColor = fi.상태 switch
        {
            "완료" => "#3388dd",
            "담음" => "#dd8833",
            _     => "#666666",
        };

        string label = $"{fi.시설명}  {fi.시료명}";

        var cb = new CheckBox
        {
            IsChecked = _basket.ContainsKey($"facility:{fi.Id}"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 2, 0)
        };
        cb.IsCheckedChanged += (_, _) =>
        {
            if (cb.IsChecked.GetValueOrDefault()) _basket[$"facility:{fi.Id}"] = label;
            else _basket.Remove($"facility:{fi.Id}");
        };
        _treeCheckboxes.Add((fi.Id, cb));

        // 상태 표시: 작은 컬러 점 (체크박스처럼 보이지 않게)
        var statusDot = new Border
        {
            Width = 7, Height = 7,
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.Parse(statusColor)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 2, 0)
        };

        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        header.Children.Add(cb);
        header.Children.Add(statusDot);
        header.Children.Add(new TextBlock
        {
            Text = fi.시료명, FontFamily = Font, FontSize = AppTheme.FontMD,
            VerticalAlignment = VerticalAlignment.Center
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
        if (_isRefreshing) return;  // 트리 재빌드 중 이벤트 억제
        if (TaskTreeView.SelectedItem is not TreeViewItem sel) return;

        string dateStr = _selectedDate.ToString("yyyy-MM-dd");

        switch (sel.Tag)
        {
            case MyTaskService.SampleRow row:
                _selectedSample = row;
                _myExtraItems.Clear();
        _removedBaseItems.Clear();
                _currentAssignments = MyTaskService.GetAnalytesForSampleRow(row.RowId, dateStr);
                // 클릭 시 바스켓에 자동 추가
                if (!_basket.ContainsKey($"meok:{row.RowId}"))
                {
                    string lbl = (string.IsNullOrEmpty(row.시료명) ? "(시료명 없음)" : row.시료명);
                    _basket[$"meok:{row.RowId}"] = lbl;
                    var entry = _treeCheckboxes.Find(t => t.Id == row.RowId);
                    if (entry.Cb != null) entry.Cb.IsChecked = true;
                }
                _markedColumns = MyTaskService.GetMarkedColumnsForRow(row.RowId);
                DetailPanelChanged?.Invoke(BuildSelectedListPanel());
                EditPanelChanged?.Invoke(BuildMyItemsPanel());
                StatsPanelChanged?.Invoke(BuildAnalyteListPanel());
                break;
            case WasteRequestItem wi:
                DetailPanelChanged?.Invoke(BuildWasteItemDetail(wi));
                EditPanelChanged?.Invoke(null);
                StatsPanelChanged?.Invoke(null);
                break;
            case FacilityWorkItem fi:
                break;
            case MyTaskService.RequestListItem ri:
                _selectedSample = null;
                _myExtraItems.Clear();
        _removedBaseItems.Clear();
                _currentAssignments = MyTaskService.GetAnalytesForWasteRow(ri.Id, dateStr);
                if (!_basket.ContainsKey($"waste:{ri.Id}"))
                {
                    _basket[$"waste:{ri.Id}"] = ri.업체명 + (string.IsNullOrEmpty(ri.관리번호) ? "" : $"  {ri.관리번호}");
                    var entry = _treeCheckboxes.Find(t => t.Id == ri.Id);
                    if (entry.Cb != null) entry.Cb.IsChecked = true;
                }
                DetailPanelChanged?.Invoke(BuildSelectedListPanel());
                EditPanelChanged?.Invoke(BuildMyItemsPanel());
                StatsPanelChanged?.Invoke(BuildOthersItemsPanel());
                break;
        }
    }

    // ── Show2 — 바스켓 선택된 아이템 목록 ────────────────────────────────────
    private Control BuildSelectedListPanel()
    {
        var root = new StackPanel { Spacing = 4, Margin = new Thickness(8) };

        if (_basket.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "선택된 항목 없음",
                FontFamily = Font, FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted")
            });
        }
        else
        {
            foreach (var kv in _basket)
            {
                var kvLocal = kv;
                var inner = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                inner.Children.Add(new TextBlock { Text = "✔", FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
                inner.Children.Add(new TextBlock
                {
                    Text = kvLocal.Value, FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center
                });

                // 수질분석센터 시료: 클릭 시 Show4에 의뢰 항목 마킹 표시
                if (_activeCategory == "수질분석센터" && kvLocal.Key.StartsWith("meok:") &&
                    int.TryParse(kvLocal.Key.Substring(5), out int rowId))
                {
                    var itemBorder = new Border
                    {
                        Background = Brushes.Transparent,
                        CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(2, 1),
                        Cursor = new Cursor(StandardCursorType.Hand),
                        Child = inner
                    };
                    itemBorder.PointerPressed += (_, _) =>
                    {
                        _markedColumns = MyTaskService.GetMarkedColumnsForRow(rowId);
                        StatsPanelChanged?.Invoke(BuildAnalyteListPanel());
                    };
                    root.Children.Add(itemBorder);
                }
                else
                {
                    root.Children.Add(inner);
                }
            }
        }

        // Excel 출력 버튼 — 분석항목은 Show3 체크리스트에서 선택
        var checkedCount = _selectedAnalytesForNote.Count;
        var btnLabel = checkedCount > 0 ? $"📊 Excel 생성  ({checkedCount}항목)" : "📊 Excel 생성";
        var btnExcel = new Button
        {
            Content = btnLabel,
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Padding = new Thickness(10, 4),
            Background = new SolidColorBrush(Color.Parse("#1565c0")),
            Foreground = Brushes.White,
            CornerRadius = new CornerRadius(4)
        };
        btnExcel.Click += async (_, _) => await ExportNotebookToExcel();
        root.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 4, 0, 0) });
        root.Children.Add(btnExcel);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }

    // ── 수질분석센터: Show2 — 선택된 시료의 전체 의뢰 항목 ─────────────────
    private Control BuildAllItemsPanel(MyTaskService.SampleRow row)
    {
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"📋  {row.약칭}  —  {row.시료명}",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });
        root.Children.Add(new TextBlock
        {
            Text = $"채취일: {_selectedDate:yyyy-MM-dd}  /  총 {_currentAssignments.Count}항목",
            FontFamily = Font, FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted"),
            Margin = new Thickness(0, 2, 0, 6)
        });

        if (_currentAssignments.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "의뢰된 분석항목이 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("FgMuted")
            });
        }
        else
        {
            foreach (var a in _currentAssignments)
            {
                var card = new Border
                {
                    Background = AppRes("PanelInnerBg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 5),
                    Margin = new Thickness(0, 1)
                };
                var row2 = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
                row2.Children.Add(new TextBlock
                {
                    Text = string.IsNullOrEmpty(a.ShortName) ? a.FullName : a.ShortName,
                    FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppRes("AppFg"), Width = 80
                });
                if (a.ShortName != a.FullName && !string.IsNullOrEmpty(a.ShortName))
                    row2.Children.Add(new TextBlock
                    {
                        Text = a.FullName, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppRes("FgMuted"), Width = 160
                    });
                if (!string.IsNullOrEmpty(a.AssignedAnalyst))
                    row2.Children.Add(new TextBlock
                    {
                        Text = a.AssignedAnalyst, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = new SolidColorBrush(Color.Parse("#88ccaa"))
                    });
                card.Child = row2;
                root.Children.Add(card);
            }
        }

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }

    // ── 수질분석센터: Show3 — 나의 담당 항목 ───────────────────────────────
    private Control BuildMyItemsPanel()
    {
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = "✏  나의 담당 항목",
            FontFamily = Font, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        var myItems = _currentAssignments
            .Where(a =>
                (_myExtraItems.Contains(a.FullName) ||
                 (a.AssignedAnalyst == _employeeName && !_removedBaseItems.Contains(a.FullName))))
            .ToList();

        if (myItems.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "담당 항목이 없습니다. Show4에서 항목을 추가하세요.",
                FontFamily = Font, FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted"),
                TextWrapping = TextWrapping.Wrap
            });
        }
        else
        {
            foreach (var a in myItems)
            {
                var aLocal = a;
                bool isExtra = _myExtraItems.Contains(aLocal.FullName);
                var card = new Border
                {
                    Background = isExtra
                        ? new SolidColorBrush(Color.Parse("#1a2a1a"))
                        : AppRes("PanelInnerBg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(6, 4),
                    Margin = new Thickness(0, 1),
                    BorderBrush = isExtra
                        ? new SolidColorBrush(Color.Parse("#336633"))
                        : Brushes.Transparent,
                    BorderThickness = new Thickness(1)
                };
                var row2 = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                row2.Children.Add(new TextBlock
                {
                    Text = isExtra ? "🔄" : "✅",
                    FontSize = AppTheme.FontBase, VerticalAlignment = VerticalAlignment.Center
                });
                row2.Children.Add(new TextBlock
                {
                    Text = string.IsNullOrEmpty(aLocal.ShortName) ? aLocal.FullName : aLocal.ShortName,
                    FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppRes("AppFg"), Width = 80, VerticalAlignment = VerticalAlignment.Center
                });
                if (aLocal.ShortName != aLocal.FullName && !string.IsNullOrEmpty(aLocal.ShortName))
                    row2.Children.Add(new TextBlock
                    {
                        Text = aLocal.FullName, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
                    });
                if (isExtra && !string.IsNullOrEmpty(aLocal.AssignedAnalyst))
                    row2.Children.Add(new TextBlock
                    {
                        Text = $"← {aLocal.AssignedAnalyst}", FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = new SolidColorBrush(Color.Parse("#ccaa44")),
                        VerticalAlignment = VerticalAlignment.Center
                    });

                // X 버튼
                var removeBtn = new Button
                {
                    Content = new TextBlock { Text = "✕", FontSize = AppTheme.FontSM },
                    Background = Brushes.Transparent,
                    Foreground = AppRes("FgMuted"),
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(4, 1),
                    Margin = new Thickness(4, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                removeBtn.Click += (_, _) =>
                {
                    if (isExtra) _myExtraItems.Remove(aLocal.FullName);
                    else _removedBaseItems.Add(aLocal.FullName);
                    EditPanelChanged?.Invoke(BuildMyItemsPanel());
                    StatsPanelChanged?.Invoke(BuildOthersItemsPanel());
                };
                row2.Children.Add(removeBtn);

                card.Child = row2;
                root.Children.Add(card);
            }
        }

        // 분석노트 출력 버튼
        var printBtn = new Button
        {
            Content = new TextBlock { Text = "🖨  분석노트 출력", FontFamily = Font, FontSize = AppTheme.FontMD },
            Background = new SolidColorBrush(Color.Parse("#1a3a2a")),
            Foreground = new SolidColorBrush(Color.Parse("#88ccaa")),
            BorderBrush = new SolidColorBrush(Color.Parse("#336644")),
            Margin = new Thickness(0, 10, 0, 0),
            Padding = new Thickness(12, 6),
            HorizontalAlignment = HorizontalAlignment.Left
        };
        printBtn.Click += (_, _) => ShowAnalysisNote(myItems);
        root.Children.Add(printBtn);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }

    // ── Show4 — 의뢰 포함 항목 (내 업무 아닌 것은 추가 버튼 표시) ──────────
    private Control BuildOthersItemsPanel()
    {
        var root = new StackPanel { Spacing = 2, Margin = new Thickness(8) };

        foreach (var a in _currentAssignments)
        {
            var aLocal = a;
            bool isMine = aLocal.AssignedAnalyst == _employeeName || _myExtraItems.Contains(aLocal.FullName);
            string displayName = string.IsNullOrEmpty(aLocal.ShortName) ? aLocal.FullName : aLocal.ShortName;
            bool hasAlias = !string.IsNullOrEmpty(aLocal.ShortName) && aLocal.ShortName != aLocal.FullName;

            var row2 = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

            // 약칭 배지
            if (hasAlias)
            {
                var badge = new Border
                {
                    Background = new SolidColorBrush(Color.Parse(isMine ? "#222222" : "#1a2a3a")),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(4, 1),
                    Margin = new Thickness(0, 0, 2, 0)
                };
                badge.Child = new TextBlock
                {
                    Text = aLocal.ShortName, FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = new SolidColorBrush(Color.Parse(isMine ? "#555555" : "#88aaee")),
                };
                row2.Children.Add(badge);
            }
            row2.Children.Add(new TextBlock
            {
                Text = aLocal.FullName,
                FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = isMine ? AppRes("FgMuted") : AppRes("AppFg"),
                Width = 160
            });
            if (!string.IsNullOrEmpty(aLocal.AssignedAnalyst))
                row2.Children.Add(new TextBlock
                {
                    Text = aLocal.AssignedAnalyst, FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = new SolidColorBrush(Color.Parse(isMine ? "#555555" : "#ccaa44")),
                    Width = 70
                });
            if (!isMine)
            {
                var addBtn = new Button
                {
                    Content = new TextBlock { Text = "+", FontFamily = Font, FontSize = AppTheme.FontBase },
                    Background = new SolidColorBrush(Color.Parse("#1a2a1a")),
                    Foreground = new SolidColorBrush(Color.Parse("#88ccaa")),
                    BorderBrush = new SolidColorBrush(Color.Parse("#336633")),
                    Padding = new Thickness(6, 1),
                };
                addBtn.Click += (_, _) =>
                {
                    _myExtraItems.Add(aLocal.FullName);
                    EditPanelChanged?.Invoke(BuildMyItemsPanel());
                    StatsPanelChanged?.Invoke(BuildOthersItemsPanel());
                };
                row2.Children.Add(addBtn);
            }
            root.Children.Add(row2);
        }

        if (_currentAssignments.Count == 0)
            root.Children.Add(new TextBlock
            {
                Text = "항목 없음", FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted")
            });

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }

    // ── 분석노트 출력 ────────────────────────────────────────────────────────
    private void ShowAnalysisNote(List<MyTaskService.AnalyteAssignment> items)
    {
        if (_selectedSample == null) return;

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(16) };

        root.Children.Add(new TextBlock
        {
            Text = "분석 기록 노트",
            FontFamily = Font, FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg")
        });
        root.Children.Add(new Border
        {
            Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 4, 0, 4)
        });

        void AddRow(string label, string value)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
            sp.Children.Add(new TextBlock
            {
                Text = label, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppRes("FgMuted"), Width = 70
            });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppRes("AppFg")
            });
            root.Children.Add(sp);
        }

        AddRow("날짜", _selectedDate.ToString("yyyy-MM-dd"));
        AddRow("업체", _selectedSample.약칭);
        AddRow("시료명", _selectedSample.시료명);
        AddRow("분석자", _employeeName);
        AddRow("항목수", $"{items.Count}항목");

        root.Children.Add(new Border
        {
            Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 6, 0, 2)
        });
        root.Children.Add(new TextBlock
        {
            Text = "분석 항목 목록",
            FontFamily = Font, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 4)
        });

        int idx = 1;
        foreach (var a in items)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock
            {
                Text = $"{idx++}.", FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppRes("FgMuted"), Width = 24
            });
            sp.Children.Add(new TextBlock
            {
                Text = a.FullName, FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppRes("AppFg")
            });
            if (!string.IsNullOrEmpty(a.ShortName) && a.ShortName != a.FullName)
                sp.Children.Add(new TextBlock
                {
                    Text = $"({a.ShortName})", FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = AppRes("FgMuted")
                });
            root.Children.Add(sp);
        }

        // Show2에 노트 표시
        DetailPanelChanged?.Invoke(new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        });
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

    // ── Show3 처리시설 — 사용자 분장항목 ────────────────────────────────────────
    private Control BuildUserAssignedPanel()
    {
        var root = new StackPanel { Spacing = 2, Margin = new Thickness(8, 6) };

        root.Children.Add(new TextBlock
        {
            Text = "🖊  나의 담당 항목",
            FontFamily = Font, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"), Margin = new Thickness(0, 0, 0, 6)
        });

        if (_userAssigned.Count == 0 && _borrowedAnalytes.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "담당 항목이 없습니다. Show4에서 항목을 추가하세요.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap
            });
        }
        else
        {
            // 본인 담당항목 체크리스트
            foreach (var a in _userAssigned)
            {
                var aLocal = a;
                bool isChecked = _selectedAnalytesForNote.Contains(a.FullName);
                var cb = new CheckBox { IsChecked = isChecked, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) };
                cb.IsCheckedChanged += (_, _) =>
                {
                    if (cb.IsChecked.GetValueOrDefault()) _selectedAnalytesForNote.Add(aLocal.FullName);
                    else _selectedAnalytesForNote.Remove(aLocal.FullName);
                    DetailPanelChanged?.Invoke(BuildSelectedListPanel());
                };

                var shortName = string.IsNullOrEmpty(a.ShortName) ? a.FullName : a.ShortName;
                var badge = new Border { Background = new SolidColorBrush(Color.Parse("#1a2a3a")), CornerRadius = new CornerRadius(3), Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center };
                badge.Child = new TextBlock { Text = shortName, FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#88aaee")) };

                var itemRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                itemRow.Children.Add(cb);
                itemRow.Children.Add(badge);
                itemRow.Children.Add(new TextBlock { Text = a.FullName, FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis });

                root.Children.Add(new Border { Background = AppRes("PanelInnerBg"), CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 5), Margin = new Thickness(0, 1), Child = itemRow });
            }

            // 추가 항목 (Show4에서 클릭으로 추가된 것)
            if (_borrowedAnalytes.Count > 0)
            {
                root.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 8, 0, 4) });
                root.Children.Add(new TextBlock
                {
                    Text = "추가 항목",
                    FontFamily = Font, FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
                    Foreground = new SolidColorBrush(Color.Parse("#88ccaa")), Margin = new Thickness(0, 0, 0, 2)
                });
                foreach (var bName in _borrowedAnalytes.ToList())
                {
                    var bLocal = bName;
                    var bShort = _allAnalytes.FirstOrDefault(x => x.FullName == bLocal)?.ShortName ?? bLocal;

                    bool bChecked = _selectedAnalytesForNote.Contains(bLocal);
                    var bCb = new CheckBox { IsChecked = bChecked, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) };
                    bCb.IsCheckedChanged += (_, _) =>
                    {
                        if (bCb.IsChecked.GetValueOrDefault()) _selectedAnalytesForNote.Add(bLocal);
                        else _selectedAnalytesForNote.Remove(bLocal);
                        DetailPanelChanged?.Invoke(BuildSelectedListPanel());
                    };

                    var bBadge = new Border { Background = new SolidColorBrush(Color.Parse("#1b5e20")), CornerRadius = new CornerRadius(3), Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center };
                    bBadge.Child = new TextBlock { Text = bShort, FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = new SolidColorBrush(Color.Parse("#a5d6a7")) };

                    var delTb = new TextBlock { Text = "×", FontFamily = Font, FontSize = AppTheme.FontLG, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0, 0, 0), Cursor = new Cursor(StandardCursorType.Hand) };
                    delTb.PointerPressed += (_, _) =>
                    {
                        _borrowedAnalytes.Remove(bLocal);
                        _selectedAnalytesForNote.Remove(bLocal);
                        StatsPanelChanged?.Invoke(BuildAnalyteListPanel());
                        EditPanelChanged?.Invoke(BuildUserAssignedPanel());
                    };

                    var bRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                    bRow.Children.Add(bCb);
                    bRow.Children.Add(bBadge);
                    bRow.Children.Add(new TextBlock { Text = bLocal, FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = new SolidColorBrush(Color.Parse("#88ccaa")), VerticalAlignment = VerticalAlignment.Center });
                    bRow.Children.Add(delTb);

                    root.Children.Add(new Border { Background = new SolidColorBrush(Color.Parse("#1a3a1a")), CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 5), Margin = new Thickness(0, 1), Child = bRow });
                }
            }
        }

        // 노트 출력 버튼
        root.Children.Add(new Border { Height = 1, Background = AppRes("BorderSubtle"), Margin = new Thickness(0, 8, 0, 4) });
        var noteCount = _selectedAnalytesForNote.Count;
        var noteBtnText = noteCount > 0 ? $"\U0001f5a8 분석노트 출력  ({noteCount}항목)" : "\U0001f5a8 분석노트 출력";
        var btnNote = new Button
        {
            Content = noteBtnText,
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Background = new SolidColorBrush(Color.Parse("#1b5e20")),
            Foreground = new SolidColorBrush(Color.Parse("#a5d6a7")),
            Padding = new Thickness(12, 6), CornerRadius = new CornerRadius(4)
        };
        btnNote.Click += async (_, _) => await ExportNotebookToExcel();
        root.Children.Add(btnNote);

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }

    // =========================================================================
    // Show4 — 항상 표시: 전체 분석항목 목록 (날짜별 담당자 + 선택 시 O표시 마킹)
    // =========================================================================
    private void LoadShow4Analytes()
    {
        string dateStr = _selectedDate.ToString("yyyy-MM-dd");
        _allAnalytes = MyTaskService.GetAllAnalytesWithAssignments(dateStr);
        _markedColumns = null;
        StatsPanelChanged?.Invoke(BuildAnalyteListPanel());
    }

    private Control BuildAnalyteListPanel()
    {
        var root = new StackPanel { Spacing = 2, Margin = new Thickness(6, 4) };

        if (_allAnalytes.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "분석항목 정보 없음", FontFamily = Font,
                FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted"),
                Margin = new Thickness(4)
            });
            return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        }

        bool hasMarks = _markedColumns != null;

        foreach (var a in _allAnalytes)
        {
            string shortDisplay = string.IsNullOrEmpty(a.ShortName) ? a.FullName : a.ShortName;
            bool isMarked = hasMarks && (_markedColumns?.Contains(a.FullName) ?? false);
            bool isBorrowed = !hasMarks && _borrowedAnalytes.Contains(a.FullName);

            // 카드 테두리
            var card = new Border
            {
                Background = isBorrowed
                    ? new SolidColorBrush(Color.Parse("#1a3a1a"))
                    : AppRes("PanelInnerBg"),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(4, 3),
                Margin = new Thickness(0, 1),
            };

            // Grid: [마커?] [badge] [전체명] [담당자/indicator]
            var grid = new Grid();
            if (hasMarks)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            if (!hasMarks)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            int col = 0;

            // 의뢰 여부 아이콘 (마킹 모드일 때만)
            if (hasMarks)
            {
                var marker = new TextBlock
                {
                    Text = isMarked ? "🟢" : "🔴",
                    FontSize = AppTheme.FontSM,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0)
                };
                Grid.SetColumn(marker, col++);
                grid.Children.Add(marker);
            }

            // 약칭 배지
            var (badgeBg, badgeFg) = isMarked
                ? ("#1a3a1a", "#88ccaa")
                : ("#1a2a3a", "#88aaee");

            var badge = new Border
            {
                Background = new SolidColorBrush(Color.Parse(badgeBg)),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1),
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            badge.Child = new TextBlock
            {
                Text = shortDisplay,
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = new SolidColorBrush(Color.Parse(badgeFg))
            };
            Grid.SetColumn(badge, col++);
            grid.Children.Add(badge);

            // 전체명
            var nameTb = new TextBlock
            {
                Text = a.FullName,
                FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = isMarked ? AppRes("AppFg") : AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(nameTb, col++);
            grid.Children.Add(nameTb);

            // 담당자
            if (!string.IsNullOrEmpty(a.AssignedAnalyst))
            {
                var analystTb = new TextBlock
                {
                    Text = a.AssignedAnalyst,
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 0, 0, 0)
                };
                Grid.SetColumn(analystTb, col++);
                grid.Children.Add(analystTb);
            }

            // +/check 표시 (마킹 모드가 아닐 때)
            if (!hasMarks)
            {
                var indicator = new TextBlock
                {
                    Text = isBorrowed ? "✓" : "+",
                    FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = isBorrowed
                        ? new SolidColorBrush(Color.Parse("#88ccaa"))
                        : AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 0, 0)
                };
                Grid.SetColumn(indicator, col);
                grid.Children.Add(indicator);
            }

            card.Child = grid;
            root.Children.Add(card);
            // 마킹 모드가 아닐 때: 클릭으로 Show3에 항목 추가 (이미 추가된 항목은 무시)
            if (!hasMarks && !isBorrowed)
            {
                var aLocal = a;
                card.Cursor = new Cursor(StandardCursorType.Hand);
                card.PointerPressed += (_, _) =>
                {
                    string aName = aLocal.FullName;
                    _borrowedAnalytes.Add(aName);
                    _selectedAnalytesForNote.Add(aName);
                    StatsPanelChanged?.Invoke(BuildAnalyteListPanel());
                    EditPanelChanged?.Invoke(BuildUserAssignedPanel());
                };
            }
        }

        return new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
    }


    // =========================================================================
    // Show2 — Excel 노트 생성
    // =========================================================================
    private async Task ExportNotebookToExcel()
    {
        if (_basket.Count == 0)
        {
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
            {
                var dlg = new Window
                {
                    Title = "알림", Width = 300, Height = 120, CanResize = false,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Content = new TextBlock { Text = "선택된 시료가 없습니다.", Margin = new Thickness(20), VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center }
                };
                dlg.Show();
            });
            return;
        }

        if (_selectedAnalytesForNote.Count == 0)
        {
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
            {
                var dlg = new Window
                {
                    Title = "알림", Width = 300, Height = 120, CanResize = false,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Content = new TextBlock { Text = "Show3에서 분석항목을 선택하세요.", Margin = new Thickness(20), VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center }
                };
                dlg.Show();
            });
            return;
        }

        var analytesToExport = _selectedAnalytesForNote.ToList();
        string path = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            $"분석노트_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

        await System.Threading.Tasks.Task.Run(() =>
        {
            using var wb = new XLWorkbook();
            foreach (var analyteForExport in analytesToExport)
            {
                var cols = ETA.Services.SERVICE3.AnalysisNoteService.GetNoteColumns(analyteForExport);
                if (cols.Count == 0) cols = new List<string> { "희석배수", "비고" };

                // 시트명 최대 31자 (Excel 제한)
                string sheetName = analyteForExport.Length > 31 ? analyteForExport.Substring(0, 31) : analyteForExport;
                var ws = wb.Worksheets.Add(sheetName);

                ws.Cell(1, 1).Value = "날짜";
                ws.Cell(1, 2).Value = "시료명";
                for (int c = 0; c < cols.Count; c++)
                    ws.Cell(1, c + 3).Value = cols[c];

                var hdr = ws.Range(1, 1, 1, 2 + cols.Count);
                hdr.Style.Font.Bold = true;
                hdr.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565c0");
                hdr.Style.Font.FontColor = XLColor.White;
                hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int row = 2;
                foreach (var kv in _basket)
                {
                    var parts = kv.Value.Split(new[] { "  " }, 2, StringSplitOptions.RemoveEmptyEntries);
                    ws.Cell(row, 1).Value = parts.Length > 0 ? parts[0].Trim() : "";
                    ws.Cell(row, 2).Value = parts.Length > 1 ? parts[1].Trim() : kv.Value;
                    row++;
                }

                var dataRange = ws.Range(1, 1, row - 1, 2 + cols.Count);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder  = XLBorderStyleValues.Hair;
                ws.Columns().AdjustToContents();
            }
            wb.SaveAs(path);
        });

        await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
        {
            var dlg = new Window
            {
                Title = "Excel 생성 완료", Width = 380, Height = 130, CanResize = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = new TextBlock
                {
                    Text = $"저장됨: {path}",
                    Margin = new Thickness(20), VerticalAlignment = VerticalAlignment.Center,
                    TextWrapping = TextWrapping.Wrap
                }
            };
            dlg.Show();
        });
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
