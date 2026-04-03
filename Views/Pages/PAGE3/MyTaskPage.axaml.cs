using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
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
    private string   _activeCategory = "측대";
    private DateTime _selectedDate   = DateTime.Today;
    private string   _employeeId     = "";
    private string   _employeeName   = "";

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // 카테고리별 활성 버튼 스타일
    private static readonly Dictionary<string, (string Bg, string Fg, string Bd)> CatStyle = new()
    {
        ["측대"]    = ("#1a2a4a", "#88aaee", "#336699"),
        ["여수"]    = ("#1e3a5a", "#88aacc", "#336699"),
        ["율촌"]    = ("#1a3a1a", "#aaccaa", "#336633"),
        ["세풍"]    = ("#3a2a1a", "#ccaa88", "#996633"),
        ["처리시설"] = ("#2a1a3a", "#bb99cc", "#663399"),
        ["의뢰목록"] = ("#1a3a2a", "#88ccaa", "#339966"),
    };

    private readonly Dictionary<string, Button> _catBtns = new();

    public MyTaskPage()
    {
        InitializeComponent();
        dpDate.SelectedDate = DateTime.Today;

        _catBtns["측대"]    = BtnMeokdae;
        _catBtns["여수"]    = BtnYeosoo;
        _catBtns["율촌"]    = BtnYulchon;
        _catBtns["세풍"]    = BtnSepung;
        _catBtns["처리시설"] = BtnFacility;
        _catBtns["의뢰목록"] = BtnRequestList;
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
    // 날짜 변경
    // =========================================================================
    private void DpDate_SelectedDateChanged(object? sender, DatePickerSelectedValueChangedEventArgs e)
    {
        if (dpDate.SelectedDate is { } d)
        {
            _selectedDate = d.DateTime;
            RefreshTree();
        }
    }

    // =========================================================================
    // 카테고리 버튼 클릭
    // =========================================================================
    private void CategoryBtn_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        _activeCategory = btn.Tag as string ?? "측대";
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
            case "측대":      LoadMeokdaeTree(dateStr);             break;
            case "여수":
            case "율촌":
            case "세풍":      LoadWasteTree(dateStr, _activeCategory); break;
            case "처리시설":  LoadFacilityTree(dateStr);            break;
            case "의뢰목록":  LoadRequestListTree(dateStr);         break;
        }
    }

    // ── 측대 트리 ─────────────────────────────────────────────────────────────
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
            Text = "🧪", FontSize = 13, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = string.IsNullOrEmpty(shortName) ? fullName : shortName,
            FontFamily = Font, FontSize = 13,
            VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(shortName) && shortName != fullName)
        {
            header.Children.Add(new TextBlock
            {
                Text = $"({fullName})", FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
        }

        var item = new TreeViewItem { Header = header };
        item.Tag = new AnalyteTag(fullName, shortName);
        return item;
    }

    private record AnalyteTag(string FullName, string ShortName);

    // ── 여수/율촌/세풍 트리 ───────────────────────────────────────────────────
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
            Text = icon, FontSize = 13, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = company, FontFamily = Font, FontSize = 13,
            Foreground = new SolidColorBrush(Color.Parse(color)),
            VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = $"{count}건", FontFamily = Font, FontSize = 11,
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
            Text = statusIcon, FontSize = 12, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = wi.항목, FontFamily = Font, FontSize = 12,
            Foreground = new SolidColorBrush(Color.Parse(statusColor)),
            VerticalAlignment = VerticalAlignment.Center
        });
        if (!string.IsNullOrEmpty(wi.배정자))
        {
            header.Children.Add(new TextBlock
            {
                Text = wi.배정자, FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
            });
        }

        var item = new TreeViewItem { Header = header };
        item.Tag = wi;
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
            Text = statusIcon, FontSize = 12, VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = fi.시설명, FontFamily = Font, FontSize = 13,
            VerticalAlignment = VerticalAlignment.Center
        });
        header.Children.Add(new TextBlock
        {
            Text = fi.시료명, FontFamily = Font, FontSize = 11,
            Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center
        });

        var item = new TreeViewItem { Header = header };
        item.Tag = fi;
        return item;
    }

    // ── 의뢰목록 트리 (폐수채수의뢰 — 선택일 기준 7일 범위) ─────────────────
    private void LoadRequestListTree(string dateStr)
    {
        DateTime.TryParse(dateStr, out var selectedDate);
        var fromDate = selectedDate.AddDays(-7);
        var from = fromDate.ToString("yyyy-MM-dd");
        var to   = dateStr;

        var items = MyTaskService.GetRequestListItemsRange(from, to);

        if (items.Count == 0)
        {
            TaskTreeView.Items.Add(MakeEmptyNode($"채수 의뢰 없음 ({from} ~ {to})"));
            return;
        }

        // 날짜별 그룹 (내림차순: 최근 날짜가 위)
        var byDate = items.GroupBy(i => i.채수일)
                          .OrderByDescending(g => g.Key);

        foreach (var dateGroup in byDate)
        {
            // 날짜 헤더 노드
            DateTime.TryParse(dateGroup.Key, out var d);
            bool isToday = d.Date == selectedDate.Date;
            string dayKr = d.DayOfWeek switch
            {
                DayOfWeek.Monday    => "월",
                DayOfWeek.Tuesday   => "화",
                DayOfWeek.Wednesday => "수",
                DayOfWeek.Thursday  => "목",
                DayOfWeek.Friday    => "금",
                DayOfWeek.Saturday  => "토",
                _                   => "일",
            };
            var dateHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            dateHeader.Children.Add(new TextBlock
            {
                Text = $"{d:MM/dd}({dayKr})", FontFamily = Font, FontSize = 13,
                FontWeight = isToday ? FontWeight.Bold : FontWeight.Normal,
                Foreground = isToday
                    ? new SolidColorBrush(Color.Parse("#aaddff"))
                    : new SolidColorBrush(Color.Parse("#888888")),
                VerticalAlignment = VerticalAlignment.Center,
            });
            dateHeader.Children.Add(new TextBlock
            {
                Text = $"{dateGroup.Count()}건", FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
            });

            var dateNode = new TreeViewItem { Header = dateHeader, IsExpanded = isToday };

            // 날짜 내 구분별 서브그룹
            var byGubun = dateGroup
                .GroupBy(i => i.구분)
                .OrderBy(g => g.Key == "여수" ? 0 : g.Key == "율촌" ? 1 : g.Key == "세풍" ? 2 : 3);

            foreach (var grp in byGubun)
            {
                var (icon, color) = grp.Key switch
                {
                    "여수" => ("🌊", "#88aacc"),
                    "율촌" => ("🏗",  "#aaccaa"),
                    "세풍" => ("🏭", "#ccaa88"),
                    _     => ("📋", "#88ccaa"),
                };

                var groupHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                groupHeader.Children.Add(new TextBlock { Text = icon, FontSize = 12, VerticalAlignment = VerticalAlignment.Center });
                groupHeader.Children.Add(new TextBlock
                {
                    Text = grp.Key, FontFamily = Font, FontSize = 12,
                    Foreground = new SolidColorBrush(Color.Parse(color)),
                    VerticalAlignment = VerticalAlignment.Center,
                });
                groupHeader.Children.Add(new TextBlock
                {
                    Text = $"{grp.Count()}건", FontFamily = Font, FontSize = 11,
                    Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
                });

                var groupNode = new TreeViewItem { Header = groupHeader, IsExpanded = isToday };

                foreach (var item in grp)
                {
                    var itemHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                    itemHeader.Children.Add(new TextBlock
                    {
                        Text = item.SN, FontFamily = Font, FontSize = 11,
                        Foreground = AppRes("FgMuted"), Width = 70,
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                    itemHeader.Children.Add(new TextBlock
                    {
                        Text = item.업체명, FontFamily = Font, FontSize = 12,
                        Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center,
                    });
                    if (!string.IsNullOrEmpty(item.확인자))
                        itemHeader.Children.Add(new TextBlock
                        {
                            Text = item.확인자, FontFamily = Font, FontSize = 11,
                            Foreground = new SolidColorBrush(Color.Parse("#66aaee")),
                            VerticalAlignment = VerticalAlignment.Center,
                        });

                    groupNode.Items.Add(new TreeViewItem { Header = itemHeader, Tag = item });
                }

                dateNode.Items.Add(groupNode);
            }

            TaskTreeView.Items.Add(dateNode);
        }
    }

    // ── 빈 노드 ───────────────────────────────────────────────────────────────
    private TreeViewItem MakeEmptyNode(string text)
    {
        return new TreeViewItem
        {
            Header = new TextBlock
            {
                Text = text, FontFamily = Font, FontSize = 12,
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
            AnalyteTag tag                      => BuildAnalyteDetail(tag),
            WasteRequestItem wi                 => BuildWasteItemDetail(wi),
            FacilityWorkItem fi                 => BuildFacilityDetail(fi),
            MyTaskService.RequestListItem req   => BuildRequestListDetail(req),
            _                                   => null,
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
            FontFamily = Font, FontSize = 15, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });
        root.Children.Add(new TextBlock
        {
            Text = tag.FullName, FontFamily = Font, FontSize = 11,
            Foreground = AppRes("FgMuted")
        });
        root.Children.Add(new TextBlock
        {
            Text = $"채취일: {dateStr}  /  총 {samples.Count}건",
            FontFamily = Font, FontSize = 12, Foreground = AppRes("FgMuted"),
            Margin = new Thickness(0, 4, 0, 4)
        });

        if (samples.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "해당 날짜에 의뢰된 시료가 없습니다.",
                FontFamily = Font, FontSize = 12, Foreground = AppRes("FgMuted")
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
                    Text = s.Company, FontFamily = Font, FontSize = 12,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("AppFg"),
                    Width = 80
                });
                row.Children.Add(new TextBlock
                {
                    Text = s.Sample, FontFamily = Font, FontSize = 12,
                    Foreground = AppRes("AppFg"), Width = 120
                });
                if (!string.IsNullOrEmpty(s.접수번호))
                    row.Children.Add(new TextBlock
                    {
                        Text = s.접수번호, FontFamily = Font, FontSize = 11,
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
            FontFamily = Font, FontSize = 15, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = 12, Foreground = AppRes("FgMuted"), Width = 70 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = 12,
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

    // ── 처리시설 상세 ────────────────────────────────────────────────────────
    private Control BuildFacilityDetail(FacilityWorkItem fi)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"⚙  {fi.시설명}",
            FontFamily = Font, FontSize = 15, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = 12, Foreground = AppRes("FgMuted"), Width = 70 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = 12,
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

    // ── 의뢰목록 상세 ─────────────────────────────────────────────────────────
    private Control BuildRequestListDetail(MyTaskService.RequestListItem req)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text = $"📋  {req.업체명}",
            FontFamily = Font, FontSize = 15, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg")
        });

        void AddRow(string label, string value, string? color = null)
        {
            if (string.IsNullOrEmpty(value)) return;
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new TextBlock { Text = label, FontFamily = Font, FontSize = 12, Foreground = AppRes("FgMuted"), Width = 60 });
            sp.Children.Add(new TextBlock
            {
                Text = value, FontFamily = Font, FontSize = 12,
                Foreground = color != null ? new SolidColorBrush(Color.Parse(color)) : AppRes("AppFg")
            });
            root.Children.Add(sp);
        }

        AddRow("SN",     req.SN);
        AddRow("구분",   req.구분);
        AddRow("채수일", req.채수일);
        AddRow("확인자", req.확인자, "#66aaee");
        AddRow("관리번호", req.관리번호);
        if (!string.IsNullOrEmpty(req.비고))
            AddRow("비고", req.비고);

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
