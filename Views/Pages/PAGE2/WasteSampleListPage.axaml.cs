using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
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
using ETA.Services.Common;
using ETA.Views;

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
    public event Action<Control?>? DetailPanelSilentChanged;  // 전환 효과 없이 교체

    private string? _selectedDate;
    private Panel?  _detailPanel;
    private readonly Dictionary<string, List<string>> _datesByMonth = new();
    private readonly HashSet<string> _loadedMonths = new();
    private string _maxEnabledMonth = "";

    // ── 드래그앤드롭 상태 ──────────────────────────────────────────────────────
    private Grid?   _dragSourceRow;          // 드래그 중인 행
    private int     _dragSampleId;
    private string  _dragDate = "";
    private string  _dragGroup = "";

    // ── Show4: 업체 트리 ─────────────────────────────────────────────────────
    private Control?   _companyTreePanel;
    private TreeView?  _companyTreeView;
    private TextBlock? _companyTreeStatus;
    private string     _companyActiveGroup = "여수";
    private Button?    _btnCo여수;
    private Button?    _btnCo율촌;
    private Button?    _btnCo세풍;
    private TextBox?   _companySearchBox;
    private List<WasteCompany> _companyCacheAll = [];
    private List<WasteCompany> _pendingCompanies = [];  // 저장 전 선택 버퍼

    // ── pending 드래그 상태 ───────────────────────────────────────────────────
    private WasteCompany? _dragPendingItem;
    private Grid?         _dragPendingRow;

    public Control CompanyTreePanel => _companyTreePanel ??= BuildCompanyTreePanel();

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    public WasteSampleListPage()
    {
        InitializeComponent();
        // 인라인 캘린더 ESC 닫기
        InlineCalendar.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && CalendarPickerBorder.IsVisible)
            { CalendarPickerBorder.IsVisible = false; e.Handled = true; }
        };
    }

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
            // 당월 말일까지 (당월 포함)
            var maxDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month,
                DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month))
                .ToString("yyyy-MM-dd");
            var dates = WasteSampleService.GetDatesUpTo(maxDate);
            _datesByMonth.Clear();
            _loadedMonths.Clear();
            // 당월 + 전월 활성화
            _maxEnabledMonth = DateTime.Today.ToString("yyyy-MM");

            // 연-월 그룹핑
            var byMonth = dates
                .GroupBy(d => d[..7])          // "YYYY-MM"
                .OrderByDescending(g => g.Key);

            foreach (var month in byMonth)
            {
                var monthDates = month.OrderByDescending(x => x).ToList();
                _datesByMonth[month.Key] = monthDates;

                bool isEnabledMonth = string.CompareOrdinal(month.Key, _maxEnabledMonth) <= 0;
                var monthNode = MakeMonthNode(month.Key, monthDates.Count, isEnabledMonth);
                monthNode.IsExpanded = false;
                DateTreeView.Items.Add(monthNode);
            }

            Log($"날짜 로드 → {dates.Count}건");
        }
        catch (Exception ex) { Log("크래시: " + ex.Message); }
    }

    // =========================================================================
    // 날짜 트리 노드
    // =========================================================================
    private static TreeViewItem MakeMonthNode(string ym, int count, bool isEnabledMonth)
    {
        if (!DateTime.TryParse(ym + "-01", out var d)) d = DateTime.MinValue;

        var node = new TreeViewItem
        {
            Tag = $"MONTH:{ym}",
            IsEnabled = isEnabledMonth,
            IsExpanded = false,
        };

        var title = $"📅  {d.Year}년 {d.Month}월";
        if (!isEnabledMonth)
            title += " (비활성)";

        var titleFg = isEnabledMonth ? AppTheme.FgMuted : AppTheme.FgDimmed;

        node.Header = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            Children =
            {
                new TextBlock
                {
                    Text = title,
                    FontWeight = FontWeight.SemiBold,
                    FontFamily = Font, Foreground = titleFg,
                    VerticalAlignment = VerticalAlignment.Center,
                }.BindMD(),
                new Border
                {
                    Background = AppTheme.BorderSubtle, CornerRadius = new CornerRadius(8),
                    Padding = new Thickness(5,1),
                    Child = new TextBlock
                    {
                        Text = count.ToString(),
                        FontFamily = Font, Foreground = AppTheme.FgDimmed,
                    }.BindXS()
                }
            }
        };

        return node;
    }

    private static TreeViewItem MakeDateNode(string dateStr)
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
                        Foreground = AppTheme.FgPrimary,
                    }.BindMD(),
                    new Border
                    {
                        Background = AppTheme.BorderSubtle, CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = "선택",
                            FontFamily = Font, Foreground = AppTheme.FgDimmed,
                        }.BindXS(),
                    },
                }
            }
        };

        return node;
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
        if (e.AddedItems[0] is not TreeViewItem tvi || tvi.Tag is not string tag) return;

        if (tag.StartsWith("MONTH:", StringComparison.Ordinal))
        {
            var ym = tag[6..];
            // 비활성 월은 무시
            if (string.CompareOrdinal(ym, _maxEnabledMonth) > 0)
                return;

            EnsureMonthDatesLoaded(tvi, ym);
            tvi.IsExpanded = true;
            return;
        }

        var date = tag;

        _selectedDate = date;
        _pendingCompanies.Clear();
        FilterCompanyTree();
        RefreshDetail();
    }

    private void EnsureMonthDatesLoaded(TreeViewItem monthNode, string ym)
    {
        if (_loadedMonths.Contains(ym)) return;
        if (!_datesByMonth.TryGetValue(ym, out var dates)) return;

        foreach (var d in dates)
            monthNode.Items.Add(MakeDateNode(d));

        _loadedMonths.Add(ym);
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
        _pendingCompanies.Clear();
        FilterCompanyTree();
        EnsureDateInTree(date);
        RefreshDetail();
    }

    // =========================================================================
    // 상세 패널 새로고침
    // =========================================================================
    private void RefreshDetail(bool silent = false)
    {
        if (_selectedDate == null) return;
        _detailPanel = BuildSamplePanel(_selectedDate);
        if (silent && DetailPanelSilentChanged != null)
            DetailPanelSilentChanged.Invoke(_detailPanel);
        else
            DetailPanelChanged?.Invoke(_detailPanel);
    }

    // =========================================================================
    // 채수 의뢰 상세 패널 (날짜 선택 시 Show2)
    // =========================================================================
    private Panel BuildSamplePanel(string date)
    {
        var rows = WasteSampleService.GetByDate(date);
        var pendingForDate = _pendingCompanies
            .Where(c => !rows.Any(r => r.업체명 == c.업체명))
            .ToList();

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(4) };

        // 헤더
        DateTime.TryParse(date, out var d);
        root.Children.Add(new TextBlock
        {
            Text = $"🗓  채수일  {d:yyyy-MM-dd} ({DayOfWeekKr(d)})",
            FontWeight = FontWeight.SemiBold, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
        }.BindLG());
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0,0,0,4) });

        // 그룹별 섹션
        foreach (var groupKey in new[] { "여수", "율촌", "세풍" })
        {
            var groupRows = rows.Where(r => r.구분 == groupKey).ToList();
            int pendingCount = pendingForDate.Count(c => CompanyGroup(c) == groupKey || (CompanyGroup(c) == "기타" && groupKey == "여수"));

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
                        Text = pendingCount > 0
                            ? $"({groupRows.Count}건 + {pendingCount}건 대기)"
                            : $"({groupRows.Count}건)",
                        FontFamily = Font,
                        Foreground = pendingCount > 0 ? AppTheme.FgWarn : AppTheme.FgDimmed,
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindSM()
                }
            };
            root.Children.Add(groupHeader);

            // 행 목록 (드래그앤드롭 지원) — Grid 래퍼로 오버레이 인디케이터 배치
            var listPanel = new StackPanel { Spacing = 2 };
            foreach (var row in groupRows)
                listPanel.Children.Add(BuildSampleRow(row, date, groupKey));

            // pending 행 — 별도 드래그앤드롭 영역
            var pendingGroup = pendingForDate
                .Where(c => CompanyGroup(c) == groupKey || (CompanyGroup(c) == "기타" && groupKey == "여수"))
                .ToList();
            if (pendingGroup.Count > 0)
            {
                int baseSeq = groupRows.Count > 0
                    ? groupRows.Max(r => r.순서) + 1
                    : WasteSampleService.NextSeq(date, groupKey);

                var pendingListPanel = new StackPanel { Spacing = 2 };
                for (int pi = 0; pi < pendingGroup.Count; pi++)
                {
                    string previewSN = WasteSample.BuildSN(date, groupKey, baseSeq + pi);
                    pendingListPanel.Children.Add(BuildPendingRow(pendingGroup[pi], previewSN));
                }

                var pendingIndicator = new Border
                {
                    Height = 2, Background = Brush.Parse("#44bb44"),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top,
                    IsHitTestVisible = false, IsVisible = false,
                    BoxShadow = new BoxShadows(new BoxShadow { Blur = 6, Color = Color.Parse("#44bb44") }),
                };
                var pendingWrapper = new Grid();
                pendingWrapper.Children.Add(pendingListPanel);
                pendingWrapper.Children.Add(pendingIndicator);

                var capturedGroup = groupKey;
                DragDrop.SetAllowDrop(pendingWrapper, true);
                pendingWrapper.AddHandler(DragDrop.DragOverEvent,  (_, e) => OnPendingDragOver(pendingListPanel, pendingIndicator, e, capturedGroup));
                pendingWrapper.AddHandler(DragDrop.DropEvent,       (_, e) => OnPendingDrop(pendingListPanel, e, capturedGroup));
                pendingWrapper.AddHandler(DragDrop.DragLeaveEvent, (_, _) => { pendingIndicator.IsVisible = false; });

                listPanel.Children.Add(pendingWrapper);
            }

            var listWrapper = new Grid();
            listWrapper.Children.Add(listPanel);
            // 드롭 인디케이터용 오버레이 (레이아웃에 영향 없음)
            var indicatorOverlay = new Border
            {
                Height = 2, Background = Brush.Parse("#4488ff"),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                IsHitTestVisible = false, IsVisible = false,
                BoxShadow = new BoxShadows(new BoxShadow { Blur = 6, Color = Color.Parse("#4488ff") }),
                Tag = "indicator",
            };
            listWrapper.Children.Add(indicatorOverlay);

            var gk2 = groupKey;
            DragDrop.SetAllowDrop(listWrapper, true);
            listWrapper.AddHandler(DragDrop.DragOverEvent, (_, e) => OnListDragOver(listPanel, indicatorOverlay, e));
            listWrapper.AddHandler(DragDrop.DropEvent, (_, e) => OnListDrop(listPanel, e, date, gk2));
            listWrapper.AddHandler(DragDrop.DragLeaveEvent, (_, _) => { indicatorOverlay.IsVisible = false; });
            root.Children.Add(listWrapper);

            // + 업체 추가 버튼
            var gk = groupKey;
            var addBtn = new Button
            {
                Content = $"+ {groupKey} 업체 추가",
                Height = 24, FontSize = AppTheme.FontSM, FontFamily = Font,
                Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
                BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderActive,
                CornerRadius = new CornerRadius(4), Padding = new Thickness(10, 0),
                Margin = new Thickness(0, 2, 0, 8),
            };
            addBtn.Click += (_, _) => ShowAddCompanyDialog(date, gk);
            root.Children.Add(addBtn);
        }

        return root;
    }

    // =========================================================================
    // pending 행 (미저장 — 연한 색, 드래그핸들 + SN + 업체명 + 삭제)
    private Grid BuildPendingRow(WasteCompany c, string previewSN)
    {
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("24,90,*,26"),
            Background = Brush.Parse("#1a2a1a"),
            Margin = new Thickness(0, 1),
            Tag = c,
        };

        // ≡ 드래그 핸들
        var handle = new TextBlock
        {
            Text = "≡", FontSize = 16, FontFamily = Font,
            Foreground = Brush.Parse("#3a6a3a"),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Cursor = new Cursor(StandardCursorType.Hand),
            [Grid.ColumnProperty] = 0,
        };
        handle.PointerPressed += async (_, e) =>
        {
            if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed) return;
            _dragPendingItem = c;
            _dragPendingRow  = row;
            row.Opacity = 0.4;
            var data = new DataObject();
            data.Set("pendingReorder", c.업체명);
            await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
            row.Opacity = 1;
            _dragPendingRow = null;
        };
        row.Children.Add(handle);

        // SN 배지
        var snBlock = new Border
        {
            Background = Brush.Parse("#2a3a1a"),
            BorderBrush = AppTheme.BorderWarn,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(5, 1),
            Margin = new Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = previewSN, FontSize = AppTheme.FontXS, FontFamily = Font,
                Foreground = AppTheme.FgWarn,
            },
        };
        Grid.SetColumn(snBlock, 1);
        row.Children.Add(snBlock);

        var nameBlock = new TextBlock
        {
            Text = c.업체명, FontSize = AppTheme.FontMD, FontFamily = Font,
            Foreground = AppTheme.FgWarn,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(nameBlock, 2);
        row.Children.Add(nameBlock);

        var btnDel = new Button
        {
            Content = "✕", Width = 24, Height = 24,
            Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Foreground = AppTheme.FgDimmed,
            Padding = new Thickness(0),
            VerticalContentAlignment = VerticalAlignment.Center,
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        btnDel.Click += (_, _) =>
        {
            _pendingCompanies.RemoveAll(p => p.업체명 == c.업체명);
            RefreshDetail(silent: true);
            if (_companyTreeStatus != null)
            {
                if (_pendingCompanies.Count == 0)
                {
                    _companyTreeStatus.Foreground = AppTheme.FgDimmed;
                    _companyTreeStatus.Text = $"{_companyActiveGroup} — 클릭하면 목록에 추가";
                }
                else
                {
                    _companyTreeStatus.Foreground = AppTheme.FgWarn;
                    _companyTreeStatus.Text = $"추가 대기 {_pendingCompanies.Count}건 — [저장] 버튼으로 확정";
                }
            }
        };
        Grid.SetColumn(btnDel, 3);
        row.Children.Add(btnDel);

        return row;
    }

    // 개별 행 (드래그핸들 + S/N + 업체명 + 삭제)
    // =========================================================================
    private Grid BuildSampleRow(WasteSample s, string date, string groupKey)
    {
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("24,90,*,26"),
            Background = Brush.Parse(s.순서 % 2 == 1 ? "#1a1a28" : "#1e1e30"),
            Margin = new Thickness(0, 1),
            Tag = s,   // 드래그용 데이터
        };

        // ≡ 드래그 핸들
        var handle = new TextBlock
        {
            Text = "≡", FontSize = 16, FontFamily = Font,
            Foreground = Brush.Parse("#555"),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Cursor = new Cursor(StandardCursorType.Hand),
            [Grid.ColumnProperty] = 0,
        };

        // 드래그 시작: 핸들 위에서 PointerPressed
        var dt = date;
        var gk = groupKey;
        handle.PointerPressed += async (sender, e) =>
        {
            if (e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
            {
                _dragSampleId = s.Id;
                _dragDate = dt;
                _dragGroup = gk;
                _dragSourceRow = row;
                row.Opacity = 0.4;

                var data = new DataObject();
                data.Set("sampleReorder", s.Id.ToString());
                await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);

                // 드래그 끝 → 복구
                row.Opacity = 1;
                _dragSourceRow = null;
            }
        };

        // S/N
        row.Children.Add(new TextBlock
        {
            Text = s.SN, FontFamily = Font,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 4),
            [Grid.ColumnProperty] = 1,
        }.BindBase());

        // 업체명
        row.Children.Add(new TextBlock
        {
            Text = s.업체명, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 4),
            [Grid.ColumnProperty] = 2,
        }.BindMD());

        // 삭제
        var btnDel = MakeIconBtn("✕", "#883333");
        btnDel.SetValue(Grid.ColumnProperty, 3);
        btnDel.Click += (_, _) =>
        {
            WasteSampleService.Delete(s.Id);
            RefreshDetail(silent: true);
        };

        row.Children.Add(handle);
        row.Children.Add(btnDel);
        return row;
    }

    // =========================================================================
    // 드래그앤드롭 핸들러 (오버레이 인디케이터 — 레이아웃 영향 없음)
    // =========================================================================
    private void OnListDragOver(StackPanel listPanel, Border indicator, DragEventArgs e)
    {
        if (!e.Data.Contains("sampleReorder")) return;
        e.DragEffects = DragDropEffects.Move;

        var pos = e.GetPosition(listPanel);
        int insertIdx = GetInsertIndex(listPanel, pos.Y);
        double y = GetInsertY(listPanel, insertIdx);

        indicator.Margin = new Thickness(0, y - 1, 0, 0);
        indicator.IsVisible = true;
    }

    private void OnListDrop(StackPanel listPanel, DragEventArgs e, string date, string groupKey)
    {
        if (!e.Data.Contains("sampleReorder")) return;

        // 인디케이터 숨기기 (래퍼 Grid의 두 번째 자식)
        if (listPanel.Parent is Grid wrapper)
            foreach (var c in wrapper.Children)
                if (c is Border b && b.Tag?.ToString() == "indicator") b.IsVisible = false;

        if (_dragGroup != groupKey) return;

        var pos = e.GetPosition(listPanel);
        int insertIdx = GetInsertIndex(listPanel, pos.Y);

        // 데이터 인덱스 계산
        int srcDataIdx = -1;
        for (int i = 0; i < listPanel.Children.Count; i++)
            if (listPanel.Children[i] is Grid g && g.Tag is WasteSample ws && ws.Id == _dragSampleId)
            { srcDataIdx = i; break; }

        if (srcDataIdx < 0) return;

        int targetIdx = insertIdx;
        if (targetIdx > srcDataIdx) targetIdx--;
        if (targetIdx == srcDataIdx) return;

        WasteSampleService.ReorderTo(_dragSampleId, date, groupKey, targetIdx);
        RefreshDetail(silent: true);
    }

    // ── pending 드래그앤드롭 ──────────────────────────────────────────────────
    private void OnPendingDragOver(StackPanel listPanel, Border indicator, DragEventArgs e, string groupKey)
    {
        if (!e.Data.Contains("pendingReorder")) return;
        e.DragEffects = DragDropEffects.Move;
        var pos = e.GetPosition(listPanel);
        double y = GetInsertY(listPanel, GetInsertIndex(listPanel, pos.Y));
        indicator.Margin = new Thickness(0, y - 1, 0, 0);
        indicator.IsVisible = true;
    }

    private void OnPendingDrop(StackPanel listPanel, DragEventArgs e, string groupKey)
    {
        if (!e.Data.Contains("pendingReorder") || _dragPendingItem == null) return;

        if (listPanel.Parent is Grid wrapper)
            foreach (var ch in wrapper.Children)
                if (ch is Border b) b.IsVisible = false;

        var pos = e.GetPosition(listPanel);
        int insertIdx = GetInsertIndex(listPanel, pos.Y);

        // pending 내 해당 그룹의 항목만 추출해서 재정렬
        string effGroup = groupKey == "기타" ? "여수" : groupKey;
        var groupItems = _pendingCompanies
            .Where(c => { var g = CompanyGroup(c); return (g == effGroup) || (g == "기타" && effGroup == "여수"); })
            .ToList();

        int srcIdx = groupItems.IndexOf(_dragPendingItem);
        if (srcIdx < 0) return;

        int targetIdx = Math.Clamp(insertIdx, 0, groupItems.Count);
        if (targetIdx > srcIdx) targetIdx--;
        if (targetIdx == srcIdx) return;

        // _pendingCompanies 전체 리스트에서 해당 항목 재배치
        int globalSrc = _pendingCompanies.IndexOf(_dragPendingItem);
        var target = groupItems[targetIdx];
        int globalTarget = _pendingCompanies.IndexOf(target);

        _pendingCompanies.RemoveAt(globalSrc);
        int insertGlobal = _pendingCompanies.IndexOf(target);
        if (globalSrc < globalTarget) insertGlobal++;
        _pendingCompanies.Insert(Math.Clamp(insertGlobal, 0, _pendingCompanies.Count), _dragPendingItem);

        RefreshDetail(silent: true);
    }

    private static int GetInsertIndex(StackPanel listPanel, double mouseY)
    {
        for (int i = 0; i < listPanel.Children.Count; i++)
        {
            if (listPanel.Children[i] is not Control c) continue;
            double mid = c.Bounds.Y + c.Bounds.Height / 2;
            if (mouseY < mid) return i;
        }
        return listPanel.Children.Count;
    }

    private static double GetInsertY(StackPanel listPanel, int insertIdx)
    {
        if (listPanel.Children.Count == 0) return 0;
        if (insertIdx >= listPanel.Children.Count)
        {
            var last = (Control)listPanel.Children[^1];
            return last.Bounds.Y + last.Bounds.Height;
        }
        return ((Control)listPanel.Children[insertIdx]).Bounds.Y;
    }

    private static Button MakeIconBtn(string text, string fg) => new()
    {
        Content = text, FontSize = AppTheme.FontXS,
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
            Background = AppTheme.BorderSubtle,
        };

        // 날짜 선택 (기본값: 현재 선택된 날짜)
        DateTime.TryParse(date, out var parsedDate);
        var dp = new CalendarDatePicker
        {
            SelectedDate = parsedDate == DateTime.MinValue
                ? DateTime.Today
                : parsedDate,
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };

        var acb = new AutoCompleteBox
        {
            ItemsSource = companies,
            Watermark = "업체명 입력 또는 선택",
            FilterMode = AutoCompleteFilterMode.Contains,
            Height = 32, FontSize = AppTheme.FontMD, FontFamily = Font,
        };

        var btnOk = new Button
        {
            Content = "추가", Width = 80, Height = 30,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };
        var btnCancel = new Button
        {
            Content = "취소", Width = 80, Height = 30,
            Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontMD,
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
            RefreshDetail(silent: true);
        };
        btnCancel.Click += (_, _) => dlg.Close();

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(16), Spacing = 10,
            Children =
            {
                new TextBlock { Text = "채수일", FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppTheme.FgMuted },
                dp,
                new TextBlock { Text = "업체명", FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppTheme.FgMuted },
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
            FontSize = AppTheme.FontXL, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle });

        var dp = new CalendarDatePicker
        {
            SelectedDate = DateTime.Today,
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };

        var btnCreate = new Button
        {
            Content = "날짜 생성",
            Height = 30, FontSize = AppTheme.FontMD, FontFamily = Font,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
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

        root.Children.Add(new TextBlock { Text = "채수일 선택", FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = AppTheme.FgMuted });
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

        // 월 노드 찾기 — Tag가 "MONTH:yyyy-MM" 형식
        TreeViewItem? monthNode = DateTreeView.Items
            .OfType<TreeViewItem>()
            .FirstOrDefault(m => (string?)m.Tag == $"MONTH:{ym}");

        // 없으면 월 노드 새로 생성
        if (monthNode == null)
        {
            bool isEnabledMonth = string.CompareOrdinal(ym, _maxEnabledMonth) <= 0;
            monthNode = MakeMonthNode(ym, 0, isEnabledMonth);
            monthNode.IsExpanded = true;
            DateTreeView.Items.Insert(0, monthNode);
            _datesByMonth.TryAdd(ym, new List<string>());
        }
        else
        {
            // 기존 월 노드 — 아직 펼쳐지지 않았으면 기존 날짜 먼저 로드
            EnsureMonthDatesLoaded(monthNode, ym);
            monthNode.IsExpanded = true;
        }

        // 날짜 노드가 이미 있으면 스킵
        bool exists = monthNode.Items.OfType<TreeViewItem>().Any(d => (string?)d.Tag == date);
        if (exists) return;

        monthNode.Items.Insert(0, MakeDateNode(date));

        if (_datesByMonth.TryGetValue(ym, out var monthDates))
            monthDates.Insert(0, date);
        else
            _datesByMonth[ym] = new List<string> { date };

        _loadedMonths.Add(ym);
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
            Text = "폐수배출업소", FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppTheme.FgMuted,
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
            Background = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(8, 4),
            Child = headerSp,
        });
        UpdateCompanyGroupTabStyles();

        // Row 1: 검색창
        _companySearchBox = new TextBox
        {
            Watermark = "🔍 업체명 검색",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Height = 28,
            Background = AppTheme.BgSecondary,
            Foreground = AppTheme.FgSecondary,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            CornerRadius = new CornerRadius(0),
            Padding = new Thickness(8, 4),
        };
        _companySearchBox.TextChanged += (_, _) => FilterCompanyTree();
        Grid.SetRow(_companySearchBox, 1);
        root.Children.Add(_companySearchBox);

        // Row 2: 상태
        _companyTreeStatus = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin = new Thickness(8, 3),
        };
        Grid.SetRow(_companyTreeStatus, 2);
        root.Children.Add(_companyTreeStatus);

        // Row 3: 트리뷰
        _companyTreeView = new TreeView { Margin = new Thickness(5) };
        // 방향키는 탐색만 — 스페이스/엔터 또는 마우스 클릭으로만 추가
        _companyTreeView.KeyDown += (_, e) =>
        {
            if (e.Key is Avalonia.Input.Key.Space or Avalonia.Input.Key.Enter)
            {
                if (_companyTreeView.SelectedItem is TreeViewItem tvi && tvi.Tag is WasteCompany c)
                    AddPendingCompany(c);
                e.Handled = true;
            }
        };
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
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Padding = new Thickness(8, 3),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
        };
        btn.Click += (_, _) =>
        {
            _companyActiveGroup = tag;
            UpdateCompanyGroupTabStyles();
            LoadCompanyTree(forceReload: true);
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
                btn.Background  = AppTheme.BgSecondary;
                btn.Foreground  = AppTheme.FgDimmed;
                btn.BorderBrush = AppTheme.BorderMuted;
            }
        }
    }

    // DB에서 전체 목록을 새로 가져와 캐시 갱신 후 트리 재렌더
    public void LoadCompanyTree(bool forceReload = false)
    {
        if (_companyTreeView == null) return;

        if (forceReload || _companyCacheAll.Count == 0)
        {
            try { _companyCacheAll = WasteCompanyService.GetAllItems(); }
            catch (Exception ex)
            {
                if (_companyTreeStatus != null)
                {
                    _companyTreeStatus.Foreground = AppTheme.FgDanger;
                    _companyTreeStatus.Text = $"오류: {ex.Message}";
                }
                return;
            }
        }

        FilterCompanyTree();
    }

    // 캐시만 사용해 트리를 재렌더 (DB 쿼리 없음)
    private void FilterCompanyTree()
    {
        if (_companyTreeView == null) return;

        var search = _companySearchBox?.Text?.Trim() ?? "";
        var companies = _companyCacheAll
            .Where(c => CompanyGroup(c) == _companyActiveGroup)
            .Where(c => string.IsNullOrEmpty(search) ||
                        c.업체명.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        c.약칭.Contains(search, StringComparison.OrdinalIgnoreCase))
            .OrderBy(c => c.관리번호)
            .ToList();

        _companyTreeView.Items.Clear();
        foreach (var c in companies)
            _companyTreeView.Items.Add(MakeCompanyLeafNode(c));

        if (_companyTreeStatus != null)
        {
            _companyTreeStatus.Foreground = AppTheme.FgDimmed;
            string statusText = $"{_companyActiveGroup} {companies.Count}개";
            if (!string.IsNullOrEmpty(search)) statusText += $"  (검색: \"{search}\")";
            _companyTreeStatus.Text = statusText + " — 클릭하면 목록에 추가";
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
                    new TextBlock { Text = icon, FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock
                    {
                        Text = name, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = Brush.Parse(color),
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new Border
                    {
                        Background = AppTheme.BorderSubtle, CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5, 1), VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = count.ToString(), FontSize = AppTheme.FontXS,
                            FontFamily = Font, Foreground = AppTheme.FgMuted,
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
                    Text = c.약칭, FontSize = AppTheme.FontXS, FontFamily = Font,
                    Foreground = Brush.Parse(fg),
                },
            });
        }

        sp.Children.Add(new TextBlock
        {
            Text = c.업체명, FontSize = AppTheme.FontMD, FontFamily = Font,
            Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text = c.관리번호, FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
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
        AddPendingCompany(company);
    }

    private void AddPendingCompany(WasteCompany company)
    {
        _companyTreeView!.SelectedItem = null;

        // 날짜 미선택
        if (string.IsNullOrEmpty(_selectedDate))
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = AppTheme.FgWarn;
                _companyTreeStatus.Text = "왼쪽에서 날짜를 먼저 선택하세요";
            }
            return;
        }

        // 이미 DB에 있는지 확인
        var existing = WasteSampleService.GetByDate(_selectedDate);
        if (existing.Any(s => s.업체명 == company.업체명))
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = AppTheme.FgWarn;
                _companyTreeStatus.Text = $"이미 추가됨: {company.업체명}";
            }
            return;
        }

        // 이미 대기 중이면 무시
        if (_pendingCompanies.Any(c => c.업체명 == company.업체명)) return;

        _pendingCompanies.Add(company);
        RefreshDetail(silent: true);   // Show2 즉시 반영

        if (_companyTreeStatus != null)
        {
            _companyTreeStatus.Foreground = AppTheme.FgWarn;
            _companyTreeStatus.Text = $"추가 대기 {_pendingCompanies.Count}건 — [저장] 버튼으로 확정";
        }
    }

    // ── 서브버튼 BT1: 대기 목록 일괄 저장 ───────────────────────────────────
    public void SavePending(IProgress<(int done, int total)>? progress = null)
    {
        if (string.IsNullOrEmpty(_selectedDate))
        {
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = AppTheme.FgWarn;
                _companyTreeStatus.Text = "날짜를 먼저 선택하세요";
            }
            return;
        }
        if (_pendingCompanies.Count == 0) return;

        var toSave = _pendingCompanies.ToList();
        int total = toSave.Count;
        int saved = 0;

        foreach (var c in toSave)
        {
            string groupKey = CompanyGroup(c);
            if (groupKey == "기타") groupKey = "여수";
            WasteSampleService.Insert(_selectedDate, groupKey, c.업체명, c.관리번호);
            saved++;
            progress?.Report((saved, total));
        }
        // UI 조작은 반드시 UI 스레드에서
        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            _pendingCompanies.Clear();
            EnsureDateInTree(_selectedDate);
            FilterCompanyTree();
            RefreshDetail(silent: true);
            if (_companyTreeStatus != null)
            {
                _companyTreeStatus.Foreground = AppTheme.FgSuccess;
                _companyTreeStatus.Text = $"{saved}건 저장 완료";
            }
        });
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [WasteSampleList] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("Logs/WasteSampleDebug.log", line + Environment.NewLine); } catch { }
    }
}
