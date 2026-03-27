using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using ETA.Services;
using ETA.ViewModels;
using ETA.Views.Pages;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ETA.Views;

public partial class MainPage : Window
{
    public static string CurrentEmployeeId { get; set; } = "";

    private string _currentMode = "None";
    
    // ── 창 위치/레이아웃 관리 ──────────────────────────────────────────
    private WindowPositionManager? _positionManager;
    private const string LayoutStorageModePrefix = "Mode_";

    // 페이지는 처음 진입 시점에 Lazy 생성
    private AnalysisPage?      _analysisPage;
    private ContractPage?      _contractPage;
    private AgentTreePage?     _agentTreePage;
    private WasteCompanyPage?  _wasteCompanyPage;
    private PurchasePage?      _purchasePage;
    private RepairPage?       _repairPage;
    private TestReportPage?    _testReportPage;
    private ReportsPanel?      _reportsPanel;           // Content4: 출력 보관함

    // ── 견적/의뢰서 전용 4-패널 ──────────────────────────────────────────
    // Content1: 발행내역 트리  Content2: 신규작성 폼
    // Content3: 분석항목 체크  Content4: 계약업체 목록
    private QuotationHistoryPanel? _quotationHistoryPanel;
    private QuotationDetailPanel?  _quotationDetailPanel;   // Content2: 세부내역
    private QuotationNewPanel?     _quotationNewPanel;      // Content2: 신규작성
    private QuotationCheckPanel?   _quotationCheckPanel;
    private QuotationPage?         _quotationPage;
    private OrderRequestEditPanel? _orderRequestEditPanel;  // Content2: 의뢰서 편집

    private System.Action? _bt1SaveAction;

    // ── 마지막으로 표시된 issue 캐시 (트리뷰 선택 null 시 복원용) ─────
    private ETA.Models.QuotationIssue? _lastShownIssue;

    // ── 분석의뢰 상세/목록 패널 ──────────────────────────────────────
    private AnalysisRequestDetailPanel?  _analysisRequestDetailPanel;
    private AnalysisRequestListPanel?    _analysisRequestListPanel;

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ApplyTheme(true);
        
        // WindowPositionManager 초기화
        _positionManager = new WindowPositionManager(CurrentUserManager.Instance.CurrentUserId);
        
        // 윈도우 이벤트 연결
        this.Opened += MainPage_Opened;
        this.Closing += MainPage_Closing;
    }

    private void MainPage_Opened(object? sender, EventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainPage] Opened 이벤트");
    }

    private void MainPage_Closing(object? sender, WindowClosingEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainPage] Closing 이벤트 - 현재 모드 레이아웃 저장");
        if (!string.IsNullOrEmpty(_currentMode) && _currentMode != "None")
        {
            SaveCurrentModeLayout();
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  분석항목 리스트 (드래그 앤 드랍용)
    // ══════════════════════════════════════════════════════════════════════

    private ListBox?    _analysisItemsListBox;
    private Control?    _content4Container;
    private DateTime    _content4QueryStart = DateTime.Today;
    private DateTime    _content4QueryEnd   = DateTime.Today;

    private Control CreateAnalysisItemsListBox()
    {
        if (_content4Container != null) return _content4Container;

        // ── 리스트박스 ────────────────────────────────────────────────────
        _analysisItemsListBox = new ListBox
        {
            Background      = new SolidColorBrush(Color.Parse("#1e1e2e")),
            BorderThickness = new Thickness(0),
            SelectionMode   = SelectionMode.Multiple,
            Margin          = new Thickness(2, 0, 2, 2),
        };
        LoadAnalysisItems();
        _analysisItemsListBox.PointerPressed += OnAnalysisItemPointerPressed;

        // ── 날짜 범위 행 ──────────────────────────────────────────────────
        var txbRange = new TextBlock
        {
            Text              = DateTime.Today.ToString("yyyy-MM-dd"),
            FontSize          = 10,
            Foreground        = new SolidColorBrush(Color.Parse("#aaaaaa")),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            MinWidth          = 160,
            FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
        };

        var btnCal = new Button
        {
            Content         = "📅",
            Width           = 28,
            Height          = 22,
            FontSize        = 11,
            Padding         = new Thickness(0),
            Background      = new SolidColorBrush(Color.Parse("#2a3a4a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555555")),
        };
        ToolTip.SetTip(btnCal, "기간 선택 (드래그로 범위 설정)");

        var btnToday = new Button
        {
            Content         = "오늘",
            Width           = 42,
            Height          = 22,
            FontSize        = 10,
            Padding         = new Thickness(4, 0),
            Background      = new SolidColorBrush(Color.Parse("#3a5a3a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#666666")),
        };

        var dateRow = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing     = 5,
            Margin      = new Thickness(4, 4, 4, 2),
        };
        dateRow.Children.Add(txbRange);
        dateRow.Children.Add(btnCal);
        dateRow.Children.Add(btnToday);

        // ── 인라인 달력 ───────────────────────────────────────────────────
        var calendar = new Avalonia.Controls.Calendar
        {
            SelectionMode = Avalonia.Controls.CalendarSelectionMode.SingleRange,
            IsVisible     = false,
            Margin        = new Thickness(4, 0, 4, 2),
            DisplayDate   = DateTime.Today,
        };
        calendar.SelectedDates.Add(DateTime.Today);

        // ── 이벤트 연결 ───────────────────────────────────────────────────
        btnCal.Click += (_, _) => calendar.IsVisible = !calendar.IsVisible;

        btnToday.Click += (_, _) =>
        {
            _content4QueryStart = _content4QueryEnd = DateTime.Today;
            txbRange.Text       = DateTime.Today.ToString("yyyy-MM-dd");
            calendar.IsVisible  = false;
            UpdateAssignmentInfo();
        };

        calendar.SelectedDatesChanged += (_, _) =>
        {
            if (calendar.SelectedDates.Count == 0) return;
            var dates = calendar.SelectedDates.Cast<DateTime>().ToList();
            _content4QueryStart = dates.Min();
            _content4QueryEnd   = dates.Max();
            txbRange.Text = _content4QueryStart == _content4QueryEnd
                ? _content4QueryStart.ToString("yyyy-MM-dd")
                : $"{_content4QueryStart:yyyy-MM-dd} ~ {_content4QueryEnd:yyyy-MM-dd}";
            UpdateAssignmentInfo();
            if (_content4QueryStart != _content4QueryEnd)
                calendar.IsVisible = false;
        };

        // ── 컨테이너 조립 (DockPanel) ─────────────────────────────────────
        var header = new StackPanel { Spacing = 0 };
        header.Children.Add(dateRow);
        header.Children.Add(calendar);

        var dock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(header, Dock.Top);
        dock.Children.Add(header);
        dock.Children.Add(_analysisItemsListBox);

        _content4Container = dock;
        return _content4Container;
    }

    private void LoadAnalysisItems()
    {
        if (_analysisItemsListBox == null) return;

        _analysisItemsListBox.Items.Clear();
        var items = AnalysisService.GetAllItems();
        Debug.WriteLine($"[LoadAnalysisItems] 로드된 분석항목 수: {items.Count}");
        
        foreach (var item in items)
        {
            // 카드 UI: Border + Grid로 구성
            var cardBorder = new Border
            {
                Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
                BorderThickness = new Thickness(1),
                BorderBrush     = new SolidColorBrush(Color.Parse("#444455")),
                CornerRadius    = new CornerRadius(3),
                Padding         = new Thickness(6, 3),
                Margin          = new Thickness(2, 1),
                Cursor          = new Cursor(StandardCursorType.Hand),
            };
            ToolTip.SetTip(cardBorder, $"{item.Category} | ES: {item.ES}");

            var cardGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,Auto"),
            };

            // Col 0: 항목명
            var nameBlock = new TextBlock
            {
                Text              = item.Analyte,
                FontSize          = 11,
                FontWeight        = FontWeight.SemiBold,
                Foreground        = new SolidColorBrush(Color.Parse("#a0d060")),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                TextTrimming      = TextTrimming.CharacterEllipsis,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            Grid.SetColumn(nameBlock, 0);
            cardGrid.Children.Add(nameBlock);

            // Col 1: 할당 배지 (UpdateAssignmentInfo에서 갱신)
            var assignBadge = new TextBlock
            {
                Text              = "·",
                FontSize          = 9,
                Foreground        = new SolidColorBrush(Color.Parse("#555566")),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Margin            = new Thickness(4, 0, 0, 0),
            };
            Grid.SetColumn(assignBadge, 1);
            cardGrid.Children.Add(assignBadge);

            cardBorder.Tag = item.Analyte;
            cardBorder.Child = cardGrid;

            // ✅ 중요: Border에 직접 PointerPressed 이벤트 등록 (ListBox 이벤트와 상충 방지)
            cardBorder.PointerPressed += (s, e) =>
            {
                if (e.GetCurrentPoint(cardBorder).Properties.IsLeftButtonPressed)
                {
                    var analyte = (cardBorder.Tag as string) ?? "";
                    if (!string.IsNullOrEmpty(analyte))
                    {
                        var data = new DataObject();
                        data.Set("analyte", analyte);
                        DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                        System.Diagnostics.Debug.WriteLine($"[DragStart] {analyte}");
                    }
                }
            };

            // ListBoxItem에 카드 넣기
            var listBoxItem = new ListBoxItem
            {
                Content = cardBorder,
                Tag = item.Analyte,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
            };
            
            _analysisItemsListBox.Items.Add(listBoxItem);
        }

        // 할당 정보 비동기로 업데이트
        UpdateAssignmentInfo();
    }

    /// <summary>각 카드에 할당 정보를 표시 (현재 선택된 기간 기준)</summary>
    private void UpdateAssignmentInfo()
    {
        if (_analysisItemsListBox == null) return;

        foreach (var listBoxItem in _analysisItemsListBox.Items.OfType<ListBoxItem>())
        {
            if (listBoxItem.Content is Border border && border.Child is Grid grid)
            {
                var analyte = listBoxItem.Tag as string;
                if (string.IsNullOrEmpty(analyte)) continue;

                var assignees = _content4QueryStart == _content4QueryEnd
                    ? AnalysisRequestService.GetAssigneesForAnalyteOnDate(analyte, _content4QueryStart)
                    : AnalysisRequestService.GetAssigneesForAnalyteInRange(analyte, _content4QueryStart, _content4QueryEnd);

                if (grid.Children.Count > 1 && grid.Children[1] is TextBlock assignBadge)
                {
                    var baseTip = ToolTip.GetTip(border)?.ToString()?.Split('\n')[0] ?? "";
                    if (assignees.Count == 0)
                    {
                        assignBadge.Text       = "미할당";
                        assignBadge.Foreground = new SolidColorBrush(Color.Parse("#555566"));
                        ToolTip.SetTip(border, baseTip);
                    }
                    else
                    {
                        // 이름별 일수 집계 (중복 포함)
                        var grouped = assignees
                            .GroupBy(a => a)
                            .OrderByDescending(g => g.Count())
                            .ToList();
                        int uniqueCount = grouped.Count;

                        // 배지: 이름을 모두 나열
                        assignBadge.Text       = string.Join(", ", grouped.Select(g => g.Key));
                        assignBadge.Foreground = new SolidColorBrush(Color.Parse("#88cc88"));

                        // 툴팁: 이름(N일) 형식으로 중복 표시
                        bool isRange = _content4QueryStart != _content4QueryEnd;
                        var detail = isRange
                            ? string.Join(", ", grouped.Select(g =>
                                g.Count() > 1 ? $"{g.Key}({g.Count()}일)" : g.Key))
                            : string.Join(", ", grouped.Select(g => g.Key));
                        ToolTip.SetTip(border, $"{baseTip}\n담당: {detail}");
                    }
                }
            }
        }
    }

    private void OnAnalysisItemPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            // Content4의 카드에서 드래그 시작
            var listBox = sender as ListBox;
            if (listBox == null) return;

            // 클릭된 control을 따라가서 ListBoxItem 찾기
            var source = e.Source as Control;
            ListBoxItem? targetItem = null;

            // visual tree를 따라 올라가면서 ListBoxItem 찾기
            var current = source;
            while (current != null)
            {
                if (current is ListBoxItem lbi)
                {
                    targetItem = lbi;
                    break;
                }
                current = current.Parent as Control;
            }

            if (targetItem != null && targetItem.Tag is string analyte)
            {
                // 드래그 데이터 설정
                var data = new DataObject();
                data.Set("analyte", analyte);
                DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
            }
        }
    }


    // ══════════════════════════════════════════════════════════════════════
    //  레이아웃 헬퍼
    // ══════════════════════════════════════════════════════════════════════

    private void SetLeftPanelWidth(double width)
    {
        var grid = this.FindControl<Grid>("MainSplitGrid");
        if (grid != null && grid.ColumnDefinitions.Count > 0)
            grid.ColumnDefinitions[0].Width = new GridLength(width);
    }

    private void SetContentLayout(
        double content2Star = 1, double content4Star = 1,
        double upperStar   = 4, double lowerStar    = 1)
    {
        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        if (rightGrid != null && rightGrid.RowDefinitions.Count >= 3)
        {
            rightGrid.RowDefinitions[0].Height = new GridLength(upperStar, GridUnitType.Star);
            rightGrid.RowDefinitions[2].Height = lowerStar > 0
                ? new GridLength(lowerStar, GridUnitType.Star)
                : new GridLength(0);

            var lowerBorder = this.FindControl<Border>("LowerBorder");
            if (lowerBorder != null) lowerBorder.IsVisible = lowerStar > 0;

            var vSplitter = this.FindControl<GridSplitter>("VerticalSplitter");
            if (vSplitter != null) vSplitter.IsVisible = lowerStar > 0;
        }

        var upperGrid = this.FindControl<Grid>("UpperContentGrid");
        if (upperGrid != null && upperGrid.ColumnDefinitions.Count >= 3)
        {
            upperGrid.ColumnDefinitions[0].Width = new GridLength(content2Star, GridUnitType.Star);
            upperGrid.ColumnDefinitions[2].Width = content4Star > 0
                ? new GridLength(content4Star, GridUnitType.Star)
                : new GridLength(0);

            var content4Border = this.FindControl<Border>("Content4Border");
            if (content4Border != null) content4Border.IsVisible = content4Star > 0;

            var hSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
            if (hSplitter != null) hSplitter.IsVisible = content4Star > 0;
        }
    }


    // ══════════════════════════════════════════════════════════════════════
    //  메뉴 클릭
    // ══════════════════════════════════════════════════════════════════════

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Agent";

        if (_agentTreePage == null)
        {
            _agentTreePage = new AgentTreePage();
            _agentTreePage.DetailPanelChanged += panel =>
            {
                ActivePageContent2.Content = panel;
                LogContentChange("ActivePageContent2", panel);
            };
        }

        ActivePageContent1.Content = _agentTreePage;
        LogContentChange("ActivePageContent1", _agentTreePage);
        ActivePageContent2.Content = null;
        LogContentChange("ActivePageContent2", null);
        ActivePageContent3.Content = null;
        LogContentChange("ActivePageContent3", null);
        
        // ✅ Content4: 분석항목 카드 리스트 (드래그&드롭용)
        ActivePageContent4.Content = CreateAnalysisItemsListBox();
        LogContentChange("ActivePageContent4", _analysisItemsListBox);
        
        // AgentTreePage에 ListBox 참조 전달
        _agentTreePage.AnalysisItemsListBox = _analysisItemsListBox;
        
        _agentTreePage.LoadData();
        _bt1SaveAction = _agentTreePage.SaveSelected;

        SetSubMenu("저장", "새로고침", "직원 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Agent");
    }

    private void WasteCompany_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteCompany";
        if (_wasteCompanyPage == null) _wasteCompanyPage = new WasteCompanyPage();

        ActivePageContent1.Content = _wasteCompanyPage;
        LogContentChange("ActivePageContent1", _wasteCompanyPage);
        ActivePageContent2.Content = null;
        LogContentChange("ActivePageContent2", null);
        ActivePageContent3.Content = null;
        LogContentChange("ActivePageContent3", null);
        ActivePageContent4.Content = null;
        LogContentChange("ActivePageContent4", null);
        _wasteCompanyPage.LoadData();
        _bt1SaveAction = null;

        SetSubMenu("저장", "새로고침", "업소 등록", "엑셀 업로드", "보고서 생성", "통계 보기", "설정");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("WasteCompany");
    }

    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Analysis";
        _analysisPage ??= new AnalysisPage();
        ActivePageContent1.Content = _analysisPage;
        ActivePageContent2.Content = null;
        ActivePageContent4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("분석 시작", "새로고침", "데이터 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(380);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Analysis");
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Contract";

        if (_contractPage == null)
        {
            _contractPage = new ContractPage();
            _contractPage.ParentMainPage = this;
            _contractPage.DetailPanelChanged += panel =>
            {
                ActivePageContent2.Content = panel;
                LogContentChange("ActivePageContent2", panel);
            };
        }

        ActivePageContent1.Content = _contractPage;
        LogContentChange("ActivePageContent1", _contractPage);
        ActivePageContent2.Content = null;
        LogContentChange("ActivePageContent2", null);
        ActivePageContent3.Content = null;
        LogContentChange("ActivePageContent3", null);
        ActivePageContent4.Content = null;
        LogContentChange("ActivePageContent4", null);
        _contractPage.LoadData();
        _bt1SaveAction = _contractPage.SaveSelected;

        SetSubMenu("저장", "새로고침", "업체 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(350);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 8, lowerStar: 2);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Contract");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _contractPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    // ── 보수요청 ──────────────────────────────────────────────────────────────
    private void Repair_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Repair";
        _repairPage ??= new RepairPage();

        ActivePageContent1.Content = _repairPage.TreeControl;
        LogContentChange("ActivePageContent1", _repairPage.TreeControl);
        ActivePageContent2.Content = _repairPage.ListControl;
        LogContentChange("ActivePageContent2", _repairPage.ListControl);
        ActivePageContent3.Content = _repairPage.FormControl;
        LogContentChange("ActivePageContent3", _repairPage.FormControl);
        ActivePageContent4.Content = null;
        LogContentChange("ActivePageContent4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "승인", "반려", "완료", "삭제", "", "설정");
        SetLeftPanelWidth(220);
        // Content2(목록) 위, Content3(폼) 아래 30% 표시
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 7, lowerStar: 3);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Repair");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _repairPage.Refresh(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    // ── 견적/의뢰서 ───────────────────────────────────────────────────────
    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Quotation";

        // ── 초기화 순서 중요: CheckPanel → Page → NewPanel → DetailPanel → HistoryPanel

        // Content3: 분석항목 체크박스 (먼저 생성)
        _quotationCheckPanel ??= new QuotationCheckPanel();

        // Content4: 계약업체 목록
        _quotationPage ??= new QuotationPage();

        // Content2-A: 신규작성 폼
        if (_quotationNewPanel == null)
        {
            _quotationNewPanel = new QuotationNewPanel();
            // 저장 완료 → 히스토리 갱신 후 저장된 issue를 DetailPanel에 표시
            _quotationNewPanel.SaveCompleted += savedIssue =>
            {
                _lastShownIssue = savedIssue;
                _quotationHistoryPanel?.LoadData();
                // DB에서 최신 row 재조회 후 DetailPanel 갱신
                _quotationDetailPanel?.ShowIssue(savedIssue);
                ActivePageContent2.Content = _quotationDetailPanel;
                LogContentChange("ActivePageContent2", _quotationDetailPanel);
            };
        }
        // CheckPanel → NewPanel 연동
        _quotationCheckPanel.SelectionChanged -= OnCheckSelectionChanged;
        _quotationCheckPanel.SelectionChanged += OnCheckSelectionChanged;
        // 업체 → NewPanel 연동
        _quotationPage.CompanySelected -= OnCompanySelected;
        _quotationPage.CompanySelected += OnCompanySelected;

        // Content2-B: 세부내역 패널
        if (_quotationDetailPanel == null)
        {
            _quotationDetailPanel = new QuotationDetailPanel();
            _quotationDetailPanel.CheckPanel = _quotationCheckPanel;

            // 🥕 당근: 재활용 → NewPanel 에서 신규 번호·날짜로 작성
            _quotationDetailPanel.CarrotRequested += issue =>
            {
                var target = issue ?? _lastShownIssue;
                if (target == null) return;
                _lastShownIssue = target;
                _quotationNewPanel!.LoadFromIssue(target);
                ActivePageContent2.Content = _quotationNewPanel;
                LogContentChange("ActivePageContent2", _quotationNewPanel);
            };

            // ✏️ 오작성 수정: 기존 Id 덮어쓰기 — 메타 수정
            _quotationDetailPanel.CorrectRequested += issue =>
            {
                var target = issue ?? _lastShownIssue;
                if (target == null) return;
                _lastShownIssue = target;
                _quotationNewPanel!.LoadFromIssueCorrect(target);
                ActivePageContent2.Content = _quotationNewPanel;
                LogContentChange("ActivePageContent2", _quotationNewPanel);
            };

            // ESC 취소 → DetailPanel 복귀 + 마지막 issue 재표시
            _quotationNewPanel!.EscapeCancelled += () =>
            {
                if (_lastShownIssue != null)
                    _quotationDetailPanel?.ShowIssue(_lastShownIssue);
                ActivePageContent2.Content = _quotationDetailPanel;
                LogContentChange("ActivePageContent2", _quotationDetailPanel);
            };

            // 📋 의뢰서 작성: 편집 패널로 전환
            _quotationDetailPanel.OrderRequestEditRequested += (issue, samples, quotedItems) =>
            {
                _orderRequestEditPanel ??= new OrderRequestEditPanel();
                _orderRequestEditPanel.SubmitCompleted += () =>
                {
                    // 제출 완료 후 세부내역으로 복귀
                    ActivePageContent2.Content = _quotationDetailPanel;
                    LogContentChange("ActivePageContent2", _quotationDetailPanel);
                };
                _orderRequestEditPanel.Cancelled += () =>
                {
                    ActivePageContent2.Content = _quotationDetailPanel;
                    LogContentChange("ActivePageContent2", _quotationDetailPanel);
                };
                _orderRequestEditPanel.Load(issue, samples, quotedItems);
                ActivePageContent2.Content = _orderRequestEditPanel;
                LogContentChange("ActivePageContent2", _orderRequestEditPanel);
            };
        }
        _quotationDetailPanel.CheckPanel = _quotationCheckPanel;

        // Content1: 발행내역 + 분석의뢰내역 토글 트리
        if (_quotationHistoryPanel == null)
        {
            _quotationHistoryPanel = new QuotationHistoryPanel();

            // 견적발행내역 노드 선택
            _quotationHistoryPanel.IssueSelected += issue =>
            {
                _lastShownIssue = issue;
                _quotationDetailPanel!.ShowIssue(issue);
                ActivePageContent2.Content = _quotationDetailPanel;
                LogContentChange("ActivePageContent2", _quotationDetailPanel);
                ActivePageContent4.Content = _quotationPage;
                LogContentChange("ActivePageContent4", _quotationPage);
            };

            // 분석의뢰내역 탭으로 전환됨 → Content4: 의뢰 리스트 + TODO 패널
            _quotationHistoryPanel.AnalysisTabActivated += () =>
            {
                _analysisRequestListPanel ??= new AnalysisRequestListPanel();
                ActivePageContent4.Content = _analysisRequestListPanel;
                LogContentChange("ActivePageContent4", _analysisRequestListPanel);
            };

            // 견적발행내역 탭으로 복귀 → Content4: 계약업체 목록 + TODO 패널
            _quotationHistoryPanel.QuotationTabActivated += () =>
            {
                ActivePageContent4.Content = _quotationPage;
                LogContentChange("ActivePageContent4", _quotationPage);
            };

            // 분석의뢰내역 노드 선택
            _quotationHistoryPanel.AnalysisRequestSelected += record =>
            {
                _analysisRequestDetailPanel ??= new AnalysisRequestDetailPanel();
                _analysisRequestDetailPanel.CheckPanel = _quotationCheckPanel;
                _analysisRequestDetailPanel.ShowRecord(record);
                ActivePageContent2.Content = _analysisRequestDetailPanel;
                LogContentChange("ActivePageContent2", _analysisRequestDetailPanel);
                _analysisRequestListPanel ??= new AnalysisRequestListPanel();
                ActivePageContent4.Content = _analysisRequestListPanel;
                LogContentChange("ActivePageContent4", _analysisRequestListPanel);
                // Content4 트리뷰에 선택된 의뢰 추가
                _analysisRequestListPanel.AddRecord(record);
            };
        }

        ActivePageContent1.Content = _quotationHistoryPanel;
        LogContentChange("ActivePageContent1", _quotationHistoryPanel);
        ActivePageContent2.Content = _quotationDetailPanel;
        LogContentChange("ActivePageContent2", _quotationDetailPanel);
        ActivePageContent3.Content = _quotationCheckPanel;
        LogContentChange("ActivePageContent3", _quotationCheckPanel);
        ActivePageContent4.Content = _quotationPage;
        LogContentChange("ActivePageContent4", _quotationPage);
        _bt1SaveAction = null;

        _quotationHistoryPanel.LoadData();
        _quotationCheckPanel.LoadData();
        _quotationPage.LoadData();

        SetSubMenu("새로고침", "신규 작성", "삭제", "엑셀 내보내기", "인쇄", "", "설정");
        SetLeftPanelWidth(430);
        // Content2(세부내역) 50% : Content4(업체목록) 50%
        // 하단(Content3 분석항목) ≈ 23%  (13 : 4 → 76% : 24%)
        SetContentLayout(content2Star: 7, content4Star: 3, upperStar: 13, lowerStar: 4);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Quotation");
    }

    private void Purchase_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Purchase";
        _purchasePage ??= new PurchasePage();

        ActivePageContent1.Content = _purchasePage.TreeControl;
        LogContentChange("ActivePageContent1", _purchasePage.TreeControl);
        ActivePageContent2.Content = _purchasePage.ListControl;
        LogContentChange("ActivePageContent2", _purchasePage.ListControl);
        ActivePageContent3.Content = _purchasePage.FormControl;
        LogContentChange("ActivePageContent3", _purchasePage.FormControl);
        ActivePageContent4.Content = null;
        LogContentChange("ActivePageContent4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "엑셀 내보내기", "승인", "반려", "완료", "삭제", "설정");
        SetLeftPanelWidth(250);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 8, lowerStar: 2);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Purchase");
    }


    private void TestReport_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "TestReport";

        if (_testReportPage == null)
        {
            _testReportPage = new TestReportPage();
            _testReportPage.ResultListChanged += panel =>
            {
                ActivePageContent2.Content = panel;
                LogContentChange("ActivePageContent2", panel);
            };
            _testReportPage.EditPanelChanged += panel =>
            {
                ActivePageContent3.Content = panel;
                LogContentChange("ActivePageContent3", panel);
            };
        }

        ActivePageContent1.Content = _testReportPage;
        LogContentChange("ActivePageContent1", _testReportPage);
        ActivePageContent2.Content = null;
        LogContentChange("ActivePageContent2", null);
        ActivePageContent3.Content = null;
        LogContentChange("ActivePageContent3", null);

        // Content4: 출력 보관함 (Reports 폴더)
        _reportsPanel ??= new ReportsPanel();
        _reportsPanel.LoadFiles();
        ActivePageContent4.Content = _reportsPanel;
        LogContentChange("ActivePageContent4", _reportsPanel);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "CSV 저장", "삭제", "엑셀 출력", "PDF 출력", "일괄 엑셀", "일괄 PDF");

        SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 8.5, lowerStar: 1.5);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("TestReport");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _testReportPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Admin";
        ActivePageContent1.Content = null;
        ActivePageContent2.Content = null;
        ActivePageContent4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("사용자 관리", "권한 설정", "로그 확인", "백업하기", "시스템 설정", "통계", "종료");
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }

    private void Permission_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Permission";
        // TODO: 권한관리 페이지 구현
        ActivePageContent1.Content = null;
        ActivePageContent2.Content = null;
        ActivePageContent4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("저장", "새로고침", "삭제", "설정", "통계", "종료", "도움말");
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }


    // ══════════════════════════════════════════════════════════════════════
    //  서브메뉴 버튼
    // ══════════════════════════════════════════════════════════════════════

    private void SetSubMenu(string bt1, string bt2, string bt3,
                            string bt4, string bt5, string bt6, string bt7)
    {
        SetBtn(BT1, bt1); SetBtn(BT2, bt2); SetBtn(BT3, bt3);
        SetBtn(BT4, bt4); SetBtn(BT5, bt5);
        SetBtn(BT6, bt6); SetBtn(BT7, bt7);
    }

    private static void SetBtn(Avalonia.Controls.Button btn, string label)
    {
        btn.IsVisible = !string.IsNullOrWhiteSpace(label);
        btn.Content   = label;
    }

    private void BT1_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.Refresh();    break;
            case "TestReport": _testReportPage?.LoadData(); break;
            case "Repair":     _repairPage?.Refresh();      break;
            default: _bt1SaveAction?.Invoke();               break;
        }
    }

    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.LoadData();      break;
            case "WasteCompany": _wasteCompanyPage?.LoadData();   break;
            case "Contract":     _contractPage?.LoadData();       break;
            case "Purchase":     _purchasePage?.ExportCsv();      break;
            case "TestReport":   _testReportPage?.SaveCsv();      break;
            case "Quotation":
                // BT2 = 신규 작성 → Content2 를 NewPanel 로 교체
                _quotationNewPanel?.Clear();
                ActivePageContent2.Content = _quotationNewPanel;
                break;
            case "Repair":       _repairPage?.ApproveSelected();  break;
            default: Debug.WriteLine($"[{_currentMode}] BT2");   break;
        }
    }

    private void BT3_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":      _agentTreePage?.ShowAddPanel();           break;
            case "Contract":   _contractPage?.ShowAddPanel();            break;
            case "Purchase":   _purchasePage?.ApproveSelected();         break;
            case "Quotation":  _quotationPage?.LoadData(); _quotationHistoryPanel?.LoadData(); break;
            case "TestReport": _ = _testReportPage?.DeleteSampleAsync(); break;
            case "Repair":     _repairPage?.RejectSelected();            break;
            default: Debug.WriteLine($"[{_currentMode}] BT3");          break;
        }
    }

    private async void BT4_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":    if (_agentTreePage  != null) await _agentTreePage.DeleteSelectedAsync();  break;
            case "Contract": if (_contractPage   != null) await _contractPage.DeleteSelectedAsync();   break;
            case "Purchase": _purchasePage?.RejectSelected();   break;
            case "TestReport": _testReportPage?.PrintExcel();   break;
            case "Repair":   _repairPage?.CompleteSelected();   break;
            default: Debug.WriteLine($"[{_currentMode}] BT4"); break;
        }
    }

    private void BT5_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.CompleteSelected();  break;
            case "TestReport": _testReportPage?.PrintPdf();        break;
            case "Repair":     _repairPage?.DeleteSelected();      break;
            default: Debug.WriteLine($"[{_currentMode}] BT5");    break;
        }
    }

    private void BT6_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.DeleteSelected();        break;
            case "TestReport": _testReportPage?.BatchPrintExcel();     break;
            default: Debug.WriteLine($"[{_currentMode}] BT6");        break;
        }
    }

    private void BT7_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.ShowSettings(this);  break;
            case "TestReport": _testReportPage?.BatchPrintPdf();   break;
            default: Debug.WriteLine($"[{_currentMode}] BT7");    break;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  테마 (DynamicResource 전체 교체 방식)
    // ══════════════════════════════════════════════════════════════════════

    private void OnShowPasswordChanged(object? sender, RoutedEventArgs e)
    {
        if (tglShowPassword == null) return;
        ApplyTheme(tglShowPassword.IsChecked == true);
    }

    private void ApplyTheme(bool isDark)
    {
        // ── Avalonia 기본 테마 변형도 함께 변경 ──────────────────────────
        var variant = isDark ? ThemeVariant.Dark : ThemeVariant.Light;
        if (Application.Current is Application app)
            app.RequestedThemeVariant = variant;
        this.RequestedThemeVariant = variant;

        // ── Window.Resources의 DynamicResource 키 값 교체 ────────────────
        // 다크/라이트 팔레트 정의
        var palette = isDark ? DarkPalette() : LightPalette();
        foreach (var (key, color) in palette)
            this.Resources[key] = new Avalonia.Media.SolidColorBrush(
                Avalonia.Media.Color.Parse(color));
    }

    // ── 다크 팔레트 ──────────────────────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> DarkPalette() => new()
    {
        ["AppBg"]            = "#1e1e26",
        ["PanelBg"]          = "#39383f",
        ["PanelInnerBg"]     = "#2d2d35",
        ["MenuBarBg"]        = "#444444",
        ["MenuItemBg"]       = "#444444",
        ["MenuItemHover"]    = "#012800",
        ["SubMenuBg"]        = "#012800",
        ["SubMenuItemBg"]    = "#444444",   // ★ 서브메뉴 항목 통일
        ["SubMenuItemHover"] = "#023d00",
        ["SubBtnBg"]         = "#342f2f",
        ["SubBtnHover"]      = "#012800",
        ["SplitterColor"]    = "#020202",
        ["AppFg"]            = "#ffffff",
        ["FgMuted"]          = "#aaaaaa",
        ["FgHover"]          = "#fd0f0f",
        ["TreeFg"]           = "#e0e0e0",
        ["TreeSelBg"]        = "#014a00",
        ["TreeSelFg"]        = "#ffffff",
        ["GridHeaderBg"]     = "#2a2a32",
        ["GridRowBg"]        = "#35343c",
        ["GridRowAltBg"]     = "#2d2d35",
        ["InputBg"]          = "#2d2d35",
        ["InputBorder"]      = "#555566",
        ["InputFg"]          = "#e8e8e8",
    };

    // ── 라이트 팔레트 ─────────────────────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> LightPalette() => new()
    {
        ["AppBg"]            = "#f0f2f5",
        ["PanelBg"]          = "#ffffff",
        ["PanelInnerBg"]     = "#f8f9fb",
        ["MenuBarBg"]        = "#2d6a4f",
        ["MenuItemBg"]       = "#2d6a4f",
        ["MenuItemHover"]    = "#1b4332",
        ["SubMenuBg"]        = "#1b4332",
        ["SubMenuItemBg"]    = "#2d6a4f",   // ★ 라이트도 통일
        ["SubMenuItemHover"] = "#145a32",
        ["SubBtnBg"]         = "#e2e8f0",
        ["SubBtnHover"]      = "#cbd5e1",
        ["SplitterColor"]    = "#cbd5e1",
        ["AppFg"]            = "#1a1a2e",
        ["FgMuted"]          = "#64748b",
        ["FgHover"]          = "#dc2626",
        ["TreeFg"]           = "#1e293b",
        ["TreeSelBg"]        = "#bbf7d0",
        ["TreeSelFg"]        = "#14532d",
        ["GridHeaderBg"]     = "#e2e8f0",
        ["GridRowBg"]        = "#ffffff",
        ["GridRowAltBg"]     = "#f8fafc",
        ["InputBg"]          = "#ffffff",
        ["InputBorder"]      = "#94a3b8",
        ["InputFg"]          = "#1e293b",
    };

    // ── Quotation 이벤트 핸들러 (중복 구독 방지용 named handler) ─────────
    private void OnCheckSelectionChanged(System.Collections.Generic.List<ETA.Models.AnalysisItem> items)
        => _quotationNewPanel?.SetSelectedAnalytes(items);

    private void OnCompanySelected(ETA.Models.Contract company)
        => _quotationNewPanel?.SetCompany(company);

    private void LogContentChange(string contentName, object? content)
    {
        Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] ContentChange: {contentName} = {content?.GetType().Name ?? "null"}");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  창 위치/레이아웃 저장 및 복원
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 현재 모드의 레이아웃 정보 저장
    /// </summary>
    private void SaveCurrentModeLayout()
    {
        if (_positionManager == null || string.IsNullOrEmpty(_currentMode) || _currentMode == "None")
            return;

        try
        {
            string modeKey = LayoutStorageModePrefix + _currentMode;
            var layout = new PageLayoutInfo();

            // 윈도우 위치/크기
            layout.WindowX = this.Position.X;
            layout.WindowY = this.Position.Y;
            layout.WindowWidth = this.Width;
            layout.WindowHeight = this.Height;

            // 왼쪽 패널 너비 조회
            var mainGrid = this.FindControl<Grid>("MainSplitGrid");
            if (mainGrid != null && mainGrid.ColumnDefinitions.Count > 0)
            {
                var colDef = mainGrid.ColumnDefinitions[0];
                if (colDef.Width.IsAbsolute)
                    layout.LeftPanelWidth = colDef.Width.Value;
            }

            // 오른쪽 상단/하단 분할 비율
            var rightGrid = this.FindControl<Grid>("RightSplitGrid");
            if (rightGrid != null && rightGrid.RowDefinitions.Count >= 3)
            {
                layout.UpperStar = rightGrid.RowDefinitions[0].Height.Value;
                layout.LowerStar = rightGrid.RowDefinitions[2].Height.Value;
            }

            // Content2/Content4 분할 비율
            var upperGrid = this.FindControl<Grid>("UpperContentGrid");
            if (upperGrid != null && upperGrid.ColumnDefinitions.Count >= 3)
            {
                layout.Content2Star = upperGrid.ColumnDefinitions[0].Width.Value;
                layout.Content4Star = upperGrid.ColumnDefinitions[2].Width.Value;
            }

            layout.SavedAt = DateTime.Now;
            _positionManager.SavePageLayout(modeKey, layout);
            System.Diagnostics.Debug.WriteLine($"[MainPage] 저장: {modeKey} - {layout}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MainPage] 레이아웃 저장 오류: {ex.Message}");
        }
    }

    /// <summary>
    /// 특정 모드의 저장된 레이아웃 정보 복원
    /// </summary>
    private void RestoreModeLayout(string modeName)
    {
        if (_positionManager == null)
            return;

        try
        {
            string modeKey = LayoutStorageModePrefix + modeName;
            var layout = _positionManager.GetPageLayout(modeKey);
            if (layout == null)
            {
                System.Diagnostics.Debug.WriteLine($"[MainPage] 복원할 레이아웃 없음: {modeKey}");
                return;
            }

            // 왼쪽 패널 너비 복원
            if (layout.LeftPanelWidth > 0)
                SetLeftPanelWidth(layout.LeftPanelWidth);

            // 레이아웃 비율 복원
            SetContentLayout(
                content2Star: layout.Content2Star,
                content4Star: layout.Content4Star,
                upperStar: layout.UpperStar,
                lowerStar: layout.LowerStar);

            System.Diagnostics.Debug.WriteLine($"[MainPage] 복원: {modeKey} - {layout}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MainPage] 레이아웃 복원 오류: {ex.Message}");
        }
    }

    /// <summary>
    /// 사용자 변경 시 WindowPositionManager 업데이트
    /// (로그인 후 호출)
    /// </summary>
    public void UpdateCurrentUser(string newUserId)
    {
        if (!string.IsNullOrWhiteSpace(newUserId))
        {
            CurrentUserManager.Instance.SetCurrentUser(newUserId);
            // 새 사용자로 WindowPositionManager 재초기화
            _positionManager = new WindowPositionManager(newUserId);
            System.Diagnostics.Debug.WriteLine($"[MainPage] 사용자 변경: {newUserId}");
        }
    }

    /// <summary>
    /// 로그 파일 경로 조회 (디버깅용)
    /// </summary>
    public string GetPositionLogFilePath()
    {
        return _positionManager?.GetLogFilePath() ?? "Unknown";
    }

    private void TEST_Click(object? sender, RoutedEventArgs e) { }
}
