using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Styling;
using ETA.ViewModels;
using ETA.Views.Pages;
using System;
using System.Diagnostics;

namespace ETA.Views;

public partial class MainPage : Window
{
    private string _currentMode = "None";

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
    }

    // ══════════════════════════════════════════════════════════════════════
    //  로그 기록 헬퍼
    // ══════════════════════════════════════════════════════════════════════

    private void LogContentChange(string contentName, object? content)
    {
        string contentType = content?.GetType().Name ?? "null";
        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {contentName} changed to {contentType}");
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
        ActivePageContent4.Content = null;
        LogContentChange("ActivePageContent4", null);
        _agentTreePage.LoadData();
        _bt1SaveAction = _agentTreePage.SaveSelected;

        SetSubMenu("저장", "새로고침", "직원 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);
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

    private void TEST_Click(object? sender, RoutedEventArgs e) { }
}
