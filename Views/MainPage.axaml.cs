using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Styling;
using ETA.ViewModels;
using ETA.Views.Pages;
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
    private TestReportPage?    _testReportPage;

    // ── 견적/의뢰서 전용 4-패널 ──────────────────────────────────────────
    // Content1: 발행내역 트리  Content2: 신규작성 폼
    // Content3: 분석항목 체크  Content4: 계약업체 목록
    private QuotationHistoryPanel? _quotationHistoryPanel;
    private QuotationDetailPanel?  _quotationDetailPanel;   // Content2: 세부내역
    private QuotationNewPanel?     _quotationNewPanel;      // 신규작성 (서브메뉴 BT)
    private QuotationCheckPanel?   _quotationCheckPanel;
    private QuotationPage?         _quotationPage;

    private System.Action? _bt1SaveAction;

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ApplyTheme(true);
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
                ActivePageContent2.Content = panel;
        }

        ActivePageContent1.Content = _agentTreePage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;
        ActivePageContent4.Content = null;
        _agentTreePage.LoadData();
        _bt1SaveAction = _agentTreePage.SaveSelected;

        SetSubMenu("저장", "새로고침", "직원 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }

    private void WasteCompany_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteCompany";
        if (_wasteCompanyPage == null) _wasteCompanyPage = new WasteCompanyPage();

        ActivePageContent1.Content = _wasteCompanyPage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;
        ActivePageContent4.Content = null;
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
        SetLeftPanelWidth(260);
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
                ActivePageContent2.Content = panel;
        }

        ActivePageContent1.Content = _contractPage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;
        ActivePageContent4.Content = null;
        _contractPage.LoadData();
        _bt1SaveAction = _contractPage.SaveSelected;

        SetSubMenu("저장", "새로고침", "업체 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(330);
        SetContentLayout(content2Star: 7, content4Star: 3, upperStar: 1, lowerStar: 0);

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _contractPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    // ── 견적/의뢰서 ───────────────────────────────────────────────────────
    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Quotation";

        // Content1: 발행내역 트리
        if (_quotationHistoryPanel == null)
        {
            _quotationHistoryPanel = new QuotationHistoryPanel();
            // 트리 항목 선택 → Content2 세부내역 갱신
            _quotationHistoryPanel.IssueSelected += issue =>
            {
                _quotationDetailPanel ??= new QuotationDetailPanel();
                _quotationDetailPanel.ShowIssue(issue);
                ActivePageContent2.Content = _quotationDetailPanel;
            };
        }

        // Content2: 초기 상태 (발행건 선택 전 빈 세부내역 패널)
        _quotationDetailPanel ??= new QuotationDetailPanel();
        // CheckPanel 참조 주입 (체크 동기화용)
        _quotationDetailPanel.CheckPanel = _quotationCheckPanel;

        // Content3: 분석항목 체크박스
        if (_quotationCheckPanel == null)
        {
            _quotationCheckPanel = new QuotationCheckPanel();
            _quotationCheckPanel.SelectionChanged += items =>
                _quotationNewPanel?.SetSelectedAnalytes(items);
        }

        // Content4: 계약업체 목록
        if (_quotationPage == null)
        {
            _quotationPage = new QuotationPage();
            _quotationPage.CompanySelected += company =>
                _quotationNewPanel?.SetCompany(company);
        }

        ActivePageContent1.Content = _quotationHistoryPanel;
        ActivePageContent2.Content = _quotationDetailPanel;
        ActivePageContent3.Content = _quotationCheckPanel;
        ActivePageContent4.Content = _quotationPage;
        _bt1SaveAction = null;

        _quotationHistoryPanel.LoadData();
        _quotationCheckPanel.LoadData();
        _quotationPage.LoadData();

        SetSubMenu("새로고침", "신규 작성", "삭제", "엑셀 내보내기", "인쇄", "", "설정");
        SetLeftPanelWidth(260);
        // Content2(세부내역) 50% : Content4(업체목록) 50%
        // 하단(Content3 분석항목) ≈ 23%  (13 : 4 → 76% : 24%)
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 13, lowerStar: 4);
    }

    private void Purchase_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Purchase";
        _purchasePage ??= new PurchasePage();

        ActivePageContent1.Content = _purchasePage.TreeControl;
        ActivePageContent2.Content = _purchasePage.ListControl;
        ActivePageContent3.Content = _purchasePage.FormControl;
        ActivePageContent4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "엑셀 내보내기", "승인", "반려", "완료", "삭제", "설정");
        SetLeftPanelWidth(250);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 4, lowerStar: 1);
    }

    private void TestReport_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "TestReport";

        if (_testReportPage == null)
        {
            _testReportPage = new TestReportPage();
            _testReportPage.ResultListChanged += panel =>
                ActivePageContent2.Content = panel;
            _testReportPage.EditPanelChanged += panel =>
                ActivePageContent3.Content = panel;
        }

        ActivePageContent1.Content = _testReportPage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;
        ActivePageContent4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "CSV 저장", "삭제", "", "", "", "");
        BT4.IsVisible = false; BT5.IsVisible = false;
        BT6.IsVisible = false; BT7.IsVisible = false;

        SetContentLayout(content2Star: 6, content4Star: 4, upperStar: 7, lowerStar: 3);

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
        BT1.IsVisible = true; BT2.IsVisible = true; BT3.IsVisible = true;
        BT4.IsVisible = true; BT5.IsVisible = true;
        BT6.IsVisible = true; BT7.IsVisible = true;
        BT1.Content = bt1; BT2.Content = bt2; BT3.Content = bt3;
        BT4.Content = bt4; BT5.Content = bt5;
        BT6.Content = bt6; BT7.Content = bt7;
    }

    private void BT1_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.Refresh();    break;
            case "TestReport": _testReportPage?.LoadData(); break;
            case "Quotation": _quotationHistoryPanel?.LoadData(); break; // 새로고침
            default:
                _bt1SaveAction?.Invoke();
                break;
        }
    }

    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.LoadData();    break;
            case "WasteCompany": _wasteCompanyPage?.LoadData(); break;
            case "Contract":     _contractPage?.LoadData();     break;
            case "Purchase":     _purchasePage?.ExportCsv();    break;
            case "TestReport":   _testReportPage?.SaveCsv();    break;
            case "Quotation": break;
            default: Debug.WriteLine($"[{_currentMode}] BT2"); break;
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
            default: Debug.WriteLine($"[{_currentMode}] BT3");          break;
        }
    }

    private async void BT4_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":    if (_agentTreePage  != null) await _agentTreePage.DeleteSelectedAsync();  break;
            case "Contract": if (_contractPage   != null) await _contractPage.DeleteSelectedAsync();   break;
            case "Purchase": _purchasePage?.RejectSelected(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT4"); break;
        }
    }

    private void BT5_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase": _purchasePage?.CompleteSelected(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT5"); break;
        }
    }

    private void BT6_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase": _purchasePage?.DeleteSelected(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT6"); break;
        }
    }

    private void BT7_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase": _purchasePage?.ShowSettings(this); break;
            default: Debug.WriteLine($"[{_currentMode}] BT7"); break;
        }
    }


    // ══════════════════════════════════════════════════════════════════════
    //  테마
    // ══════════════════════════════════════════════════════════════════════

    private void OnShowPasswordChanged(object? sender, RoutedEventArgs e)
    {
        if (tglShowPassword == null) return;
        ApplyTheme(tglShowPassword.IsChecked == true);
    }

    private void ApplyTheme(bool isDarkMode)
    {
        if (Application.Current is not Application app) return;
        var theme = isDarkMode ? ThemeVariant.Dark : ThemeVariant.Light;
        app.RequestedThemeVariant = theme;
        this.RequestedThemeVariant = theme;
    }

    private void TEST_Click(object? sender, RoutedEventArgs e) { }
}
