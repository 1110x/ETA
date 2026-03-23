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
    private AnalysisPage?     _analysisPage;
    private ContractPage?     _contractPage;
    private QuotationPage?    _quotationPage;
    private AgentTreePage?    _agentTreePage;
    private WasteCompanyPage? _wasteCompanyPage;
    private PurchasePage?      _purchasePage;
    private TestReportPage?   _testReportPage;

    // ── 서브메뉴 BT1 의 "현재 모드 저장 액션"  ──────────────────────────────
    // 각 페이지가 자신의 저장 로직을 여기에 등록한다
    private System.Action? _bt1SaveAction;

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ApplyTheme(true);
    }

    // ── 메뉴 클릭 ────────────────────────────────────────────────────────────

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Agent";

        // Lazy 생성 + 이벤트 연결 (최초 1회만)
        if (_agentTreePage == null)
        {
            _agentTreePage = new AgentTreePage();

            // 선택 → ActivePageContent2 에 상세 패널 표시
            _agentTreePage.DetailPanelChanged += panel =>
            {
                ActivePageContent2.Content = panel;
            };
        }

        ActivePageContent1.Content = _agentTreePage;
        ActivePageContent2.Content = null;          // 선택 전 오른쪽 초기화

        _agentTreePage.LoadData();

        // BT1 저장 액션 = AgentTreePage.SaveSelected
        _bt1SaveAction = _agentTreePage.SaveSelected;

        SetSubMenu("저장", "새로고침", "직원 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
    }

    private void WasteCompany_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteCompany";

        if (_wasteCompanyPage == null)
            _wasteCompanyPage = new WasteCompanyPage();

        ActivePageContent1.Content = _wasteCompanyPage;
        ActivePageContent2.Content = null;

        _wasteCompanyPage.LoadData();

        // WasteCompanyPage 에도 같은 패턴으로 DetailPanelChanged / SaveSelected 추가 예정
        _bt1SaveAction = null;  // TODO: _wasteCompanyPage.SaveSelected

        SetSubMenu("저장", "새로고침", "업소 등록", "엑셀 업로드", "보고서 생성", "통계 보기", "설정");
    }

    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Analysis";
        _analysisPage ??= new AnalysisPage();
        ActivePageContent1.Content = _analysisPage;
        ActivePageContent2.Content = null;
        _bt1SaveAction = null;
        SetSubMenu("분석 시작", "새로고침", "데이터 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Contract";

        if (_contractPage == null)
        {
            _contractPage = new ContractPage();
            _contractPage.DetailPanelChanged += panel =>
            {
                ActivePageContent2.Content = panel;
            };
        }

        ActivePageContent1.Content = _contractPage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;

        _contractPage.LoadData();
        _bt1SaveAction = _contractPage.SaveSelected;

        SetSubMenu("저장", "새로고침", "업체 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
    }

    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Quotation";
        _quotationPage ??= new QuotationPage();
        ActivePageContent1.Content = _quotationPage;
        ActivePageContent2.Content = null;
        _bt1SaveAction = null;
        SetSubMenu("견적 등록", "새로고침", "의뢰서 보기", "삭제", "엑셀 내보내기", "인쇄", "설정");
    }

    private void Purchase_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Purchase";

        _purchasePage ??= new PurchasePage();

        ActivePageContent1.Content = _purchasePage.TreeControl;
        ActivePageContent2.Content = _purchasePage.ListControl;
        ActivePageContent3.Content = _purchasePage.FormControl;
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "엑셀 내보내기", "승인", "반려", "완료", "삭제", "설정");
    }

    private void TestReport_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "TestReport";

        if (_testReportPage == null)
        {
            _testReportPage = new TestReportPage();
            _testReportPage.ResultListChanged += panel =>
                ActivePageContent2.Content = panel;
            _testReportPage.EditPanelChanged  += panel =>
                ActivePageContent3.Content = panel;
        }

        ActivePageContent1.Content = _testReportPage;
        ActivePageContent2.Content = null;
        ActivePageContent3.Content = null;
        _bt1SaveAction             = null;

        SetSubMenu("새로고침", "📄 엑셀 출력", "📋 PDF 저장", "📦 일괄 엑셀", "📦 일괄 PDF", "☑ 전체선택", "☐ 전체해제");

        // Content 꽂은 뒤 레이아웃 완료 이후 LoadData 실행
        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _testReportPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Admin";
        ActivePageContent1.Content = null;
        ActivePageContent2.Content = null;
        _bt1SaveAction = null;
        SetSubMenu("사용자 관리", "권한 설정", "로그 확인", "백업하기", "시스템 설정", "통계", "종료");
    }

    // ── 서브메뉴 버튼 텍스트 일괄 설정 ─────────────────────────────────────
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

    // ── 서브메뉴 버튼 클릭 ──────────────────────────────────────────────────

    // BT1: 모드별 1번 액션 (Agent=저장, Purchase=새로고침)
    private void BT1_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":    _purchasePage?.Refresh();  break;
            case "TestReport":  _testReportPage?.LoadData(); break;
            default:
                if (_bt1SaveAction != null) _bt1SaveAction.Invoke();
                else Debug.WriteLine($"[{_currentMode}] BT1: 액션 없음");
                break;
        }
    }

    // BT2: 모드별 2번 액션
    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.LoadData();    break;
            case "WasteCompany": _wasteCompanyPage?.LoadData(); break;
            case "Contract":     _contractPage?.LoadData();      break;
            case "Purchase":     _purchasePage?.ExportCsv();     break;
            case "TestReport":   _testReportPage?.PrintExcel();  break;
            // BT1=출력, BT2=CSV저장이므로 새로고침은 BT1 에 없음
            default: Debug.WriteLine($"[{_currentMode}] BT2"); break;
        }
    }

    private void BT3_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":    _agentTreePage?.ShowAddPanel();    break;
            case "Contract":    _contractPage?.ShowAddPanel();       break;
            case "Purchase":    _purchasePage?.ApproveSelected();    break;
            case "TestReport":   _testReportPage?.PrintPdf();    break;
            default: Debug.WriteLine($"[{_currentMode}] BT3: {BT3.Content}"); break;
        }
    }

    private async void BT4_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":      if (_agentTreePage != null) await _agentTreePage.DeleteSelectedAsync(); break;
            case "TestReport": _testReportPage?.BatchPrintExcel(); break;
            case "Contract": if (_contractPage != null) await _contractPage.DeleteSelectedAsync(); break;
            case "Purchase": _purchasePage?.RejectSelected();  break;
            default: Debug.WriteLine($"[{_currentMode}] BT4: {BT4.Content}"); break;
        }
    }

    private void BT5_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":   _purchasePage?.CompleteSelected();       break;
            case "TestReport": _testReportPage?.BatchPrintPdf();          break;
            default: Debug.WriteLine($"[{_currentMode}] BT5: {BT5.Content}"); break;
        }
    }

    private void BT6_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":    _purchasePage?.DeleteSelected();              break;
            case "TestReport":  _testReportPage?.SelectAllSamples(true);     break;
            default: Debug.WriteLine($"[{_currentMode}] BT6: {BT6.Content}"); break;
        }
    }

    private void BT7_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":    _purchasePage?.ShowSettings(this);           break;
            case "TestReport":  _testReportPage?.SelectAllSamples(false);    break;
            default: Debug.WriteLine($"[{_currentMode}] BT7: {BT7.Content}"); break;
        }
    }

    // ── 테마 ────────────────────────────────────────────────────────────────
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