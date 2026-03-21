using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using ETA.ViewModels;
using ETA.Views.Pages;
using System.Diagnostics;   // ← Debug.WriteLine 때문에 필수

namespace ETA.Views;

public partial class MainPage : Window
{
    private string currentMode = "None";   // ← 여기로 이동 (class 안쪽!)

    private readonly AnalysisPage _analysisPage = new();
    private readonly ContractPage _contractPage = new();
    private readonly QuotationPage _quotationPage = new();
    private readonly AgentTreePage _agentTreePage = new();
    private readonly WasteCompanyPage _wastecompanyTreePage = new();

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ActivePageContent1.Content = _analysisPage; // 기본 페이지
        ApplyTheme(true);
    }

    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "Analysis";
        ActivePageContent1.Content = _analysisPage;
        BT1.Content = "분석 시작"; BT2.Content = "새로고침"; BT3.Content = "데이터 추가";
        BT4.Content = "선택 삭제"; BT5.Content = "엑셀 내보내기"; BT6.Content = "인쇄"; BT7.Content = "설정";
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "Contract";
        ActivePageContent1.Content = _contractPage;
        BT1.Content = "계약 등록"; BT2.Content = "새로고침"; BT3.Content = "계약서 보기";
        BT4.Content = "삭제"; BT5.Content = "엑셀 내보내기"; BT6.Content = "인쇄"; BT7.Content = "설정";
    }

    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "Quotation";
        ActivePageContent1.Content = _quotationPage;
        BT1.Content = "견적 등록"; BT2.Content = "새로고침"; BT3.Content = "의뢰서 보기";
        BT4.Content = "삭제"; BT5.Content = "엑셀 내보내기"; BT6.Content = "인쇄"; BT7.Content = "설정";
    }

    private void WasteCompany_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "WasteCompany";
        ActivePageContent1.Content = _wastecompanyTreePage;
        _wastecompanyTreePage.LoadData();

        BT1.Content = "업소 등록"; BT2.Content = "비용 계산"; BT3.Content = "목록 새로고침";
        BT4.Content = "엑셀 업로드"; BT5.Content = "보고서 생성"; BT6.Content = "통계 보기"; BT7.Content = "설정";

       
    }

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "Agent";
        ActivePageContent1.Content = _agentTreePage;

        BT1.Content = "전체 저장"; BT2.Content = "새로고침"; BT3.Content = "직원 추가";
        BT4.Content = "선택 삭제"; BT5.Content = "엑셀 내보내기"; BT6.Content = "인쇄"; BT7.Content = "설정";

        _agentTreePage.LoadData();


    }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        currentMode = "Admin";
        ActivePageContent1.Content = null;

        BT1.Content = "사용자 관리"; BT2.Content = "권한 설정"; BT3.Content = "로그 확인";
        BT4.Content = "백업하기"; BT5.Content = "시스템 설정"; BT6.Content = "통계"; BT7.Content = "종료";
    }

    private void OnShowPasswordChanged(object? sender, RoutedEventArgs e)
    {
        if (tglShowPassword == null) return;
        bool isDarkMode = tglShowPassword.IsChecked == true;
        ApplyTheme(isDarkMode);
    }

    private void ApplyTheme(bool isDarkMode)
    {
        if (Application.Current is not Application app) return;
        var newTheme = isDarkMode ? ThemeVariant.Dark : ThemeVariant.Light;
        app.RequestedThemeVariant = newTheme;
        this.RequestedThemeVariant = newTheme;
    }

    private void TEST_Click(object? sender, RoutedEventArgs e)
    {
        // 테스트용
    }

    // ====================== BT1 ~ BT7 실제 기능 ======================
    private void BT1_Click(object? sender, RoutedEventArgs e)
    {
        switch (currentMode)
        {
            case "Agent": Debug.WriteLine("관리자 - 사용자 추가"); break;
            case "WasteCompany": Debug.WriteLine("폐수배출업소 - 업소 등록 실행"); break;
            case "Analysis": Debug.WriteLine("분석 - 분석 시작"); break;
            case "Contract": Debug.WriteLine("계약 - 계약 등록"); break;
            case "Quotation": Debug.WriteLine("견적 - 견적 등록"); break;
            case "Admin": Debug.WriteLine("관리자 - 사용자 추가"); break;
            default: Debug.WriteLine("BT1: 현재 모드에서 기능 없음"); break;
        }
    }

    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (currentMode)
        {
            case "Agent": _agentTreePage.LoadData(); break;
            default: Debug.WriteLine(currentMode + " - BT2 클릭"); break;
        }
    }

    private void BT3_Click(object? sender, RoutedEventArgs e) => Debug.WriteLine(currentMode + " - BT3 클릭");
    private void BT4_Click(object? sender, RoutedEventArgs e) => Debug.WriteLine(currentMode + " - BT4 클릭");
    private void BT5_Click(object? sender, RoutedEventArgs e) => Debug.WriteLine(currentMode + " - BT5 클릭");
    private void BT6_Click(object? sender, RoutedEventArgs e) => Debug.WriteLine(currentMode + " - BT6 클릭");
    private void BT7_Click(object? sender, RoutedEventArgs e) => Debug.WriteLine(currentMode + " - BT7 클릭");
}