using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;

using ETA.ViewModels;           // ← MainViewModel 오류 해결
using ETA.Views.Pages;          // AgentTreePage 등

namespace ETA.Views;

public partial class MainPage : Window
{
    private readonly AnalysisPage _analysisPage = new();
    private readonly ContractPage _contractPage = new();
    private readonly QuotationPage _quotationPage = new();
    private readonly AgentTreePage _agentTreePage = new();

    public MainPage()
    {
        InitializeComponent();
        
        DataContext = new MainWindowViewModel();   // ← 필수!
        
        ActivePageContent.Content = _analysisPage;  // 기본 페이지
    }

    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        ActivePageContent.Content = _analysisPage;
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        ActivePageContent.Content = _contractPage;
    }

    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        ActivePageContent.Content = _quotationPage;
    }

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        ActivePageContent.Content = _agentTreePage;

// ★★★ 핵심: 데이터 강제 로드 (이 줄만 추가) ★★★
    _agentTreePage.LoadData();

    // cmbCategory 안전하게 찾기 (기존 코드 유지)
    var cmb = this.FindControl<ComboBox>("cmbCategory");
    if (cmb != null) cmb.IsVisible = false;
    }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        // 관리자 화면 (나중에 구현)
    }

    private void OnShowPasswordChanged(object? sender, RoutedEventArgs e)
    {
        // 기존 로직
    }

    private void TEST_Click(object? sender, RoutedEventArgs e)
    {
        // 테스트용
    }
}