using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.ObjectModel;
using ETA.ViewModels;
using System.Linq;
using Avalonia;   // ← 이게 핵심! Thickness는 Avalonia 네임스페이스에 있습니다.
namespace ETA.Views;

using Avalonia.Input;
using Avalonia.Media;

public partial class Page1 : Window
{
    private readonly ObservableCollection<AnalysisItem> _analysisItems = new();
    private readonly ObservableCollection<Contract> _contracts = new();
    private readonly ObservableCollection<ContractPrice> _quotationItems = new();

    public Page1()
    {
        InitializeComponent();

        // 초기 바인딩
        AnalysisGrid.ItemsSource = _analysisItems;
        ContractGrid.ItemsSource = _contracts;
        QuotationGrid.ItemsSource = _quotationItems;
        // DataContext = new AnalysisViewModel();  // ← new 뒤에 AnalysisViewModel (단수!)
        // 처음에는 메인 패널만 보이게
        HideAllPanels();
        MainPanel.IsVisible = true;


    }

    private void LoadAnalysisData()
    {
        Console.WriteLine("[LoadAnalysisData] 시작");
        var items = AnalysisService.GetAllItems();
        Console.WriteLine($"[Analysis] 항목 수: {items.Count}");
        _analysisItems.Clear();
        foreach (var item in items) _analysisItems.Add(item);
        var categories = _analysisItems
    .Select(item => item.Category)
    .Where(c => !string.IsNullOrEmpty(c))
    .Distinct()
    .OrderBy(c => c)
    .ToList();

        cmbCategory.ItemsSource = categories;

        Console.WriteLine($"ComboBox에 넣은 고유 Category 개수: {categories.Count}");
    }

    private void LoadContractData()
    {
        var items = ContractService.GetAllContracts();
        Console.WriteLine($"[Contract] 항목 수: {items.Count}");
        _contracts.Clear();
        foreach (var item in items) _contracts.Add(item);
    }

    private void LoadQuotationData()
    {
        MainPanel.IsVisible = true;
        try
        {
            Console.WriteLine("[Quotation] 로드 시작 - DB 연결 시도 중");

            var itemsFromDb = ContractPriceService.GetAllContractPrices();

            Console.WriteLine($"[Quotation] 성공 - 읽은 행 수: {itemsFromDb.Count}");

            _quotationItems.Clear();
            foreach (var item in itemsFromDb)
            {
                _quotationItems.Add(item);
            }

            Console.WriteLine("[Quotation] 데이터 바인딩 완료");
        }
        catch (Exception ex)
        {
            Console.WriteLine("=====================================");
            Console.WriteLine("[CRASH] 분석단가 로드 중 치명적 오류 발생");
            Console.WriteLine($"오류 메시지: {ex.Message}");
            Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            Console.WriteLine("=====================================");
        }
    }
    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        Console.WriteLine("[클릭] 분석관리");
        HideAllPanels();
        MainPanel.IsVisible = true;
        AnalysisGrid.IsVisible = true;
        cmbCategory.IsVisible = true;  // 카테고리 콤보박스도 같이 보이게
        Btn13.IsVisible = true;
        LoadAnalysisData();
        AnalysisGrid.Margin = new Thickness(10, 50, 10, 10);
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        Console.WriteLine("[클릭] 업체관리");
        HideAllPanels();
        MainPanel.IsVisible = true;
        ContractGrid.IsVisible = true;
        cmbCategory.IsVisible = false;  // 카테고리 콤보박스도 같이 보이게
        Btn13.IsVisible = false;
        LoadContractData();
    }

    // private void Quotation_Click(object? sender, RoutedEventArgs e)
    // {
    //     Console.WriteLine("[클릭] 분석단가 (견적정보)");
    //     HideAllPanels();
    //     QuotationGrid.IsVisible = true;
    //    LoadQuotationData();
    // }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        var main = new MainWindow();
        main.Show();

        // 관리자 전용 화면이 있다면 여기서 켜기
        // 예: AdminPanel.IsVisible = true;
        // 지금은 아무것도 안 열리고 로그만 남김
        cmbCategory.IsVisible = false;  // 카테고리 콤보박스도 같이 보이게
        Btn13.IsVisible = false;
    }

    private void HideAllPanels()
    {
        MainPanel.IsVisible = true;
        AnalysisGrid.IsVisible = false;
        ContractGrid.IsVisible = false;
        QuotationGrid.IsVisible = false;   // ← 이 줄이 핵심! 누락되면 깜빡임 발생
    }
    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            Console.WriteLine("[Quotation_Click] 버튼 클릭 감지됨");
            HideAllPanels();
            QuotationGrid.IsVisible = true;
            LoadQuotationData();
            cmbCategory.IsVisible = false;  // 카테고리 콤보박스도 같이 보이게
            Btn13.IsVisible = false;
        }
        catch (Exception ex)
        {
            Console.WriteLine("[Quotation_Click] 이벤트 자체에서 오류");
            Console.WriteLine($"오류: {ex.Message}\n{ex.StackTrace}");
        }
    }
    private void OnShowPasswordChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {

    }
    private void CmbCategory_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox comboBox && comboBox.SelectedItem is string selectedCategory)
        {
            Console.WriteLine($"선택된 카테고리: {selectedCategory}");

            // 여기서 선택된 카테고리에 따라 DataGrid 필터링하거나 다른 작업
            // 예: AnalysisGrid.ItemsSource를 필터링
            var filtered = _analysisItems.Where(item => item.Category == selectedCategory).ToList();
            AnalysisGrid.ItemsSource = filtered;
        }
    }

    private void TEST_Click(object? sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("테스트 출력");
    }


}