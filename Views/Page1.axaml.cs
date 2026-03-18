using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.ObjectModel;

namespace ETA.Views;

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
    }

    private void LoadContractData()
    {
        Console.WriteLine("[LoadContractData] 시작");
        var items = ContractService.GetAllContracts();
        Console.WriteLine($"[Contract] 항목 수: {items.Count}");
        _contracts.Clear();
        foreach (var item in items) _contracts.Add(item);
    }

    private void LoadQuotationData()
    {
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
        AnalysisGrid.IsVisible = true;
        LoadAnalysisData();
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        Console.WriteLine("[클릭] 업체관리");
        HideAllPanels();
        ContractGrid.IsVisible = true;
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
        Console.WriteLine("[클릭] 관리자");
        HideAllPanels();
        // 관리자 전용 화면이 있다면 여기서 켜기
        // 예: AdminPanel.IsVisible = true;
        // 지금은 아무것도 안 열리고 로그만 남김
    }

    private void HideAllPanels()
    {
        MainPanel.IsVisible = false;
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
        }
        catch (Exception ex)
        {
            Console.WriteLine("[Quotation_Click] 이벤트 자체에서 오류");
            Console.WriteLine($"오류: {ex.Message}\n{ex.StackTrace}");
        }
    }
}