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

    public Page1()
    {
        InitializeComponent();
        AnalysisGrid.ItemsSource = _analysisItems;
    }

 private void LoadAnalysisData()
{
    var itemsFromDb = AnalysisService.GetAllItems();

    Console.WriteLine($"[Page1] DB에서 읽은 항목 수: {itemsFromDb.Count}");
    if (itemsFromDb.Count > 0)
    {
        var first = itemsFromDb[0];
        Console.WriteLine($"[Page1] 첫 번째 항목 예시: {first.Category} / {first.Analyte} / {first.unit}");
    }
    else
    {
        Console.WriteLine("[Page1] 데이터가 0개입니다.");
    }

    _analysisItems.Clear();
    foreach (var item in itemsFromDb)
    {
        _analysisItems.Add(item);
    }

    // 필요 시 강제 새로고침 (거의 안 써도 됨)
    // AnalysisGrid.InvalidateArrange();   // 또는
    // AnalysisGrid.InvalidateVisual();
}
    private void Analysis_Click(object sender, RoutedEventArgs e)
    {
        MainPanel.IsVisible = false;
        MainList.IsVisible = false;
        AnalysisGrid.IsVisible = true;

        LoadAnalysisData();  // 매번 새로 불러오기
    }

    private void Admin_Click(object sender, RoutedEventArgs e)
    {
        MainPanel.IsVisible = false;
        AnalysisGrid.IsVisible = false;
        MainList.IsVisible = true;
    }
}