using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages;

/// <summary>
/// DB Migration — Show1: ETA 계약 DB 업체 목록 (약칭 배지 포함)
/// 업체 선택 시 Show4 패널(DbMigrationPointPanel)에 해당 업체의 측정인 채취지점을 표시
/// </summary>
public partial class DbMigrationPage : UserControl
{
    public event Action<Contract?>? CompanySelected;

    private List<Contract> _allCompanies = [];
    private bool _activeOnly = true;

    public DbMigrationPage()
    {
        InitializeComponent();
    }

    public void LoadData()
    {
        _allCompanies = QuotationService.GetContractCompanies(activeOnly: _activeOnly);
        ApplyFilter(txbSearch.Text ?? "");
    }

    private void ApplyFilter(string keyword)
    {
        IEnumerable<Contract> source = _allCompanies;
        if (!string.IsNullOrWhiteSpace(keyword))
        {
            var kw = keyword.Trim();
            source = source.Where(c =>
                (c.C_CompanyName  ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                (c.C_Abbreviation ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));
        }
        var filtered = source.ToList();
        lbxCompanies.ItemsSource = filtered;
        txbCount.Text = $"{filtered.Count}개 업체";
    }

    private void TglActiveOnly_Changed(object? sender, RoutedEventArgs e)
    {
        _activeOnly = tglActiveOnly.IsChecked == true;
        LoadData();
    }

    private void TxbSearch_TextChanged(object? sender, TextChangedEventArgs e)
    {
        ApplyFilter(txbSearch.Text ?? "");
    }

    private void LbxCompanies_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (lbxCompanies.SelectedItem is Contract c)
        {
            txbSelected.Text      = $"▶ {c.C_CompanyName}";
            txbSelected.IsVisible = true;
            CompanySelected?.Invoke(c);
        }
        else
        {
            txbSelected.IsVisible = false;
            CompanySelected?.Invoke(null);
        }
    }
}
