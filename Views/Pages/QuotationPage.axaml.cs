using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using System.Collections.Generic;
using System.Linq;
using System;

namespace ETA.Views.Pages;

/// <summary>
/// Content1 에 표시되는 계약업체 목록 패널.
/// 업체를 선택하면 MainPage 가 Content4(분석항목 트리) 와
/// Content2(신규/발행내역) 를 연동해 준다.
/// </summary>
public partial class QuotationPage : UserControl
{
    // ── 이벤트: 업체 선택 시 MainPage 에 알림 ─────────────────────────────
    public event Action<Contract>? CompanySelected;

    // ── 내부 데이터 ───────────────────────────────────────────────────────
    private List<Contract> _allCompanies = [];
    private bool _useContractDb = false;   // false = 측정인 DB (기본), true = 계약 DB

    public QuotationPage()
    {
        InitializeComponent();
    }

    // ── 외부에서 호출: 데이터 로드 ───────────────────────────────────────
    public void LoadData()
    {
        if (_useContractDb)
        {
            _allCompanies = QuotationService.GetContractCompanies(activeOnly: true);
            txbHeader.Text = "계약업체";
        }
        else
        {
            MeasurerService.SyncAbbrFromContractDb();
            _allCompanies = MeasurerService.GetCompaniesAsContracts();
            txbHeader.Text = "측정인 업체";
        }
        ApplyFilter(txbSearch.Text ?? "");
    }

    // ── 필터 적용 ─────────────────────────────────────────────────────────
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
        txbCompanyCount.Text = $"{filtered.Count}개 업체";
    }

    // ── 이벤트 핸들러 ────────────────────────────────────────────────────

    private void TglActiveOnly_Changed(object? sender, RoutedEventArgs e)
    {
        _useContractDb = tglActiveOnly.IsChecked == true;
        LoadData();
    }

    private void TxbSearch_TextChanged(object? sender, TextChangedEventArgs e)
    {
        ApplyFilter(txbSearch.Text ?? "");
    }

    private void LbxCompanies_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (lbxCompanies.SelectedItem is Contract c)
            CompanySelected?.Invoke(c);
    }
}
