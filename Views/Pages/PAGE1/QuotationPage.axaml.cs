using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System.Collections.Generic;
using System.Linq;
using System;

namespace ETA.Views.Pages.PAGE1;

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
    private List<Contract> _allCompanies = new();
    public List<Contract> Companies => _filteredCompanies;
    private List<Contract> _filteredCompanies = new();
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

        _filteredCompanies = source.ToList();
        // 기존 선택 항목 기억
        var prevSelected = lbCompanies.SelectedItem as Contract;
        lbCompanies.ItemsSource = _filteredCompanies;
        txbCompanyCount.Text = $"{_filteredCompanies.Count}개 업체";
        // 기존 선택 복원
        if (prevSelected != null)
        {
            var match = _filteredCompanies.FirstOrDefault(c => c.C_CompanyName == prevSelected.C_CompanyName);
            if (match != null)
                lbCompanies.SelectedItem = match;
        }
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

    // ListBox SelectionChanged 핸들러
    private void LbCompanies_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (lbCompanies.SelectedItem is Contract c)
            CompanySelected?.Invoke(c);
    }
}
