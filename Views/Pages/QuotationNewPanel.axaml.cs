using Avalonia.Controls;
using Avalonia.Interactivity;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ETA.Views.Pages;

/// <summary>Content2 — 신규 견적 작성 폼</summary>
public partial class QuotationNewPanel : UserControl
{
    private Contract?          _selectedCompany;
    private List<AnalysisItem> _selectedAnalytes = [];

    public QuotationNewPanel()
    {
        InitializeComponent();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbQuotationNo.Text = GenerateNo();
    }

    // ── 외부 연동 ─────────────────────────────────────────────────────────
    public void SetCompany(Contract company)
    {
        _selectedCompany = company;
        txbCompany.Text  = $"{company.C_CompanyName}  [{company.C_Abbreviation}]";
    }

    public void SetSelectedAnalytes(List<AnalysisItem> items)
    {
        _selectedAnalytes = items;
        if (items.Count == 0)
        {
            txbAnalytesSummary.Text = "아래 패널에서 분석 항목을 선택하세요.";
            txbTotalAmount.Text     = "선택 0개 항목";
            return;
        }
        var names = items.Select(a => a.Analyte).ToList();
        txbAnalytesSummary.Text =
            string.Join(", ", names.Take(8))
            + (names.Count > 8 ? $"  외 {names.Count - 8}개" : "");
        txbTotalAmount.Text = $"선택 {items.Count}개 항목";
    }

    // ── 버튼 ─────────────────────────────────────────────────────────────
    private void BtnSave_Click(object? sender, RoutedEventArgs e)
    {
        if (_selectedCompany == null)
        { Debug.WriteLine("[NewPanel] 업체 미선택"); return; }

        var issue = new QuotationIssue
        {
            발행일   = txbIssueDate.Text   ?? DateTime.Today.ToString("yyyy-MM-dd"),
            업체명   = _selectedCompany.C_CompanyName  ?? "",
            약칭     = _selectedCompany.C_Abbreviation ?? "",
            시료명   = txbSampleName.Text  ?? "",
            견적번호 = txbQuotationNo.Text ?? GenerateNo(),
            견적구분 = (cmbQuotationType.SelectedItem as ComboBoxItem)
                       ?.Content?.ToString() ?? "",
            총금액   = 0,
        };

        bool ok = QuotationService.Insert(issue);
        Debug.WriteLine(ok
            ? $"[NewPanel] 저장 완료 → {issue.견적번호}"
            : "[NewPanel] 저장 실패");

        if (ok) txbQuotationNo.Text = GenerateNo();
    }

    private void BtnPreview_Click(object? sender, RoutedEventArgs e)
        => Debug.WriteLine("[NewPanel] 미리보기 — 추후 연동");

    private static string GenerateNo()
        => $"Q-{DateTime.Today:yyyyMM}-{DateTime.Now:HHmm}";
}
