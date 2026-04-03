using Avalonia.Controls;
using ETA.Views;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// DB Migration — Show2: 시료명 매핑 테이블
/// 기존 시료명 | 새 시료명(채취지점) 2열 구조
/// Show4에서 채취지점을 클릭하면 선택된 행에 채취지점명이 채워짐
/// 저장 시 분석의뢰및결과.시료명 을 일괄 UPDATE
/// </summary>
public partial class DbMigrationMappingPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly FontFamily FontM =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 상태 ─────────────────────────────────────────────────────────────────
    private Contract?  _company;
    private int        _selectedIndex = -1;

    // 각 행: (원본시료명, 새시료명, 행Border)
    private readonly List<(string OldName, string NewName, Border RowBorder, TextBlock NewTxb)> _rows = new();

    public DbMigrationMappingPanel()
    {
        InitializeComponent();
    }

    // ── 외부에서 호출: 업체 변경 시 테이블 로드 ─────────────────────────────
    public void LoadCompany(Contract company)
    {
        _company       = company;
        _selectedIndex = -1;
        _rows.Clear();
        spRows.Children.Clear();
        txbSaveStatus.Text = "";

        // 배지 표시
        txbCompanyBadge.Text   = $"{company.C_CompanyName}  ({company.C_Abbreviation})";
        bdgCompany.IsVisible   = true;
        btnSave.IsEnabled      = false;

        var abbr = company.C_Abbreviation;
        if (string.IsNullOrWhiteSpace(abbr))
        {
            txbRowCount.Text = "약칭 없음 — 검색 불가";
            return;
        }

        var names = AnalysisRequestService.GetDistinctSampleNames(abbr);
        txbRowCount.Text = $"{names.Count}개 시료명";

        if (names.Count == 0)
        {
            spRows.Children.Add(new TextBlock
            {
                Text       = "해당 약칭의 시료명 없음",
                FontSize   = AppTheme.FontBase, FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Avalonia.Thickness(8, 6),
            });
            return;
        }

        for (int i = 0; i < names.Count; i++)
        {
            var idx     = i;
            var oldName = names[i];
            var newTxb  = new TextBlock { Classes = { "col-new" }, Margin = new Avalonia.Thickness(4, 0), Text = "" };
            var oldTxb  = new TextBlock { Classes = { "col-old" }, Margin = new Avalonia.Thickness(4, 0), Text = oldName };

            var rowGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*") };
            rowGrid.Children.Add(oldTxb);
            Grid.SetColumn(newTxb, 1);
            rowGrid.Children.Add(newTxb);

            var border = new Border
            {
                Classes = { "row-item" },
                Padding = new Avalonia.Thickness(6, 5),
                Child   = rowGrid,
            };

            var capturedIdx = idx;
            border.PointerPressed += (_, _) => SelectRow(capturedIdx);

            _rows.Add((oldName, "", border, newTxb));
            spRows.Children.Add(border);
        }
    }

    // ── Show4에서 채취지점 클릭 시 → 선택된 행에 새 이름 채움 ───────────────
    public void SetMappingPoint(string pointName)
    {
        if (_selectedIndex < 0 || _selectedIndex >= _rows.Count) return;

        var (old, _, border, txb) = _rows[_selectedIndex];
        txb.Text = pointName;
        _rows[_selectedIndex] = (old, pointName, border, txb);

        // 저장 버튼 활성화 여부 갱신
        btnSave.IsEnabled = _rows.Any(r => !string.IsNullOrWhiteSpace(r.NewName));

        // 자동으로 다음 행 선택
        int next = _selectedIndex + 1;
        if (next < _rows.Count) SelectRow(next);
    }

    // ── 행 선택 ──────────────────────────────────────────────────────────────
    private void SelectRow(int idx)
    {
        // 이전 선택 해제
        if (_selectedIndex >= 0 && _selectedIndex < _rows.Count)
        {
            _rows[_selectedIndex].RowBorder.Classes.Remove("row-selected");
            _rows[_selectedIndex].RowBorder.Classes.Add("row-item");
        }

        _selectedIndex = idx;

        if (idx >= 0 && idx < _rows.Count)
        {
            _rows[idx].RowBorder.Classes.Remove("row-item");
            _rows[idx].RowBorder.Classes.Add("row-selected");
        }
    }

    // ── 저장 ─────────────────────────────────────────────────────────────────
    private void BtnSave_Click(object? sender, RoutedEventArgs e)
    {
        if (_company == null) return;

        var abbr    = _company.C_Abbreviation;
        int total   = 0;
        var errors  = new List<string>();

        foreach (var (oldName, newName, _, _) in _rows)
        {
            if (string.IsNullOrWhiteSpace(newName)) continue;
            if (string.Equals(oldName, newName, StringComparison.Ordinal)) continue;

            int updated = AnalysisRequestService.RenameSampleName(abbr, oldName, newName);
            if (updated > 0) total += updated;
            else             errors.Add(oldName);
        }

        if (errors.Count == 0)
            txbSaveStatus.Text = $"✅ {total}건 저장 완료";
        else
            txbSaveStatus.Text = $"⚠️ {total}건 저장, 실패 {errors.Count}건";

        // 저장 후 목록 갱신
        if (_company != null) LoadCompany(_company);
    }
}
