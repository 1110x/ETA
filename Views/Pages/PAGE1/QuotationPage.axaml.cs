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
    private bool _useContractDb = true;   // false = 측정인 DB, true = 계약 DB (기본)

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
        txbCompanyCount.Text = $"{_filteredCompanies.Count}개 업체";

        BuildCompanyGrid();
    }

    // ── 업체 그리드 빌드 (2컬럼) ────────────────────────────────────────────
    private void BuildCompanyGrid()
    {
        var spCompanies = this.FindControl<StackPanel>("spCompanies");
        if (spCompanies == null) return;

        spCompanies.Children.Clear();

        // 2컬럼 그리드
        var gridRow = new Grid
        {
            ColumnDefinitions = new Avalonia.Controls.ColumnDefinitions("*,*"),
            ColumnSpacing = 4,
            RowSpacing = 4,
        };

        int col = 0, row = 0;
        foreach (var c in _filteredCompanies)
        {
            var card = BuildCompanyCard(c);
            Avalonia.Controls.Grid.SetColumn(card, col);
            Avalonia.Controls.Grid.SetRow(card, row);
            gridRow.Children.Add(card);

            // 로우 정의 동적 추가
            if (col == 1)
            {
                row++;
                var rowDef = new Avalonia.Controls.RowDefinition(Avalonia.Controls.GridLength.Auto);
                gridRow.RowDefinitions.Add(rowDef);
            }

            col = (col + 1) % 2;
        }

        spCompanies.Children.Add(gridRow);
    }

    // ── 업체 카드 빌드 ──────────────────────────────────────────────────────
    private Avalonia.Controls.Border BuildCompanyCard(Contract c)
    {
        var cardContent = new Avalonia.Controls.StackPanel { Spacing = 4 };

        // 업체명 + 약칭 뱃지 (한 줄)
        var namePanel = new Avalonia.Controls.StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        // 약칭 뱃지
        if (!string.IsNullOrEmpty(c.C_Abbreviation))
        {
            try
            {
                var badgeColor = BadgeColorHelper.GetBadgeColor(c.C_Abbreviation);
                var badge = new Avalonia.Controls.Border
                {
                    Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse(badgeColor.Bg)),
                    CornerRadius = new Avalonia.CornerRadius(3),
                    Padding = new Avalonia.Thickness(4, 1),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    Child = new Avalonia.Controls.TextBlock
                    {
                        Text = c.C_Abbreviation,
                        FontSize = AppFonts.XS,
                        FontWeight = Avalonia.Media.FontWeight.Bold,
                        FontFamily = new Avalonia.Media.FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse(badgeColor.Fg)),
                    }
                };
                namePanel.Children.Add(badge);
            }
            catch { }
        }

        // 업체명
        namePanel.Children.Add(new Avalonia.Controls.TextBlock
        {
            Text = c.C_CompanyName,
            FontSize = AppFonts.MD,
            FontWeight = Avalonia.Media.FontWeight.SemiBold,
            FontFamily = new Avalonia.Media.FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = Avalonia.Media.Brush.Parse("#e0e0e0"),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        });

        cardContent.Children.Add(namePanel);

        // 계약 정보 (한 줄)
        var infoPanel = new Avalonia.Controls.StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 8,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        // 계약시작
        infoPanel.Children.Add(new Avalonia.Controls.TextBlock
        {
            Text = $"시작: {c.C_ContractStartStr}",
            FontSize = AppFonts.XS,
            Foreground = Avalonia.Media.Brush.Parse("#888888"),
            FontFamily = new Avalonia.Media.FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        });

        // 계약종료
        infoPanel.Children.Add(new Avalonia.Controls.TextBlock
        {
            Text = $"종료: {c.C_ContractEndStr}",
            FontSize = AppFonts.XS,
            Foreground = Avalonia.Media.Brush.Parse("#888888"),
            FontFamily = new Avalonia.Media.FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        });

        // 잔여일수
        infoPanel.Children.Add(new Avalonia.Controls.TextBlock
        {
            Text = c.DaysLeftText,
            FontSize = AppFonts.XS,
            Foreground = Avalonia.Media.Brush.Parse("#aaaaaa"),
            FontWeight = Avalonia.Media.FontWeight.Bold,
            FontFamily = new Avalonia.Media.FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        });

        cardContent.Children.Add(infoPanel);

        // 카드 Border
        var card = new Avalonia.Controls.Border
        {
            Background = Avalonia.Media.Brush.Parse("#1a1a28"),
            BorderBrush = Avalonia.Media.Brush.Parse("#333333"),
            BorderThickness = new Avalonia.Thickness(1),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(8, 6),
            Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child = cardContent,
        };

        // 클릭 이벤트
        var capturedContract = c;
        card.PointerPressed += (_, _) => CompanySelected?.Invoke(capturedContract);

        return card;
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

}
