using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Linq;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// DB Migration — Show4: 측정인_채취지점 트리뷰
/// 업체명(루트노드) → 채취지점명(자식노드) 구조
/// </summary>
public partial class DbMigrationPointPanel : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly Avalonia.Media.FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    /// <summary>채취지점명 노드 클릭 시 발생 — Show2로 전달</summary>
    public event Action<string>? SamplingPointSelected;

    private Contract? _selectedCompany;
    private bool      _showAll = false;

    public DbMigrationPointPanel()
    {
        InitializeComponent();
    }

    // ── 외부에서 호출: 선택된 ETA 업체 변경 ──────────────────────────────────
    public void SetSelectedCompany(Contract? company)
    {
        _selectedCompany = company;
        txbHint.IsVisible = company == null && !_showAll;
        BuildTree();
    }

    public void LoadAll()
    {
        _showAll = true;
        tglShowAll.IsChecked = true;
        txbHint.IsVisible = false;
        BuildTree();
    }

    // ── 트리 빌드 ─────────────────────────────────────────────────────────────
    private void BuildTree()
    {
        PointTreeView.Items.Clear();

        var allData = MeasurerService.GetAllData();  // (계약번호, 약칭, 계약기간, 업체명, 채취지점명)

        if (!_showAll && _selectedCompany != null)
        {
            // 선택된 ETA 업체명으로 퍼지 매칭
            var companyName = _selectedCompany.C_CompanyName;
            var abbr        = _selectedCompany.C_Abbreviation;

            var matched = allData.Where(d =>
                d.업체명.Contains(companyName, StringComparison.OrdinalIgnoreCase) ||
                companyName.Contains(d.업체명,  StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(abbr) &&
                 (d.약칭.Contains(abbr, StringComparison.OrdinalIgnoreCase) ||
                  abbr.Contains(d.약칭, StringComparison.OrdinalIgnoreCase) ||
                  d.업체명.Contains(abbr, StringComparison.OrdinalIgnoreCase)))
            ).ToList();

            if (matched.Count == 0)
            {
                PointTreeView.Items.Add(new TreeViewItem {
                    Header = MakeTextBlock("⚠️ 매칭되는 채취지점 없음", "#cc8800")
                });
                txbPointCount.Text = "매칭 없음";
                txbHint.IsVisible = false;
                return;
            }

            // 그룹 by 업체명
            foreach (var grp in matched.GroupBy(d => d.업체명).OrderBy(g => g.Key))
                PointTreeView.Items.Add(MakeCompanyNode(grp.Key, grp.First().약칭, grp.First().계약기간, grp.Select(d => d.채취지점명).ToList(), expand: true));

            int total = matched.Count;
            txbPointCount.Text = $"{total}개 채취지점";
        }
        else if (_showAll)
        {
            foreach (var grp in allData.GroupBy(d => d.업체명).OrderBy(g => g.Key))
                PointTreeView.Items.Add(MakeCompanyNode(grp.Key, grp.First().약칭, grp.First().계약기간, grp.Select(d => d.채취지점명).ToList(), expand: false));

            txbPointCount.Text = $"{allData.Count}개 채취지점 / {allData.Select(d => d.업체명).Distinct().Count()}개 업체";
        }

        txbHint.IsVisible = false;
    }

    private TreeViewItem MakeCompanyNode(string company, string abbr, string period, System.Collections.Generic.List<string> points, bool expand)
    {
        // 헤더: 업체명 + 약칭 배지 + 건수
        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        header.Children.Add(new TextBlock { Text = "🏭", FontSize = AppTheme.FontXL, VerticalAlignment = VerticalAlignment.Center });
        header.Children.Add(new TextBlock {
            Text = company, FontSize = AppTheme.FontMD, FontFamily = Font,
            Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center
        });

        if (!string.IsNullOrWhiteSpace(abbr))
        {
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);
            header.Children.Add(new Border {
                Background   = new SolidColorBrush(Avalonia.Media.Color.Parse(bg)),
                CornerRadius = new Avalonia.CornerRadius(3),
                Padding      = new Avalonia.Thickness(4, 1),
                Child        = new TextBlock { Text = abbr, FontSize = AppTheme.FontSM,
                                               Foreground = new SolidColorBrush(Avalonia.Media.Color.Parse(fg)) }
            });
        }

        header.Children.Add(new TextBlock {
            Text = $"({points.Count})", FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = new SolidColorBrush(Avalonia.Media.Color.Parse("#888888")),
            VerticalAlignment = VerticalAlignment.Center
        });

        var node = new TreeViewItem { Header = header, IsExpanded = expand };

        foreach (var pt in points.OrderBy(p => p))
        {
            var capturedPt = pt;
            var leafNode = new TreeViewItem {
                Header = new StackPanel {
                    Orientation = Orientation.Horizontal, Spacing = 6,
                    Children = {
                        new TextBlock { Text = "📍", FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = pt, FontSize = AppTheme.FontBase, FontFamily = Font,
                                        Foreground = new SolidColorBrush(Avalonia.Media.Color.Parse("#bbddff")),
                                        VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            leafNode.Tapped += (_, _) => SamplingPointSelected?.Invoke(capturedPt);
            node.Items.Add(leafNode);
        }

        return node;
    }

    private static TextBlock MakeTextBlock(string text, string color) =>
        new() { Text = text, FontSize = AppTheme.FontMD, FontFamily = Font,
                Foreground = new SolidColorBrush(Avalonia.Media.Color.Parse(color)) };

    private void TglShowAll_Changed(object? sender, RoutedEventArgs e)
    {
        _showAll = tglShowAll.IsChecked == true;
        txbHint.IsVisible = !_showAll && _selectedCompany == null;
        BuildTree();
    }
}
