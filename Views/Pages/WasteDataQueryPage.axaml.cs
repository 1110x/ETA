using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Models;
using ETA.Services;

namespace ETA.Views.Pages;

public partial class WasteDataQueryPage : UserControl
{
    // ── 외부 이벤트 ─────────────────────────────────────────────────────────
    public event Action<WasteCompany>? CompanySelected;

    // ── 상태 ────────────────────────────────────────────────────────────────
    private string _activeGroup = "여수";

    public WasteDataQueryPage()
    {
        InitializeComponent();
    }

    // ── 데이터 로드 ─────────────────────────────────────────────────────────
    public void LoadData()
    {
        WasteCompanyTreeView.Items.Clear();
        try
        {
            var items = WasteCompanyService.GetAllItems()
                            .Where(c => GetGroupLabel(c) == _activeGroup)
                            .OrderBy(c => c.관리번호)
                            .ToList();

            foreach (var item in items)
                WasteCompanyTreeView.Items.Add(CreateCompanyNode(item));
        }
        catch { }
    }

    // ── 그룹 탭 토글 ────────────────────────────────────────────────────────
    public void GroupTab_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        _activeGroup = btn.Tag as string ?? "여수";
        UpdateTabStyle();
        LoadData();
    }

    // ── 업체 선택 → 이벤트 발생 ─────────────────────────────────────────────
    private void WasteCompanyTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        WasteCompany? company = null;
        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is WasteCompany c1) company = c1;
            else if (e.AddedItems[0] is WasteCompany c2) company = c2;
        }
        if (company != null) CompanySelected?.Invoke(company);
    }

    // ── 노드 생성 ───────────────────────────────────────────────────────────
    private static TreeViewItem CreateCompanyNode(WasteCompany company)
    {
        var sp = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 4,
            VerticalAlignment = VerticalAlignment.Center,
        };
        sp.Children.Add(new TextBlock
        {
            Text      = company.업체명,
            FontSize  = 12,
            FontFamily = Font,
            Foreground = Brushes.WhiteSmoke,
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text      = company.관리번호,
            FontSize  = 9,
            FontFamily = Font,
            Foreground = new SolidColorBrush(Color.Parse("#666666")),
            VerticalAlignment = VerticalAlignment.Center,
        });
        return new TreeViewItem { Tag = company, Header = sp };
    }

    // ── 그룹 라벨 ───────────────────────────────────────────────────────────
    private static string GetGroupLabel(WasteCompany c)
    {
        var p = c.프로젝트명 ?? "";
        if (p.Contains("여수")) return "여수";
        if (p.Contains("세풍")) return "세풍";
        if (p.Contains("율촌")) return "율촌";
        return "기타";
    }

    // ── 탭 스타일 ───────────────────────────────────────────────────────────
    private static readonly Dictionary<string, (string bg, string fg, string border)> TabColors = new()
    {
        ["여수"] = ("#1e3a5a", "#88aacc", "#336699"),
        ["율촌"] = ("#1a3a1a", "#aaccaa", "#336633"),
        ["세풍"] = ("#3a2a1a", "#ccaa88", "#996633"),
    };

    private void UpdateTabStyle()
    {
        foreach (var btn in new[] { BtnYeosoo, BtnYulchon, BtnSepung })
        {
            var group  = btn.Tag as string ?? "";
            bool active = group == _activeGroup;
            if (active && TabColors.TryGetValue(group, out var c))
            {
                btn.Background  = new SolidColorBrush(Color.Parse(c.bg));
                btn.Foreground  = new SolidColorBrush(Color.Parse(c.fg));
                btn.BorderBrush = new SolidColorBrush(Color.Parse(c.border));
            }
            else
            {
                btn.Background  = new SolidColorBrush(Color.Parse("#222222"));
                btn.Foreground  = new SolidColorBrush(Color.Parse("#666666"));
                btn.BorderBrush = new SolidColorBrush(Color.Parse("#444444"));
            }
        }
    }

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
}
