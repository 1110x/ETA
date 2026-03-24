using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ETA.Views.Pages;

/// <summary>Content1 — 견적 발행 내역 트리뷰</summary>
public partial class QuotationHistoryPanel : UserControl
{
    // 발행건 선택 → Content2 세부내역 표시용
    public event Action<QuotationIssue>? IssueSelected;

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    private int _selectedRowid = -1;

    public QuotationHistoryPanel()
    {
        InitializeComponent();
    }

    // ── 외부 호출 ─────────────────────────────────────────────────────────
    public void LoadData()
    {
        try
        {
            tvHistory.Items.Clear();
            _selectedRowid = -1;
            txbInfo.Text   = "";

            var all = QuotationService.GetAllIssues();
            if (all.Count == 0)
            {
                tvHistory.Items.Add(new TreeViewItem
                {
                    Header     = new TextBlock
                    {
                        Text       = "발행 내역이 없습니다.",
                        FontSize   = 12,
                        FontFamily = Font,
                        Foreground = Brush.Parse("#555"),
                    },
                });
                return;
            }

            int thisYear  = DateTime.Today.Year;
            int thisMonth = DateTime.Today.Month;

            foreach (var yg in all.GroupBy(i => ParseYear(i.발행일))
                                   .OrderByDescending(g => g.Key))
            {
                var yearNode = new TreeViewItem
                {
                    Header     = MakeYearHeader(yg.Key),
                    Tag        = yg.Key,
                    IsExpanded = true,
                };

                foreach (var mg in yg.GroupBy(i => ParseMonth(i.발행일))
                                      .OrderByDescending(g => g.Key))
                {
                    bool isCurrent = yg.Key == thisYear && mg.Key == thisMonth;

                    var monthNode = new TreeViewItem
                    {
                        Header     = MakeMonthHeader(yg.Key, mg.Key,
                                         mg.Count(), isCurrent),
                        Tag        = (yg.Key, mg.Key),
                        IsExpanded = isCurrent,
                    };

                    foreach (var issue in mg.OrderByDescending(i => i.발행일))
                        monthNode.Items.Add(MakeIssueNode(issue));

                    yearNode.Items.Add(monthNode);
                }

                tvHistory.Items.Add(yearNode);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[HistoryPanel] 오류: {ex.Message}");
            tvHistory.Items.Clear();
            tvHistory.Items.Add(new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text       = $"로드 오류: {ex.Message}",
                    FontSize   = 12,
                    FontFamily = Font,
                    Foreground = Brush.Parse("#f88"),
                },
            });
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  노드 빌더
    // ══════════════════════════════════════════════════════════════════════

    // ── 년 노드 ───────────────────────────────────────────────────────────
    private static TreeViewItem MakeYearHeader(int year)
    {
        var sp = Row(6);
        sp.Children.Add(Txt("📁", 14));
        sp.Children.Add(new TextBlock
        {
            Text       = $"{year}년",
            FontSize   = 13,
            FontFamily = Font,
            FontWeight = FontWeight.Bold,
            Foreground = Brush.Parse("#a0d4a0"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        return new TreeViewItem { Header = sp, Tag = year, IsExpanded = true };
    }

    // ── 월 노드 ───────────────────────────────────────────────────────────
    private static TreeViewItem MakeMonthHeader(int year, int month,
                                                int count, bool isCurrent)
    {
        var sp = Row(6);
        sp.Children.Add(Txt(isCurrent ? "📂" : "📁", 13));
        sp.Children.Add(new TextBlock
        {
            Text       = $"  📅 {year}-{month:D2}  ({count}건)",
            FontSize   = 12,
            FontFamily = Font,
            Foreground = Brush.Parse(isCurrent ? "#aaaacc" : "#777799"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        return new TreeViewItem
        {
            Header     = sp,
            Tag        = (year, month),
            IsExpanded = isCurrent,
        };
    }

    // ── 발행건(시료명) 노드 ───────────────────────────────────────────────
    private static TreeViewItem MakeIssueNode(QuotationIssue issue)
    {
        // 서브 라인: 업체명 · 견적번호 · 합계금액
        var sub = string.Join("  ·  ", new[]
        {
            issue.업체명,
            string.IsNullOrEmpty(issue.견적번호) ? null : issue.견적번호,
            issue.총금액 > 0 ? $"{issue.총금액:#,0}원" : null,
        }.Where(s => !string.IsNullOrEmpty(s)));

        var inner = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing     = 1,
            Children =
            {
                new TextBlock
                {
                    Text       = string.IsNullOrEmpty(issue.시료명) ? "(시료명 없음)" : issue.시료명,
                    FontSize   = 12,
                    FontFamily = Font,
                    Foreground = Brushes.WhiteSmoke,
                },
                new TextBlock
                {
                    Text       = sub,
                    FontSize   = 10,
                    FontFamily = Font,
                    Foreground = Brush.Parse("#777799"),
                },
            },
        };

        var row = Row(4);
        row.Margin = new Thickness(4, 0, 0, 0);
        row.Children.Add(Txt("🧪", 13));
        row.Children.Add(inner);

        var node = new TreeViewItem { Header = row, Tag = issue };
        return node;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  이벤트
    // ══════════════════════════════════════════════════════════════════════

    private void TvHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (tvHistory.SelectedItem is TreeViewItem item &&
            item.Tag is QuotationIssue issue)
        {
            _selectedRowid = issue.Id;
            txbInfo.Text   = $"{issue.업체명}  {issue.시료명}  {issue.발행일}";
            IssueSelected?.Invoke(issue);   // ← Content2 세부내역 표시 트리거
        }
        else
        {
            _selectedRowid = -1;
            txbInfo.Text   = "";
        }
    }

    private void BtnRefresh_Click(object? sender, RoutedEventArgs e) => LoadData();

    private void BtnDelete_Click(object? sender, RoutedEventArgs e)
    {
        if (_selectedRowid < 0) return;
        if (QuotationService.Delete(_selectedRowid))
            LoadData();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════

    private static StackPanel Row(double spacing) => new()
    {
        Orientation       = Orientation.Horizontal,
        Spacing           = spacing,
        VerticalAlignment = VerticalAlignment.Center,
    };

    private static TextBlock Txt(string text, double size) => new()
    {
        Text              = text,
        FontSize          = size,
        VerticalAlignment = VerticalAlignment.Center,
    };

    private static int ParseYear(string d)
        => DateTime.TryParse(d, out var dt) ? dt.Year
         : d.Length >= 4 && int.TryParse(d[..4], out int y) ? y : 0;

    private static int ParseMonth(string d)
        => DateTime.TryParse(d, out var dt) ? dt.Month
         : d.Length >= 7 && int.TryParse(d.Substring(5, 2), out int m) ? m : 0;
}
