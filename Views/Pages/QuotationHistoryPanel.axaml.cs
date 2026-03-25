using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

/// <summary>Content1 — 견적 발행 내역 트리뷰</summary>
public partial class QuotationHistoryPanel : UserControl
{
    public event Action<QuotationIssue>? IssueSelected;

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    private int           _selectedRowid = -1;
    private TreeViewItem? _selectedNode  = null;

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
            _selectedNode  = null;
            txbInfo.Text   = "";

            var all = QuotationService.GetAllIssues();
            if (all.Count == 0)
            {
                tvHistory.Items.Add(new TreeViewItem
                {
                    Header = new TextBlock
                    {
                        Text = "발행 내역이 없습니다.", FontSize = 11,
                        FontFamily = Font, Foreground = Brush.Parse("#555"),
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
                    Header     = MakeHeader($"📁  {yg.Key}년", 12, "#a0d4a0", bold: true),
                    Tag        = yg.Key,
                    IsExpanded = true,
                };

                foreach (var mg in yg.GroupBy(i => ParseMonth(i.발행일))
                                      .OrderByDescending(g => g.Key))
                {
                    bool isCurrent = yg.Key == thisYear && mg.Key == thisMonth;
                    var monthNode = new TreeViewItem
                    {
                        Header = MakeHeader(
                            $"{(isCurrent ? "📂" : "📁")}  {yg.Key}-{mg.Key:D2}  ({mg.Count()}건)",
                            11, isCurrent ? "#aaaacc" : "#666688"),
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
                    Text = $"로드 오류: {ex.Message}", FontSize = 11,
                    FontFamily = Font, Foreground = Brush.Parse("#f88"),
                },
            });
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  노드 빌더
    // ══════════════════════════════════════════════════════════════════════

    private static TextBlock MakeHeader(string text, double size, string color,
                                        bool bold = false) => new()
    {
        Text      = text,
        FontSize  = size,
        FontFamily = Font,
        FontWeight = bold ? FontWeight.SemiBold : FontWeight.Normal,
        Foreground = Brush.Parse(color),
        VerticalAlignment = VerticalAlignment.Center,
    };

    private TreeViewItem MakeIssueNode(QuotationIssue issue)
    {
        var sub = string.Join("  ·  ", new[]
        {
            issue.업체명,
            string.IsNullOrEmpty(issue.견적번호) ? null : issue.견적번호,
            issue.총금액 > 0 ? $"{issue.총금액:#,0}원" : null,
        }.Where(s => !string.IsNullOrEmpty(s)));

        var header = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing     = 0,
            Children    =
            {
                new TextBlock
                {
                    Text       = $"🧪  {(string.IsNullOrEmpty(issue.시료명) ? "(시료명 없음)" : issue.시료명)}",
                    FontSize   = 11,
                    FontFamily = Font,
                    Foreground = Brushes.WhiteSmoke,
                },
                new TextBlock
                {
                    Text       = "    " + sub,
                    FontSize   = 9,
                    FontFamily = Font,
                    Foreground = Brush.Parse("#777799"),
                },
            },
        };

        var node = new TreeViewItem { Header = header, Tag = issue };

        // handledEventsToo=true: Avalonia가 내부적으로 처리한 이벤트도 받음
        node.AddHandler(
            PointerReleasedEvent,
            (_, _) => OnIssueNodeTapped(node, issue),
            handledEventsToo: true);

        return node;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  선택 처리
    // ══════════════════════════════════════════════════════════════════════

    // 발행건 탭 → 이전 노드 배경 해제 + 현재 노드 하이라이트
    private void OnIssueNodeTapped(TreeViewItem node, QuotationIssue issue)
    {
        if (_selectedNode != null && _selectedNode != node)
            _selectedNode.Background = Brushes.Transparent;

        node.Background = Brush.Parse("#1a3a1a");
        _selectedNode   = node;
        _selectedRowid  = issue.Id;
        txbInfo.Text    = $"🧪 {issue.시료명}  ·  {issue.업체명}  ·  {issue.발행일}";

        IssueSelected?.Invoke(issue);
    }

    // TreeView SelectionChanged: 년/월 클릭 시 이전 발행건 선택 유지
    private void TvHistory_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (tvHistory.SelectedItem is TreeViewItem item &&
            item.Tag is QuotationIssue issue)
        {
            OnIssueNodeTapped(item, issue);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  버튼
    // ══════════════════════════════════════════════════════════════════════

    private void BtnRefresh_Click(object? sender, RoutedEventArgs e) => LoadData();

    private async void BtnDelete_Click(object? sender, RoutedEventArgs e)
    {
        if (_selectedNode == null || _selectedNode.Tag is not QuotationIssue issue)
        {
            await ShowMsg("발행건(🧪 시료명)을 선택한 후 삭제해주세요.\n년/월 노드는 삭제할 수 없습니다.");
            return;
        }

        bool confirmed = await ShowConfirm(
            $"아래 발행건을 삭제하시겠습니까?\n\n" +
            $"  시료명 : {issue.시료명}\n" +
            $"  업체명 : {issue.업체명}\n" +
            $"  발행일 : {issue.발행일}\n\n" +
            "삭제 후 복구할 수 없습니다.");

        if (!confirmed) return;

        if (QuotationService.Delete(issue.Id))
            LoadData();
    }

    // ── 확인 다이얼로그 ───────────────────────────────────────────────────
    private async Task<bool> ShowConfirm(string message)
    {
        bool result = false;
        var owner   = TopLevel.GetTopLevel(this) as Window;

        var dlg = new Window
        {
            Title = "삭제 확인", Width = 340, Height = 190,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Brush.Parse("#1e1e2a"),
            SystemDecorations = SystemDecorations.BorderOnly,
        };

        var yesBtn = new Button
        {
            Content = "삭제", Width = 80, Height = 28,
            Background = Brush.Parse("#6a2a2a"), Foreground = Brush.Parse("#f0aeae"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11,
        };
        var noBtn = new Button
        {
            Content = "취소", Width = 80, Height = 28,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11,
        };

        yesBtn.Click += (_, _) => { result = true;  dlg.Close(); };
        noBtn.Click  += (_, _) => { result = false; dlg.Close(); };

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text = message, FontFamily = Font, FontSize = 11,
                    Foreground = Brush.Parse("#dddddd"),
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    LineHeight = 18,
                },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 8,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children = { yesBtn, noBtn },
                }
            }
        };

        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();

        return result;
    }

    private async Task ShowMsg(string message)
    {
        var owner = TopLevel.GetTopLevel(this) as Window;
        var dlg   = new Window
        {
            Title = "안내", Width = 300, Height = 140,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Brush.Parse("#1e1e2a"),
            SystemDecorations = SystemDecorations.BorderOnly,
        };
        var okBtn = new Button
        {
            Content = "확인", Width = 70, Height = 26,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = 11,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        okBtn.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock
                {
                    Text = message, FontFamily = Font, FontSize = 11,
                    Foreground = Brush.Parse("#dddddd"),
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                okBtn,
            }
        };
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static int ParseYear(string d)
        => DateTime.TryParse(d, out var dt) ? dt.Year
         : d.Length >= 4 && int.TryParse(d[..4], out int y) ? y : 0;

    private static int ParseMonth(string d)
        => DateTime.TryParse(d, out var dt) ? dt.Month
         : d.Length >= 7 && int.TryParse(d.Substring(5, 2), out int m) ? m : 0;
}
