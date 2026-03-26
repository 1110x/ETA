using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

/// <summary>
/// Content4 (분석의뢰 탭): 선택된 분석의뢰를 트리뷰로 관리
/// - 노드 추가: IsExpanded=false
/// - 우클릭 → 컨텍스트 메뉴 → 삭제 확인
/// - 📌 TODO 전송:
///     부모노드 1개 → 중간 Task 1개  "(03/25) 보임 폐수조#Apond-폐수"
///     자식노드 N개 → 최종 Task N개  "(03/31) BOD 김지은"
/// - 🖨 의뢰서 출력
/// </summary>
public class AnalysisRequestListPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    private static readonly IBrush BrushChildBg  = Brush.Parse("#1e1e2e");
    private static readonly IBrush BrushChildFg  = Brush.Parse("#aaaacc");
    private static readonly IBrush BrushAbbrBg   = Brush.Parse("#3a1a1a");
    private static readonly IBrush BrushAbbrFg   = Brush.Parse("#cc8888");
    private static readonly IBrush BrushAccNo    = Brush.Parse("#445566");
    private static readonly IBrush BrushParentFg = Brush.Parse("#88ee88");

    private readonly TreeView _tree = new();
    private readonly HashSet<string> _addedKeys = new();

    // DB row 캐시: "접수번호:Id" → row (TODO 전송/출력 시 재사용)
    private readonly Dictionary<string, Dictionary<string, string>> _rowCache = new();

    // 상태 표시
    private TextBlock? _txbStatus;

    public AnalysisRequestListPanel()
    {
        Content = BuildUI();
    }

    private Control BuildUI()
    {
        var headerGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto,Auto"),
            Margin = new Avalonia.Thickness(0, 0, 0, 4),
        };
        headerGrid.Children.Add(new TextBlock
        {
            Text = "📋  분석의뢰 선택 목록",
            FontSize = 12, FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = Brush.Parse("#e0e0e0"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        // 📌 TODO 전송
        var btnTodo = new Button
        {
            Content = "📌 TODO 전송",
            Height = 24, FontSize = 10, FontFamily = Font,
            Background = Brush.Parse("#2a2a1a"),
            Foreground = Brush.Parse("#e0c060"),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(10, 0),
            Margin = new Avalonia.Thickness(4, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        };
        btnTodo.Click += BtnTodo_Click;
        headerGrid.Children.Add(btnTodo);

        // 🖨 의뢰서 출력
        var btnPrint = new Button
        {
            Content = "🖨 의뢰서 출력",
            Height = 24, FontSize = 10, FontFamily = Font,
            Background = Brush.Parse("#1a2a3a"),
            Foreground = Brush.Parse("#88ccff"),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(10, 0),
            Margin = new Avalonia.Thickness(4, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 2,
        };
        btnPrint.Click += BtnPrint_Click;
        headerGrid.Children.Add(btnPrint);

        // 전체 삭제
        var btnClear = new Button
        {
            Content = "전체 삭제",
            Height = 24, FontSize = 10, FontFamily = Font,
            Background = Brush.Parse("#3a1a1a"),
            Foreground = Brush.Parse("#f0aeae"),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(8, 0),
            Margin = new Avalonia.Thickness(4, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 3,
        };
        btnClear.Click += (_, _) =>
        {
            _tree.Items.Clear();
            _addedKeys.Clear();
            _rowCache.Clear();
            SetStatus("");
        };
        headerGrid.Children.Add(btnClear);

        _txbStatus = new TextBlock
        {
            Text = "",
            FontSize = 9, FontFamily = Font,
            Foreground = Brush.Parse("#88aa88"),
            Margin = new Avalonia.Thickness(0, 2, 0, 2),
        };

        _tree.Background      = Brushes.Transparent;
        _tree.BorderThickness = new Avalonia.Thickness(0);

        return new Border
        {
            Padding = new Avalonia.Thickness(8),
            Child   = new Grid
            {
                RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
                Children =
                {
                    headerGrid,
                    new Border { [Grid.RowProperty]=1, Height=1,
                                 Background=Brush.Parse("#333"),
                                 Margin=new Avalonia.Thickness(0,0,0,3) },
                    _txbStatus,
                    new ScrollViewer
                    {
                        [Grid.RowProperty] = 3,
                        VerticalScrollBarVisibility =
                            Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                        Content = _tree,
                    }
                }
            }
        };
    }

    // ══════════════════════════════════════════════════════════════════════
    //  외부 API — Content1 노드 선택 시 호출
    // ══════════════════════════════════════════════════════════════════════
    public async void AddRecord(AnalysisRequestRecord rec)
    {
        string key = $"{rec.접수번호}:{rec.Id}";
        if (_addedKeys.Contains(key)) return;
        _addedKeys.Add(key);

        Dictionary<string, string> row;
        try { row = await Task.Run(() => AnalysisRequestService.GetRecordRow(rec.Id)); }
        catch { row = new Dictionary<string, string>(); }

        _rowCache[key] = row;

        var fixedCols = FixedColsForAnalyte();

        var analyteItems = row
            .Where(kv => !fixedCols.Contains(kv.Key.Trim()) &&
                         !string.IsNullOrWhiteSpace(kv.Value))
            .ToList();

        // ── 부모 노드 ────────────────────────────────────────────────────
        var parentNode = new TreeViewItem
        {
            Header     = MakeRecordHeader(rec, row),
            Tag        = new ParentTag(rec),
            IsExpanded = false,
        };

        bool isAnalyzing(string v) =>
            v.Trim().Equals("O", StringComparison.OrdinalIgnoreCase);

        // ── 자식 노드 (분석항목) ─────────────────────────────────────────
        foreach (var kv in analyteItems)
        {
            string col     = kv.Key.Trim();
            bool   pending = isAnalyzing(kv.Value);

            var childGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,60"),
                Background        = BrushChildBg,
            };
            childGrid.Children.Add(new TextBlock
            {
                Text = col, FontSize = 10, FontFamily = Font,
                Foreground = BrushChildFg,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(8, 2),
                [Grid.ColumnProperty] = 0,
            });
            childGrid.Children.Add(new TextBlock
            {
                Text = pending ? "🔴 분석중" : kv.Value,
                FontSize = 9, FontFamily = Font,
                Foreground = pending ? Brush.Parse("#ff6666") : Brush.Parse("#88cc88"),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(4, 2),
                [Grid.ColumnProperty] = 1,
            });

            // 자식 Tag에 분석항목 전체명 저장
            var childNode = new TreeViewItem
            {
                Header = childGrid,
                Tag    = new ChildTag(col),
            };
            AttachContextMenu(childNode, isParent: false);
            parentNode.Items.Add(childNode);
        }

        AttachContextMenu(parentNode, isParent: true);
        _tree.Items.Add(parentNode);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  📌 TODO 전송
    //  부모노드 → 중간 Task  /  자식노드 → 최종 Task
    // ══════════════════════════════════════════════════════════════════════
    private async void BtnTodo_Click(object? sender, RoutedEventArgs e)
    {
        var parentNodes = GetParentNodes();
        if (parentNodes.Count == 0)
        {
            SetStatus("⚠ 전송할 항목이 없습니다.");
            return;
        }

        SetStatus("⏳ TODO 전송 중...");

        // (rec, row, analyteNames) 목록 구성
        var items = new List<(AnalysisRequestRecord rec,
                               Dictionary<string, string> row,
                               List<string> analyteNames)>();

        foreach (var node in parentNodes)
        {
            var rec = ((ParentTag)node.Tag!).Rec;
            string key = $"{rec.접수번호}:{rec.Id}";

            if (!_rowCache.TryGetValue(key, out var row))
            {
                try { row = await Task.Run(() => AnalysisRequestService.GetRecordRow(rec.Id)); }
                catch { row = new Dictionary<string, string>(); }
                _rowCache[key] = row;
            }

            // 자식 노드에서 분석항목 전체명 목록 수집
            var analyteNames = node.Items
                .OfType<TreeViewItem>()
                .Where(c => c.Tag is ChildTag)
                .Select(c => ((ChildTag)c.Tag!).AnalyteName)
                .ToList();

            items.Add((rec, row, analyteNames));
        }

        try
        {
            await TodoService.SendAnalysisRequestsToTodoAsync(items);

            int totalTasks = items.Sum(i => i.analyteNames.Count);
            SetStatus($"✅ 중간Task {items.Count}개 + 최종Task {totalTasks}개 전송 완료 ({DateTime.Now:HH:mm})");
        }
        catch (Exception ex)
        {
            var msg = ex.Message.Length > 50 ? ex.Message[..50] + "…" : ex.Message;
            SetStatus($"❌ 전송 실패: {msg}");
            System.Diagnostics.Debug.WriteLine($"[BtnTodo_Click] 오류: {ex}");
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  🖨 분석의뢰서 출력
    // ══════════════════════════════════════════════════════════════════════
    private async void BtnPrint_Click(object? sender, RoutedEventArgs e)
    {
        List<TreeViewItem> targets;
        if (_tree.SelectedItem is TreeViewItem sel && sel.Tag is ParentTag)
            targets = new List<TreeViewItem> { sel };
        else
            targets = GetParentNodes();

        if (targets.Count == 0)
        {
            await ShowMessageAsync("출력할 항목이 없습니다.", "분석의뢰서 출력");
            return;
        }

        foreach (var node in targets)
        {
            var rec = ((ParentTag)node.Tag!).Rec;
            string key = $"{rec.접수번호}:{rec.Id}";
            if (!_rowCache.TryGetValue(key, out var row))
            {
                try { row = await Task.Run(() => AnalysisRequestService.GetRecordRow(rec.Id)); }
                catch { row = new Dictionary<string, string>(); }
            }
            await PrintAnalysisRequestAsync(rec, row);
        }
    }

    private async Task PrintAnalysisRequestAsync(
        AnalysisRequestRecord rec, Dictionary<string, string> row)
    {
        // ── 출력 연동 포인트 ─────────────────────────────────────────────
        // await ReportService.PrintAnalysisRequestAsync(rec, row);
        // ────────────────────────────────────────────────────────────────

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("[ 분석의뢰서 출력 ]");
        sb.AppendLine($"약칭     : {rec.약칭}");
        sb.AppendLine($"시료명   : {rec.시료명}");
        sb.AppendLine($"접수번호 : {rec.접수번호}");
        sb.AppendLine($"의뢰일   : {rec.의뢰일}");
        if (row.TryGetValue("채취일자", out var cd) && !string.IsNullOrWhiteSpace(cd))
            sb.AppendLine($"채취일자 : {cd}");
        if (row.TryGetValue("담당자",   out var mg) && !string.IsNullOrWhiteSpace(mg))
            sb.AppendLine($"담당자   : {mg}");
        int cnt = row.Count(kv =>
            !FixedColsForAnalyte().Contains(kv.Key.Trim()) &&
            !string.IsNullOrWhiteSpace(kv.Value));
        sb.AppendLine($"분석항목 : {cnt}종");

        await ShowMessageAsync(sb.ToString(), "분석의뢰서 출력 확인");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  우클릭 컨텍스트 메뉴 — 삭제 확인
    // ══════════════════════════════════════════════════════════════════════
    private void AttachContextMenu(TreeViewItem node, bool isParent)
    {
        var menuDelete = new MenuItem
        {
            Header = isParent ? "🗑  이 의뢰 삭제 (하위 항목 포함)" : "🗑  이 항목 삭제",
        };
        menuDelete.Click += (_, _) => ConfirmAndDelete(node, isParent);
        var ctx = new ContextMenu();
        ctx.Items.Add(menuDelete);
        node.ContextMenu = ctx;
    }

    private async void ConfirmAndDelete(TreeViewItem node, bool isParent)
    {
        var dlg = new Window
        {
            Title = "삭제 확인", Width = 320, Height = 140,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false, Background = Brush.Parse("#1a1a2a"),
            SystemDecorations = Avalonia.Controls.SystemDecorations.Full,
        };

        var msg = isParent
            ? "이 의뢰와 모든 분석항목을 목록에서 삭제할까요?"
            : "이 분석항목을 목록에서 삭제할까요?";

        var btnOk = new Button
        {
            Content = "삭제", Width = 80, Height = 28, FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#4a1a1a"), Foreground = Brush.Parse("#f0aeae"),
            BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        var btnCancel = new Button
        {
            Content = "취소", Width = 80, Height = 28, FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#aaa"),
            BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            Margin = new Avalonia.Thickness(10, 0, 0, 0),
        };

        bool confirmed = false;
        btnOk.Click     += (_, _) => { confirmed = true; dlg.Close(); };
        btnCancel.Click += (_, _) => dlg.Close();

        dlg.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = msg, FontSize = 11, FontFamily = Font,
                    Foreground = Brush.Parse("#dddddd"),
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Children = { btnOk, btnCancel },
                },
            },
        };

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();

        if (!confirmed) return;

        if (isParent)
        {
            RemoveParentKey(node);
            _tree.Items.Remove(node);
        }
        else
        {
            foreach (TreeViewItem parent in _tree.Items.OfType<TreeViewItem>())
            {
                if (!parent.Items.Contains(node)) continue;
                parent.Items.Remove(node);
                if (parent.Items.Count == 0)
                {
                    RemoveParentKey(parent);
                    _tree.Items.Remove(parent);
                }
                break;
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  부모 노드 헤더 UI — 채취일자·담당자 표시 포함
    // ══════════════════════════════════════════════════════════════════════
    private Control MakeRecordHeader(AnalysisRequestRecord rec, Dictionary<string, string> row)
    {
        string sampleDisp = "";
        if (row.TryGetValue("채취일자", out var rawDate) && !string.IsNullOrWhiteSpace(rawDate))
            sampleDisp = DateTime.TryParse(rawDate, out var dt)
                ? dt.ToString("MM/dd")
                : (rawDate.Length >= 10 ? rawDate[5..10] : rawDate);

        row.TryGetValue("담당자", out var manager);

        var sp = new StackPanel { Spacing = 1, Margin = new Avalonia.Thickness(2, 3) };

        // 1행: [약칭] 시료명  의뢰일
        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
        topRow.Children.Add(new Border
        {
            Background = BrushAbbrBg, CornerRadius = new Avalonia.CornerRadius(3),
            Padding = new Avalonia.Thickness(4, 1), Margin = new Avalonia.Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 0,
            Child = new TextBlock { Text = rec.약칭, FontSize = 9, FontFamily = Font, Foreground = BrushAbbrFg },
        });
        topRow.Children.Add(new TextBlock
        {
            Text = rec.시료명, FontSize = 11, FontFamily = Font, Foreground = BrushParentFg,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        });
        topRow.Children.Add(new TextBlock
        {
            Text = rec.의뢰일.Length >= 10 ? rec.의뢰일[..10] : rec.의뢰일,
            FontSize = 9, FontFamily = Font, Foreground = BrushAccNo,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(6, 0),
            [Grid.ColumnProperty] = 2,
        });
        sp.Children.Add(topRow);

        // 2행: 접수번호  채취 MM/dd  담당자
        var subRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,Auto,*") };
        subRow.Children.Add(new TextBlock
        {
            Text = rec.접수번호, FontSize = 9, FontFamily = Font,
            Foreground = BrushAccNo, Margin = new Avalonia.Thickness(0, 0, 8, 1),
            [Grid.ColumnProperty] = 0,
        });
        if (!string.IsNullOrEmpty(sampleDisp))
            subRow.Children.Add(new TextBlock
            {
                Text = $"채취 {sampleDisp}", FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse("#8888cc"),
                Margin = new Avalonia.Thickness(0, 0, 8, 1),
                [Grid.ColumnProperty] = 1,
            });
        if (!string.IsNullOrWhiteSpace(manager))
            subRow.Children.Add(new TextBlock
            {
                Text = manager, FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse("#88aa88"),
                Margin = new Avalonia.Thickness(0, 0, 0, 1),
                [Grid.ColumnProperty] = 2,
            });
        sp.Children.Add(subRow);

        return sp;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  공통 헬퍼
    // ══════════════════════════════════════════════════════════════════════

    private List<TreeViewItem> GetParentNodes() =>
        _tree.Items.OfType<TreeViewItem>()
             .Where(n => n.Tag is ParentTag)
             .ToList();

    private void RemoveParentKey(TreeViewItem node)
    {
        if (node.Tag is ParentTag pt)
        {
            string key = $"{pt.Rec.접수번호}:{pt.Rec.Id}";
            _addedKeys.Remove(key);
            _rowCache.Remove(key);
        }
    }

    private void SetStatus(string msg)
    {
        if (_txbStatus != null) _txbStatus.Text = msg;
    }

    private static HashSet<string> FixedColsForAnalyte() =>
        new(StringComparer.OrdinalIgnoreCase)
        {
            "약칭","시료명","접수번호","의뢰일","업체명","대표자",
            "담당자","연락처","이메일","견적번호","비고",
            "채취일자","채취시간","의뢰사업장","입회자",
            "시료채취자-1","시료채취자-2","방류허용기준 적용유무",
            "정도보증유무","분석완료일자","견적구분",
        };

    private async Task ShowMessageAsync(string message, string title = "알림")
    {
        var dlg = new Window
        {
            Title = title, Width = 360, Height = 220,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false, Background = Brush.Parse("#1a1a2a"),
            SystemDecorations = Avalonia.Controls.SystemDecorations.Full,
        };
        var btnOk = new Button
        {
            Content = "확인", Width = 80, Height = 28, FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#2a3a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        btnOk.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = message, FontSize = 11, FontFamily = Font,
                    Foreground = Brush.Parse("#dddddd"),
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Children = { btnOk },
                },
            },
        };
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  Tag 타입 — 강타입으로 부모/자식 구분 (ValueTuple 대신)
    // ══════════════════════════════════════════════════════════════════════

    private sealed class ParentTag
    {
        public AnalysisRequestRecord Rec { get; }
        public ParentTag(AnalysisRequestRecord rec) => Rec = rec;
    }

    private sealed class ChildTag
    {
        public string AnalyteName { get; } // 분석항목 전체명 (분장표준처리 컬럼 헤더와 매칭)
        public ChildTag(string analyteName) => AnalyteName = analyteName;
    }
}
