using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Views;
using ETA.Views.Controls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

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
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static readonly IBrush BrushChildFg  = AppRes("FgMuted");
    private static readonly IBrush BrushAbbrBg   = AppTheme.BgDanger;
    private static readonly IBrush BrushAbbrFg   = Brush.Parse("#cc8888");
    private static readonly IBrush BrushParentFg = AppTheme.FgSuccess;

    private readonly TreeView _tree = new();
    private readonly HashSet<string> _addedKeys = new();

    // DB row 캐시: "접수번호:Id" → row (TODO 전송/출력 시 재사용)
    private readonly Dictionary<string, Dictionary<string, string>> _rowCache = new();

    // 상태 표시
    private TextBlock?    _txbStatus;
    private ProgressBar?  _progressBar;

    // 측정인 전송 버튼
    private Button _btnMeasurer = new();

    // Show4 → Show2: 부모 노드 클릭 시 상세 패널 요청
    public event Action<AnalysisRequestRecord>? RecordSelected;

    public AnalysisRequestListPanel()
    {
        Content = BuildUI();
    }

    // wire-v01: pill 버튼 (StatusBadge 토큰 기반)
    private static Button MakePill(string text, BadgeStatus status, int padX = 12)
    {
        var (bg, fg, bd) = StatusBadge.GetBrushes(status);
        return new Button
        {
            Content = text, Height = 26, FontSize = AppTheme.FontSM,
            FontFamily = new("avares://ETA/Assets/Fonts#Pretendard"),
            FontWeight = FontWeight.SemiBold,
            Background = bg, Foreground = fg, BorderBrush = bd,
            BorderThickness = new Avalonia.Thickness(1),
            CornerRadius = new Avalonia.CornerRadius(999),
            Padding = new Avalonia.Thickness(padX, 0),
            Margin = new Avalonia.Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Cursor = new Cursor(StandardCursorType.Hand),
        };
    }

    private Control BuildUI()
    {
        // ── 1행: 제목 ────────────────────────────────────────────────────
        var titleRow = new TextBlock
        {
            Text = "📋  분석의뢰 선택 목록",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(0, 0, 0, 4),
        };

        // ── 2행: 버튼 모음 ───────────────────────────────────────────────
        var btnTodo  = MakePill("📌 TODO 전송",  BadgeStatus.Warn);
        var btnPrint = MakePill("🖨 의뢰서 출력", BadgeStatus.Info);
        var btnClear = MakePill("전체 삭제",       BadgeStatus.Bad, padX: 10);
        _btnMeasurer = MakePill("🌐 측정인 전송", BadgeStatus.Ok);

        btnTodo.Click    += BtnTodo_Click;
        btnPrint.Click   += BtnPrint_Click;
        btnClear.Click   += (_, _) =>
        {
            _tree.Items.Clear();
            _addedKeys.Clear();
            _rowCache.Clear();
            UpdateStatusCount();
        };
        _btnMeasurer.Click += BtnMeasurer_Click;

        var btnRow = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Avalonia.Thickness(0, 0, 0, 4),
            Children = { btnTodo, btnPrint, btnClear, _btnMeasurer },
        };

        var headerPanel = new StackPanel { Spacing = 0 };
        headerPanel.Children.Add(titleRow);
        headerPanel.Children.Add(btnRow);

        _txbStatus = new TextBlock
        {
            Text = "",
            FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppRes("TreeFg"),
            Margin = new Avalonia.Thickness(0, 2, 0, 2),
            [Grid.RowProperty] = 3,
        };

        _tree.Background      = Brushes.Transparent;
        _tree.BorderThickness = new Avalonia.Thickness(0);

        // 드래그앤드롭으로 순서 변경
        DragDrop.SetAllowDrop(_tree, true);
        _tree.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("reorder-node") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });
        _tree.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
        {
            if (!e.Data.Contains("reorder-node")) return;
            var draggedNode = e.Data.Get("reorder-node") as TreeViewItem;
            if (draggedNode == null) return;

            // 드롭 위치의 TreeViewItem 찾기
            var dropTarget = FindTreeViewItemAtPoint(e.GetPosition(_tree));
            if (dropTarget == null || dropTarget == draggedNode) return;

            // 부모 노드만 재정렬 (자식 노드 무시)
            int dragIdx = _tree.Items.IndexOf(draggedNode);
            int dropIdx = _tree.Items.IndexOf(dropTarget);
            if (dragIdx < 0 || dropIdx < 0) return;

            _tree.Items.RemoveAt(dragIdx);
            _tree.Items.Insert(dropIdx, draggedNode);
        });

        return new Border
        {
            Padding = new Avalonia.Thickness(8),
            Child   = new Grid
            {
                RowDefinitions = new RowDefinitions("Auto,Auto,*,Auto"),
                Children =
                {
                    headerPanel,
                    new Border { [Grid.RowProperty]=1, Height=1,
                                 Background=AppRes("InputBorder"),
                                 Margin=new Avalonia.Thickness(0,0,0,3) },
                    new ScrollViewer
                    {
                        [Grid.RowProperty] = 2,
                        VerticalScrollBarVisibility =
                            Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                        Content = _tree,
                    },
                    _txbStatus,
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
                ColumnDefinitions = new ColumnDefinitions("*,Auto"),
                Background        = AppRes("PanelInnerBg"),
            };
            childGrid.Children.Add(new TextBlock
            {
                Text = col, FontSize = AppTheme.FontSM, FontFamily = Font,
                Foreground = BrushChildFg,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(8, 2),
                [Grid.ColumnProperty] = 0,
            });
            var statusBadge = pending
                ? StatusBadge.Bad("분석중")
                : StatusBadge.Ok(kv.Value, withIcon: false);
            statusBadge.HorizontalAlignment = HorizontalAlignment.Right;
            statusBadge.VerticalAlignment   = VerticalAlignment.Center;
            statusBadge.Margin = new Avalonia.Thickness(4, 2, 6, 2);
            Grid.SetColumn(statusBadge, 1);
            childGrid.Children.Add(statusBadge);

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
        UpdateStatusCount();
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
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  🖨 분석기록부 출력
    // ══════════════════════════════════════════════════════════════════════
    private async void BtnPrint_Click(object? sender, RoutedEventArgs e)
    {
        var parentNodes = GetParentNodes();
        if (parentNodes.Count == 0)
        {
            await ShowMessageAsync("출력할 항목이 없습니다.", "분석기록부 출력");
            return;
        }

        // 1) 분석항목별 (rec, row) 목록 구성
        var groups = new Dictionary<string, List<(AnalysisRequestRecord rec, Dictionary<string, string> row)>>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var node in parentNodes)
        {
            var rec = ((ParentTag)node.Tag!).Rec;
            string key = $"{rec.접수번호}:{rec.Id}";
            if (!_rowCache.TryGetValue(key, out var row))
            {
                try { row = await Task.Run(() => AnalysisRequestService.GetRecordRow(rec.Id)); }
                catch { row = new Dictionary<string, string>(); }
                _rowCache[key] = row ?? new Dictionary<string, string>();
            }

            var analyteNames = node.Items
                .OfType<TreeViewItem>()
                .Where(c => c.Tag is ChildTag)
                .Select(c => ((ChildTag)c.Tag!).AnalyteName)
                .ToList();

            foreach (var analyteName in analyteNames)
            {
                if (!groups.ContainsKey(analyteName))
                    groups[analyteName] = new();
                groups[analyteName].Add((rec, _rowCache[key]));
            }
        }

        if (groups.Count == 0)
        {
            await ShowMessageAsync("분석항목이 없습니다.", "분석기록부 출력");
            return;
        }

        // 2) 방류기준표 로드 및 출력 디렉터리
        var 방류기준표 = await Task.Run(() => AnalysisRecordService.Load방류기준표());
        string outputDir = AnalysisRecordService.GetOutputDir();

        // 3) 프로그레스 창 — 분석자별 진행 바
        var today = DateTime.Today;

        var progressDlg = new Window
        {
            Title = "분석기록부 출력 중",
            Width = 420,
            Height = 160,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
            Background = AppRes("PanelBg"),
            SystemDecorations = Avalonia.Controls.SystemDecorations.Full,
        };

        var overallBar = new ProgressBar
        {
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 14, IsIndeterminate = true,
            Foreground = AppRes("TreeFg"),
            Background = AppRes("SubBtnBg"),
        };
        var statusLbl = new TextBlock
        {
            Text = $"분장 조회 중... ({today:yyyy-MM-dd} 기준)",
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppRes("FgMuted"),
            Margin = new Avalonia.Thickness(0, 4, 0, 0),
        };

        progressDlg.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(16, 16),
            Spacing = 6,
            Children =
            {
                new TextBlock
                {
                    Text = $"분석항목 {groups.Count}개 — 분석자별 기록부 생성",
                    FontSize = AppTheme.FontBase, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                },
                overallBar,
                statusLbl,
            },
        };

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) progressDlg.Show(owner);
        else progressDlg.Show();

        // 4) groups의 row dict 구성 (rec 정보 보완)
        var finalGroups = new Dictionary<string, List<Dictionary<string, string>>>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var kv in groups)
        {
            finalGroups[kv.Key] = kv.Value.Select(x =>
            {
                var d = new Dictionary<string, string>(x.row, StringComparer.OrdinalIgnoreCase);
                if (!d.ContainsKey("견적번호") || string.IsNullOrEmpty(d["견적번호"]))
                    d["견적번호"] = x.rec.접수번호;
                if (!d.ContainsKey("약칭") || string.IsNullOrEmpty(d["약칭"]))
                    d["약칭"] = x.rec.약칭;
                if (!d.ContainsKey("시료명") || string.IsNullOrEmpty(d["시료명"]))
                    d["시료명"] = x.rec.시료명;
                return d;
            }).ToList();
        }

        // 5) 분석자별 파일 생성
        var results = await Task.Run(() =>
            AnalysisRecordService.GenerateByAssignee(finalGroups, 방류기준표, outputDir, today));

        overallBar.IsIndeterminate = false;
        overallBar.Value = 100;

        if (results.Count > 0)
        {
            var names = string.Join(", ", results.Select(r => r.Assignee));
            statusLbl.Text = $"완료: {names} ({results.Count}개 파일)";
            statusLbl.Foreground = AppTheme.FgSuccess;
        }
        else
        {
            statusLbl.Text = "생성 실패 — 템플릿 파일을 확인하세요.";
            statusLbl.Foreground = AppTheme.FgDanger;
        }

        await Task.Delay(1200);
        progressDlg.Close();

        if (results.Count > 0)
        {
            try { Process.Start(new ProcessStartInfo(outputDir) { UseShellExecute = true }); }
            catch { }
            var summary = string.Join(", ", results.Select(r => r.Assignee));
            SetStatus($"✅ 분석기록부 {results.Count}개 생성 완료 ({summary}) {DateTime.Now:HH:mm}");
        }
        else
        {
            await ShowMessageAsync("파일 생성에 실패했습니다.\n템플릿 파일 또는 분장 데이터를 확인하세요.", "분석기록부 출력");
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  우클릭 컨텍스트 메뉴 — 삭제 확인
    // ══════════════════════════════════════════════════════════════════════
    private void AttachContextMenu(TreeViewItem node, bool isParent)
    {
        var menuDelete = new MenuItem
        {
            Header = isParent ? "\ud83d\uddd1  \uc774 \uc758\ub8b0 \uc0ad\uc81c (\ud558\uc704 \ud56d\ubaa9 \ud3ec\ud568)" : "\ud83d\uddd1  \uc774 \ud56d\ubaa9 \uc0ad\uc81c",
        };
        menuDelete.Click += (_, _) => ConfirmAndDelete(node, isParent);
        var ctx = new ContextMenu();
        ctx.Items.Add(menuDelete);
        node.ContextMenu = ctx;

        // 부모 노드에 드래그 재정렬 + 클릭 시 Show2 전달
        if (isParent)
        {
            AttachDragReorder(node);
            node.Tapped += (_, _) =>
            {
                if (node.Tag is ParentTag pt)
                    RecordSelected?.Invoke(pt.Rec);
            };
        }
    }

    /// <summary>TreeViewItem에 드래그 시작 핸들러 추가 (순서 변경용)</summary>
    private void AttachDragReorder(TreeViewItem node)
    {
        Point? pressPos    = null;
        bool   dragStarted = false;

        node.PointerPressed += (s, e) =>
        {
            if (!e.GetCurrentPoint(node).Properties.IsLeftButtonPressed) return;
            pressPos    = e.GetCurrentPoint(node).Position;
            dragStarted = false;
        };

        node.PointerMoved += async (s, e) =>
        {
            if (pressPos == null || dragStarted) return;
            if (!e.GetCurrentPoint(node).Properties.IsLeftButtonPressed)
            {
                pressPos = null;
                return;
            }
            var diff = e.GetCurrentPoint(node).Position - pressPos.Value;
            if (Math.Abs(diff.X) > 5 || Math.Abs(diff.Y) > 5)
            {
                dragStarted = true;
                pressPos    = null;
                var data = new DataObject();
                data.Set("reorder-node", node);
                await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                dragStarted = false;
            }
        };

        node.PointerReleased += (s, e) =>
        {
            pressPos    = null;
            dragStarted = false;
        };
    }

    /// <summary>TreeView 내 특정 위치의 TreeViewItem 찾기</summary>
    private TreeViewItem? FindTreeViewItemAtPoint(Point pt)
    {
        foreach (TreeViewItem item in _tree.Items.OfType<TreeViewItem>())
        {
            var pos = item.TranslatePoint(new Point(0, 0), _tree);
            if (pos == null) continue;
            if (pt.Y >= pos.Value.Y && pt.Y <= pos.Value.Y + item.Bounds.Height)
                return item;
        }
        return null;
    }

    private async void ConfirmAndDelete(TreeViewItem node, bool isParent)
    {
        var dlg = new Window
        {
            Title = "삭제 확인", Width = 320, Height = 140,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false, Background = AppRes("PanelBg"),
            SystemDecorations = Avalonia.Controls.SystemDecorations.Full,
        };

        var msg = isParent
            ? "이 의뢰와 모든 분석항목을 목록에서 삭제할까요?"
            : "이 분석항목을 목록에서 삭제할까요?";

        var btnOk = new Button
        {
            Content = "삭제", Width = 80, Height = 28, FontSize = AppTheme.FontBase, FontFamily = Font,
            Background = Brush.Parse("#4a1a1a"), Foreground = AppTheme.FgDanger,
            BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        var btnCancel = new Button
        {
            Content = "취소", Width = 80, Height = 28, FontSize = AppTheme.FontBase, FontFamily = Font,
            Background = AppRes("SubBtnBg"), Foreground = AppRes("FgMuted"),
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
                new TextBlock { Text = msg, FontSize = AppTheme.FontBase, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
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
        UpdateStatusCount();
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
        var (abg, afg) = BadgeColorHelper.GetBadgeColor(rec.약칭);
        topRow.Children.Add(new Border
        {
            Background = Brush.Parse(abg), CornerRadius = new Avalonia.CornerRadius(3),
            Padding = new Avalonia.Thickness(4, 1), Margin = new Avalonia.Thickness(0, 0, 5, 0),
            [Grid.ColumnProperty] = 0,
            Child = new TextBlock { Text = rec.약칭, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brush.Parse(afg) },
        });
        topRow.Children.Add(new TextBlock
        {
            Text = rec.시료명, FontSize = AppTheme.FontBase, FontFamily = Font, Foreground = BrushParentFg,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 1,
        });
        topRow.Children.Add(new TextBlock
        {
            Text = rec.의뢰일.Length >= 10 ? rec.의뢰일[..10] : rec.의뢰일,
            FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(6, 0),
            [Grid.ColumnProperty] = 2,
        });
        sp.Children.Add(topRow);

        // 2행: 접수번호  채취 MM/dd  담당자
        var subRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,Auto,*") };
        subRow.Children.Add(new TextBlock
        {
            Text = rec.접수번호, FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppRes("FgMuted"), Margin = new Avalonia.Thickness(0, 0, 8, 1),
            [Grid.ColumnProperty] = 0,
        });
        if (!string.IsNullOrEmpty(sampleDisp))
            subRow.Children.Add(new TextBlock
            {
                Text = $"채취 {sampleDisp}", FontSize = AppTheme.FontXS, FontFamily = Font,
                Foreground = AppRes("FgMuted"),
                Margin = new Avalonia.Thickness(0, 0, 8, 1),
                [Grid.ColumnProperty] = 1,
            });
        if (!string.IsNullOrWhiteSpace(manager))
            subRow.Children.Add(new TextBlock
            {
                Text = manager, FontSize = AppTheme.FontXS, FontFamily = Font,
                Foreground = AppRes("TreeFg"),
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
        Log(msg);
    }

    private static readonly object _logLock = new();
    private static void Log(string msg)
    {
        try
        {
            Directory.CreateDirectory("Logs");
            lock (_logLock)
            {
                File.AppendAllText("Logs/MeasurerDebug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}" + Environment.NewLine);
            }
        }
        catch { }
    }

    private void UpdateStatusCount()
    {
        var parentNodes = GetParentNodes();
        if (parentNodes.Count == 0)
        {
            SetStatus("");
            return;
        }

        int totalAnalytes = parentNodes.Sum(p => p.Items.Count);
        SetStatus($"📊 의뢰 {parentNodes.Count}개 + 분석항목 {totalAnalytes}개 수집됨 ({DateTime.Now:HH:mm})");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  🌐 측정인 전송 — CDP로 의뢰내역 주입
    // ══════════════════════════════════════════════════════════════════════
    private async void BtnMeasurer_Click(object? sender, RoutedEventArgs e)
    {
        Log("═══ 측정인 전송 시작 ═══");
        var parentNodes = GetParentNodes();
        Log($"parentNodes.Count = {parentNodes.Count}");
        if (parentNodes.Count == 0)
        {
            SetStatus("⚠ 전송할 의뢰 항목이 없습니다.");
            return;
        }

        // ── 사전 설정 다이얼로그 ──────────────────────────────────────
        var allAgents = AgentService.GetAllItems();
        var dialogRecords = parentNodes.Select(n =>
        {
            var pt = (ParentTag)n.Tag!;
            var analytes = n.Items.OfType<TreeViewItem>()
                .Where(c => c.Tag is ChildTag)
                .Select(c => ((ChildTag)c.Tag!).AnalyteName)
                .ToArray();
            string key = $"{pt.Rec.접수번호}:{pt.Rec.Id}";
            _rowCache.TryGetValue(key, out var row);
            string? company = null;
            row?.TryGetValue("업체명", out company);
            return (sample: pt.Rec.시료명, analytes: string.Join(", ", analytes), company: company ?? pt.Rec.약칭);
        }).ToList();

        // 캘린더/DB 시료채취자 → 사전 선택 이름 수집
        var preNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var n in parentNodes)
        {
            var pt = (ParentTag)n.Tag!;
            string key = $"{pt.Rec.접수번호}:{pt.Rec.Id}";
            if (!_rowCache.TryGetValue(key, out var row) || row == null) continue;
            foreach (var col in new[] { "시료채취자-1", "시료채취자-2", "시료채취1", "시료채취2", "채수담당자", "시료채취자1", "시료채취자2" })
            {
                if (!row.TryGetValue(col, out var raw) || string.IsNullOrWhiteSpace(raw)) continue;
                foreach (var tok in raw.Split(new[] { ',', '/', ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var t = tok.Trim();
                    if (t.Length > 0) preNames.Add(t);
                }
            }
        }
        Log($"preNames=[{string.Join(",", preNames)}]");

        var owner = TopLevel.GetTopLevel(this) as Window;
        var dlg   = new MeasurerSendDialog(dialogRecords, allAgents, preNames.ToList());
        if (owner != null)
            await dlg.ShowDialog(owner);
        else
        {
            var tcs = new TaskCompletionSource();
            dlg.Closed += (_, _) => tcs.TrySetResult();
            dlg.Show();
            await tcs.Task;
        }
        if (!dlg.Confirmed) { SetStatus("⚠ 취소됨"); return; }

        var purposeValues   = dlg.PurposeValues;      // List<string> per-record
        var empIdsPerRecord = dlg.EmpIdsPerRecord;    // List<List<string>> per-record
        var selectedNames   = dlg.SelectedAgentNames;
        var selectedAgents  = dlg.SelectedAgents;
        Log($"dlg.Confirmed=true, purposeValues.Count={purposeValues?.Count ?? 0}, empIdsPerRecord.Count={empIdsPerRecord?.Count ?? 0}, selectedNames=[{string.Join(",", selectedNames)}]");
        if (empIdsPerRecord != null)
            for (int i = 0; i < empIdsPerRecord.Count; i++)
                Log($"  record[{i}]: purpose='{(purposeValues != null && i < purposeValues.Count ? purposeValues[i] : "?")}', empIds=[{string.Join(",", empIdsPerRecord[i])}]");

        // ── 시료채취자 DB 저장 + 캘린더 자동 등록 ────────────────────
        if (selectedNames.Count > 0)
        {
            // 의뢰별 시료채취자 컬럼 UPDATE
            foreach (var n in parentNodes)
            {
                var pt = (ParentTag)n.Tag!;
                try { AnalysisRequestService.UpdateSamplers(pt.Rec.Id, selectedNames); }
                catch (Exception ex) { Log($"UpdateSamplers 실패 rowId={pt.Rec.Id}: {ex.Message}"); }
            }

            // 캘린더(일정) 등록 — (채취일자, 약칭) 그룹 단위로 인원별 1건
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string registrar = ETA.Views.MainPage.CurrentEmployeeId ?? "";
            var groups = parentNodes
                .Select(n => (ParentTag)n.Tag!)
                .GroupBy(pt => (date: (pt.Rec.채취일자 ?? "").Trim(), abbr: (pt.Rec.약칭 ?? "").Trim()));
            foreach (var g in groups)
            {
                if (string.IsNullOrWhiteSpace(g.Key.date)) continue;
                foreach (var (name, empId) in selectedAgents)
                {
                    try
                    {
                        ScheduleService.Insert(new ScheduleEntry
                        {
                            날짜     = g.Key.date,
                            직원명   = name,
                            직원id   = empId,
                            분류     = "채수",
                            사이트   = "",
                            업체약칭 = g.Key.abbr,
                            제목     = "",
                            내용     = "",
                            시작시간 = "",
                            종료시간 = "",
                            첨부파일 = "",
                            등록일시 = now,
                            등록자   = registrar,
                        });
                        Log($"ScheduleService.Insert: {g.Key.date} {name}({empId}) 채수 {g.Key.abbr}");
                    }
                    catch (Exception ex) { Log($"ScheduleService.Insert 실패 {g.Key.date}/{name}: {ex.Message}"); }
                }
            }
        }
        // ─────────────────────────────────────────────────────────────

        SetStatus($"⏳ {parentNodes.Count}건 개별 전송 준비 중...");
        _btnMeasurer.IsEnabled = false;
        try
        {
            SetStatus("⏳ 측정인 로그인 창을 엽니다...");
            var loginWin = new MeasurerLoginWindow { CloseOnLoginSuccess = true };
            if (owner != null)
                await loginWin.ShowDialog(owner);
            else
            {
                var tcs = new TaskCompletionSource();
                loginWin.Closed += (_, _) => tcs.TrySetResult();
                loginWin.Show();
                await tcs.Task;
            }

            if (loginWin.LoginSucceeded)
            {
                Log("loginWin.LoginSucceeded=true");
                SetStatus($"⏳ 로그인 완료 — {parentNodes.Count}건 개별 전송 중...");
                await Task.Delay(1500); // 브라우저 다음 페이지 로드 대기
                bool injected = await TryInjectRequestDataAsync(parentNodes, purposeValues, empIdsPerRecord);
                Log($"TryInjectRequestDataAsync → {injected}");
                if (injected)
                    SetStatus($"✅ 의뢰계획 {parentNodes.Count}건 개별 전달 완료 ({DateTime.Now:HH:mm})");
                else
                    SetStatus("⚠ 로그인은 되었으나 탭 연결에 실패했습니다. 다시 시도해주세요.");
            }
            else
            {
                Log("loginWin.LoginSucceeded=false");
                SetStatus("⚠ 로그인 취소됨");
            }
        }
        catch (Exception ex)
        {
            Log($"EXCEPTION in BtnMeasurer_Click: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            var msg = ex.Message.Length > 60 ? ex.Message[..60] + "…" : ex.Message;
            SetStatus($"❌ 전달 실패: {msg}");
        }
        finally { _btnMeasurer.IsEnabled = true; Log("═══ 측정인 전송 종료 ═══"); }
    }

    /// <summary>
    /// CDP port 9222로 측정인 탭을 찾아 window.__etaRequestData에 의뢰내역 주입.
    /// 탭을 찾으면 true, 없으면 false.
    /// </summary>
    private async Task<bool> TryInjectRequestDataAsync(
        List<TreeViewItem> parentNodes,
        IReadOnlyList<string>? purposeValues = null,
        IReadOnlyList<IReadOnlyList<string>>? empIdsPerRecord = null)
    {
        const int port = 9222;
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };

        string tabsJson;
        try { tabsJson = await http.GetStringAsync($"http://127.0.0.1:{port}/json"); }
        catch (Exception ex) { Log($"CDP /json fetch 실패: {ex.Message}"); return false; }
        Log($"CDP tabs fetched, length={tabsJson.Length}");

        // 측정인.kr 탭 우선, 없으면 field_water 포함 탭
        using var jdoc = JsonDocument.Parse(tabsJson);
        string? wsUrl = null;
        foreach (var tab in jdoc.RootElement.EnumerateArray())
        {
            string url  = tab.TryGetProperty("url",  out var u)  ? u.GetString()  ?? "" : "";
            string type = tab.TryGetProperty("type", out var t)  ? t.GetString()  ?? "" : "";
            if (type != "page") continue;
            if (url.Contains("sijeong") || url.Contains("field_water"))
            {
                wsUrl = tab.TryGetProperty("webSocketDebuggerUrl", out var ws) ? ws.GetString() : null;
                break;
            }
        }
        // fallback: 아무 page 탭
        if (wsUrl == null)
        {
            foreach (var tab in jdoc.RootElement.EnumerateArray())
            {
                string type = tab.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";
                if (type != "page") continue;
                wsUrl = tab.TryGetProperty("webSocketDebuggerUrl", out var ws) ? ws.GetString() : null;
                break;
            }
        }
        if (wsUrl == null) { Log("wsUrl null — 측정인 탭을 찾지 못함"); return false; }
        Log($"wsUrl 찾음: {wsUrl.Substring(0, Math.Min(80, wsUrl.Length))}");

        // 레코드 데이터 직렬화
        var records = parentNodes.Select(n =>
        {
            var pt = (ParentTag)n.Tag!;
            string key = $"{pt.Rec.접수번호}:{pt.Rec.Id}";
            _rowCache.TryGetValue(key, out var row);
            row ??= new Dictionary<string, string>();

            var analytes = n.Items.OfType<TreeViewItem>()
                .Where(c => c.Tag is ChildTag)
                .Select(c => ((ChildTag)c.Tag!).AnalyteName)
                .ToArray();

            row.TryGetValue("업체명", out var company);
            row.TryGetValue("담당자", out var manager);
            row.TryGetValue("채취일자", out var sampleDate);
            row.TryGetValue("의뢰사업장", out var workSite);
            row.TryGetValue("비고", out var note);
            row.TryGetValue("견적번호", out var quoteNo);

            // 계약 DB의 C_ContractType(계약근거) + C_PlaceName(처리시설) 사용. 업체명→약칭 순으로 조회
            string finalContractNo = "";
            string finalPlaceName  = "";
            (string 계약번호, string 약칭, string 계약기간, string 업체명, string 채취지점명) finalContract = ("", "", "", "", "");
            try
            {
                var allContracts = ContractService.GetAllContracts();
                Contract? mainContract = null;
                if (!string.IsNullOrEmpty(company))
                    mainContract = allContracts.FirstOrDefault(c => c.C_CompanyName == company);
                // 업체명 조회 실패 시 약칭으로 폴백
                if (mainContract == null && !string.IsNullOrEmpty(pt.Rec.약칭))
                {
                    mainContract = allContracts.FirstOrDefault(c => c.C_Abbreviation == pt.Rec.약칭);
                    if (mainContract != null) Log($"[근거계약] 약칭 '{pt.Rec.약칭}'로 계약 찾음 (업체명='{company}')");
                }

                if (mainContract == null)
                    Log($"[근거계약] 계약 DB에 업체 없음 (업체명='{company}', 약칭='{pt.Rec.약칭}')");
                else if (string.IsNullOrEmpty(mainContract.C_ContractType))
                    Log($"[근거계약] '{mainContract.C_CompanyName}'의 C_ContractType 미설정 (계약업체 페이지에서 근거계약 선택 필요)");
                else
                {
                    finalContractNo = mainContract.C_ContractType;
                    finalPlaceName  = mainContract.C_PlaceName ?? "";
                    var basisContractData = MeasurerService.GetAllData()
                        .FirstOrDefault(d => d.계약번호 == finalContractNo);
                    if (!string.IsNullOrEmpty(basisContractData.계약번호))
                        finalContract = basisContractData;
                    else
                        Log($"[근거계약] 계약번호 '{finalContractNo}' 찾음, 그러나 측정인 DB에 부가정보 없음");
                }
            }
            catch (Exception ex) { Log($"[근거계약] 조회 오류: {ex.Message}"); }

            return new
            {
                id      = pt.Rec.Id,
                abbr    = pt.Rec.약칭,
                sample  = pt.Rec.시료명,
                accNo   = pt.Rec.접수번호,
                quoteNo,
                date    = pt.Rec.의뢰일,
                company,
                manager,
                sampleDate,
                workSite,
                note,
                contractNo = finalContractNo,
                contractCompany = finalContract.업체명,
                contractPoint = !string.IsNullOrWhiteSpace(finalPlaceName) ? finalPlaceName : finalContract.채취지점명,
                contractPeriod = finalContract.계약기간,
                contractLabel = string.Join(" / ", new[]
                {
                    finalContract.계약번호,
                    finalContract.업체명,
                    finalContract.채취지점명,
                    finalContract.계약기간,
                }.Where(x => !string.IsNullOrWhiteSpace(x))),
                planText = string.Join(", ", analytes),
                analytes,
            };
        }).ToList();

        var opts   = new JsonSerializerOptions
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        string jsonData = JsonSerializer.Serialize(records, opts);

        // ── 코드 매핑용 마스터 데이터 (루프 밖에서 1회 조회) ──
        var allMeasurerItems = MeasurerService.GetAllAnalysisItems();
        var allAgents = AgentService.GetAllItems();

        using var cts    = new CancellationTokenSource(TimeSpan.FromSeconds(60 * records.Count));
        using var socket = new ClientWebSocket();
        try { await socket.ConnectAsync(new Uri(wsUrl), cts.Token); }
        catch (Exception ex) { Log($"WebSocket connect 실패: {ex.Message}"); return false; }
        Log($"WebSocket 연결됨, records.Count={records.Count}");

        // ── 레코드별 반복: 각 부모 노드마다 모달 열기 → 입력 → 저장 ──
        for (int ri = 0; ri < records.Count; ri++)
        {
            var rec = records[ri];
            Log($"── record[{ri}] 시작: sample='{rec.sample}', abbr='{rec.abbr}', accNo='{rec.accNo}', analytes={rec.analytes.Length}건 ──");

            // ── 이 레코드의 분석항목 코드 매핑 (자기 자식만) ──
            var analyteCodes = new List<string>();
            foreach (var analyte in rec.analytes)
            {
                var mi = allMeasurerItems.FirstOrDefault(m =>
                    m.항목명.Equals(analyte, StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(mi.코드값))
                    mi = allMeasurerItems.FirstOrDefault(m =>
                        m.항목명.Contains(analyte, StringComparison.OrdinalIgnoreCase)
                        || analyte.Contains(m.항목명, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(mi.코드값))
                    analyteCodes.Add(mi.코드값);
                else
                    Log($"  ⚠ analyte '{analyte}' → 코드값 없음 (측정인_분석항목 미등록)");
            }
            Log($"  analyteCodes: [{string.Join(",", analyteCodes)}] ({analyteCodes.Count}/{rec.analytes.Length})");

            // ── 이 레코드의 인력 매핑 ──
            // 각 레코드마다 다이얼로그 선택 우선, fallback: 시료채취자 DB 조회
            var recEmpIds = empIdsPerRecord != null && ri < empIdsPerRecord.Count
                ? empIdsPerRecord[ri]
                : null;

            List<string> empIds;
            if (recEmpIds != null && recEmpIds.Count > 0)
            {
                empIds = recEmpIds.ToList();
            }
            else
            {
                empIds = new List<string>();
                var recPt = (ParentTag)parentNodes[ri].Tag!;
                string recKey = $"{recPt.Rec.접수번호}:{recPt.Rec.Id}";
                _rowCache.TryGetValue(recKey, out var recRow);
                recRow ??= new Dictionary<string, string>();
                recRow.TryGetValue("시료채취자-1", out var s1);
                recRow.TryGetValue("시료채취자-2", out var s2);
                foreach (var nm in new[] { s1, s2 })
                {
                    if (string.IsNullOrWhiteSpace(nm)) continue;
                    string prefix = nm.Trim().Length >= 3 ? nm.Trim()[..3] : nm.Trim();
                    var ag = allAgents.FirstOrDefault(a =>
                        a.성명.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrEmpty(a.측정인고유번호));
                    if (ag != null) empIds.Add(ag.측정인고유번호);
                }
            }
            Log($"  empIds: [{string.Join(",", empIds)}] ({empIds.Count}건)");

            // ── 모달 열기 (addFieldPlan 클릭) ──
            string openScript = $$"""
                (function(){
                    var n = document.getElementById('__eta_req_notice__');
                    if (!n) {
                        n = document.createElement('div');
                        n.id = '__eta_req_notice__';
                        n.style.cssText =
                            'position:fixed;top:12px;right:12px;z-index:99999;' +
                            'background:#1a6f2a;color:#fff;padding:10px 16px;' +
                            'border-radius:6px;font-size:13px;box-shadow:0 2px 8px #0005;';
                        document.body.appendChild(n);
                    }
                    n.textContent = 'ETA {{ri + 1}}/{{records.Count}} 처리 중...';
                    var btn = document.getElementById('addFieldPlan');
                    if (btn && btn.offsetParent !== null) { btn.click(); }
                    setTimeout(function(){ if(n && n.parentNode) n.parentNode.removeChild(n); }, 4000);
                    return 'OK';
                })()
                """;
            await CdpEvalAsync(socket, openScript, cts.Token);

            // ── 모달 대기 — add_meas_cont_no 옵션 존재 확인 ──
            await Task.Delay(1000, cts.Token);
            bool modalReady = false;
            for (int attempt = 0; !modalReady && attempt < 20; attempt++)
            {
                await Task.Delay(400, cts.Token);
                string chk = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
                    var el = document.getElementById('add_meas_cont_no');
                    return (el && el.options && el.options.length > 0) ? 'YES' : 'NO';
                })()", cts.Token));
                modalReady = chk == "YES";
            }
            Log($"  modalReady={modalReady}");
            if (!modalReady) { Log("  ⚠ modal 안 열림 → skip"); continue; }

            // ── A. 근거계약 선택 — 측정인_채취지점 DB 기준 매칭 ─────────────
            string contOptsRaw = await CdpEvalAsync(socket, @"(function(){
                var sel = document.getElementById('add_meas_cont_no');
                if (!sel) return '[]';
                var opts = [];
                for (var i = 0; i < sel.options.length; i++) {
                    var o = sel.options[i];
                    if (!o.value) continue;
                    opts.push({ value: o.value, text: o.text.trim() });
                }
                return JSON.stringify(opts);
            })()", cts.Token);
            string contOptsJson = ExtractCdpValue(contOptsRaw);

            string chosenContVal = "";
            if (!string.IsNullOrEmpty(contOptsJson) && contOptsJson != "[]")
            {
                // rec.contractNo = 계약 DB의 C_ContractType(사용자 선택 근거계약)
                string wNo = (rec.contractNo ?? "").Trim();
                if (!string.IsNullOrEmpty(wNo))
                {
                    string wNoLow = wNo.ToLowerInvariant();
                    using var cDoc = JsonDocument.Parse(contOptsJson);
                    foreach (var o in cDoc.RootElement.EnumerateArray())
                    {
                        string v = o.GetProperty("value").GetString() ?? "";
                        if (string.IsNullOrEmpty(v)) continue;
                        if (v.ToLowerInvariant() == wNoLow) { chosenContVal = v; break; }
                    }
                }
            }
            Log($"  계약매칭: DB계약번호='{rec.contractNo}' → chosenContVal='{chosenContVal}'"
                + (string.IsNullOrEmpty(chosenContVal) && !string.IsNullOrEmpty(rec.contractNo?.ToString())
                    ? " (드랍박스 옵션에 일치하는 계약번호 없음)" : ""));

            // VBA: option.selected=true → option.click() → change + select2 trigger
            if (!string.IsNullOrEmpty(chosenContVal))
            {
                string cv = chosenContVal.Replace("'", "\\'").Replace("\"", "\\\"");
                string selScript = $@"(function(){{
                    var sel = document.getElementById('add_meas_cont_no');
                    if (!sel) return 'NO_SEL';
                    var opts = sel.getElementsByTagName('option');
                    for (var i = 0; i < opts.length; i++) {{
                        if (opts[i].value == '{cv}') {{
                            opts[i].selected = true;
                            try {{ opts[i].click(); }} catch(e) {{}}
                            break;
                        }}
                    }}
                    sel.dispatchEvent(new Event('change'));
                    if (window.$ && window.$('#add_meas_cont_no').length) {{
                        var $el = window.$('#add_meas_cont_no');
                        $el.trigger({{ type:'select2:select', params:{{ data:{{ id:'{cv}', text: $el.find('option:selected').text() }} }} }});
                        $el.trigger('change');
                    }}
                    return 'OK';
                }})()";
                string selResult = ExtractCdpValue(await CdpEvalAsync(socket, selScript, cts.Token));
                Log($"  근거계약 선택 결과: {selResult} (value={chosenContVal})");

                // 현장 목록 로딩 대기 (최대 5초)
                for (int w = 0; w < 5000; w += 400)
                {
                    await Task.Delay(400, cts.Token);
                    string plcChk = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
                        var sel = document.getElementById('cmb_emis_cmpy_plc_no');
                        if (!sel || sel.options.length <= 1) return '0';
                        return String(sel.options.length);
                    })()", cts.Token));
                    if (int.TryParse(plcChk, out int pc) && pc > 1) break;
                }
            }

            // ── B. 현장 선택 (VBA: option.selected + change) ─────────────
            //   우선순위: 계약 DB의 C_PlaceName(= rec.contractPoint, 측정인 옵션 텍스트 그대로) → 시료 의뢰사업장
            string wsName = (rec.contractPoint ?? "").Trim();
            if (string.IsNullOrEmpty(wsName)) wsName = (rec.workSite ?? "").Trim();
            if (!string.IsNullOrEmpty(wsName))
            {
                string wsE = wsName.Replace("\\", "\\\\").Replace("'", "\\'");
                string plcScript = $@"(function(){{
                    var sel = document.getElementById('cmb_emis_cmpy_plc_no');
                    if (!sel || sel.options.length <= 1) return 'NO_OPT';
                    var want = '{wsE}'.toLowerCase();
                    for (var i = 0; i < sel.options.length; i++) {{
                        if (sel.options[i].text.toLowerCase().indexOf(want) >= 0) {{
                            sel.value = sel.options[i].value;
                            sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                            return 'OK:' + sel.options[i].text;
                        }}
                    }}
                    if (sel.options.length > 1) {{
                        sel.value = sel.options[1].value;
                        sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'FALLBACK';
                    }}
                    return 'NO_MATCH';
                }})()";
                string plcResult = ExtractCdpValue(await CdpEvalAsync(socket, plcScript, cts.Token));
                Log($"  현장 선택 결과: {plcResult} (want={wsName})");

                // 채취지점 로딩 대기
                for (int w = 0; w < 5000; w += 400)
                {
                    await Task.Delay(400, cts.Token);
                    string facChk = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
                        var el = document.getElementById('add_emis_fac_no');
                        if (!el || el.options.length <= 1) return '0';
                        return String(el.options.length);
                    })()", cts.Token));
                    if (int.TryParse(facChk, out int fc) && fc > 1) break;
                }
            }

            // ── C. 환경기술인 입력 (VBA CommandButton3: Clear.SendKeys) ──
            string techName = (rec.manager ?? "").Trim();
            if (!string.IsNullOrEmpty(techName))
            {
                string tnE = techName.Replace("'", "\\'");
                await CdpEvalAsync(socket, $@"(function(){{
                    var el = document.getElementById('add_env_psic_name');
                    if (!el) return 'NO';
                    el.value = '';
                    el.value = '{tnE}';
                    el.dispatchEvent(new Event('input', {{bubbles:true}}));
                    el.dispatchEvent(new Event('change', {{bubbles:true}}));
                    el.blur();
                    return 'OK';
                }})()", cts.Token);
            }

            // ── D. 분석시작일자 입력 (VBA CommandButton4: input.value + blur) ──
            string dateStr = (rec.sampleDate ?? rec.date ?? "").Trim();
            if (!string.IsNullOrEmpty(dateStr))
            {
                string dtE = dateStr.Replace(".", "-").Replace("/", "-");
                if (dtE.Length >= 10) dtE = dtE[..10];
                await CdpEvalAsync(socket, $@"(function(){{
                    var el = document.getElementById('add_meas_start_dt');
                    if (!el) return 'NO';
                    el.value = '{dtE}';
                    el.blur();
                    return 'OK';
                }})()", cts.Token);
            }

            // ── D2. 계획번호 입력 — meas_no_yn 체크 + add_meas_no / edit_meas_no (둘 다 시도) ──
            {
                string baseNo = (rec.quoteNo ?? rec.accNo ?? "").Trim();
                string measNo = $"{baseNo} {rec.sample ?? ""}".Trim();
                string measNoE = measNo.Replace("\\", "\\\\").Replace("'", "\\'");
                string measNoResult = ExtractCdpValue(await CdpEvalAsync(socket, $@"(function(){{
                    var chk = document.getElementById('meas_no_yn');
                    var chkState = 'NO_CHK';
                    if (chk) {{
                        if (!chk.checked) {{
                            chk.click();  // 토글 + onClickMeas_no_yn() 실행
                        }}
                        chkState = chk.checked ? 'CHECKED' : 'UNCHECKED';
                    }}
                    var inp = document.getElementById('add_meas_no') || document.getElementById('edit_meas_no');
                    if (inp) {{
                        inp.removeAttribute('disabled');
                        inp.removeAttribute('readonly');
                        inp.value = '{measNoE}';
                        inp.dispatchEvent(new Event('input', {{bubbles:true}}));
                        inp.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return chkState + '|INP:' + inp.id + '=' + inp.value;
                    }}
                    return chkState + '|NO_INP';
                }})()", cts.Token));
                Log($"  계획번호 입력 결과: {measNoResult} (measNo='{measNo}')");
            }

            // ── H. 차량 선택 — edit_meas_car_no / add_meas_car_no → 1000004880 ──
            {
                await CdpEvalAsync(socket, @"(function(){
                    var carSel = document.getElementById('add_meas_car_no')
                               || document.getElementById('edit_meas_car_no');
                    if (!carSel) return 'NO_SEL';
                    var target = '1000004880';
                    var carOpt = carSel.querySelector('option[value=""' + target + '""]');
                    if (carOpt) {
                        carOpt.selected = true;
                        try { carOpt.click(); } catch(e) {}
                        carSel.dispatchEvent(new Event('change', {bubbles:true}));
                    }
                    if (window.$ && carSel.id) {
                        var $car = window.$('#' + carSel.id);
                        if ($car.length && $car.data && $car.data('select2')) {
                            $car.trigger({ type:'select2:select', params:{ data:{ id:target, text:'' } } });
                            $car.trigger('change');
                            try { $car.data('select2').close(); } catch(e) {}
                        }
                    }
                    return carOpt ? 'OK' : 'NO_OPT';
                })()", cts.Token);
            }

            // ── E. 분석항목 선택 — 이 레코드의 자식노드만 (VBA CommandButton6) ──
            if (analyteCodes.Count > 0)
            {
                string codesJs = string.Join(",", analyteCodes.Select(c => $"'{c.Replace("'", "\\\'")}'"));
                string itemScript = $@"(function(){{
                    var sel = document.getElementById('add_meas_item');
                    if (!sel) return 'NO_SEL';
                    Array.from(sel.options).forEach(function(opt){{ opt.selected = false; }});
                    var codes = [{codesJs}];
                    var found = 0;
                    codes.forEach(function(code){{
                        Array.from(sel.options).forEach(function(opt){{
                            if (opt.value === code) {{ opt.selected = true; found++; }}
                        }});
                    }});
                    sel.dispatchEvent(new Event('change'));
                    if (window.$ && window.$('#add_meas_item').data('select2')) {{
                        window.$('#add_meas_item').data('select2').close();
                    }}
                    return 'ITEMS:' + found + '/' + codes.length;
                }})()";
                string itemResult = ExtractCdpValue(await CdpEvalAsync(socket, itemScript, cts.Token));
                Log($"  분석항목 선택 결과: {itemResult}");
            }

            // ── F. 인력 선택 (VBA CommandButton9: 기존 해제 → 다중 option.selected + change) ──
            if (empIds.Count > 0)
            {
                await CdpEvalAsync(socket, @"(function(){
                    var sel = document.getElementById('add_emp_id');
                    if (!sel) return 'NO';
                    while (sel.selectedOptions.length > 0) sel.selectedOptions[0].selected = false;
                    sel.dispatchEvent(new Event('change', {bubbles:true}));
                    return 'CLR';
                })()", cts.Token);

                string empJs = string.Join(",", empIds.Select(id => $"'{id.Replace("'", "\\\'")}'"));
                string empScript = $@"(function(){{
                    var sel = document.getElementById('add_emp_id');
                    if (!sel) return 'NO_SEL';
                    var ids = [{empJs}];
                    var found = 0;
                    ids.forEach(function(id){{
                        var opt = sel.querySelector('option[value=""' + id + '""]');
                        if (opt) {{ opt.selected = true; found++; }}
                    }});
                    sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'EMP:' + found + '/' + ids.length;
                }})()";
                string empResult = ExtractCdpValue(await CdpEvalAsync(socket, empScript, cts.Token));
                Log($"  인력 선택 결과: {empResult}");
            }

            // ── F-2. 측정장비 — 현장측정장비 테이블의 코드값으로 option 매칭 ──
            {
                List<(string 장비명, string 코드값)> fieldEquip;
                try { fieldEquip = MeasurerService.GetFieldMeasEquipment(); }
                catch (Exception ex) { Log($"  ⚠ GetFieldMeasEquipment 실패: {ex.Message}"); fieldEquip = new List<(string, string)>(); }
                Log($"  현장측정장비: {fieldEquip.Count}건 [{string.Join(";", fieldEquip.Select(x => x.장비명))}]");

                if (fieldEquip.Count > 0)
                {
                    string codesJs = string.Join(",", fieldEquip
                        .Select(x => x.코드값)
                        .Where(c => !string.IsNullOrWhiteSpace(c))
                        .Select(c => $"'{c.Trim().Replace("'", "\\'")}'"));
                    string equipResult = ExtractCdpValue(await CdpEvalAsync(socket, $@"(function(){{
                        var ids = ['edit_meas_equip_no','add_meas_equip_no'];
                        var sel = null, usedId = null;
                        for (var k=0;k<ids.length;k++){{
                            var s = document.getElementById(ids[k]);
                            if (s) {{ sel = s; usedId = ids[k]; break; }}
                        }}
                        if (!sel) return 'NO_SEL';
                        while (sel.selectedOptions.length > 0) sel.selectedOptions[0].selected = false;
                        var codes = [{codesJs}];
                        var picked = 0;
                        var miss = [];
                        codes.forEach(function(c){{
                            var opt = sel.querySelector('option[value=""' + c + '""]');
                            if (opt) {{ opt.selected = true; picked++; }}
                            else miss.push(c);
                        }});
                        sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                        if (window.$ && window.$(sel).data('select2')) {{
                            window.$(sel).trigger('change.select2');
                        }}
                        return 'EQUIP('+usedId+'):' + picked + '/' + codes.length + (miss.length ? ' MISS=['+miss.join('|')+']' : '');
                    }})()", cts.Token));
                    Log($"  측정장비 선택 결과: {equipResult}");
                }
                else
                {
                    Log("  측정장비 스킵 (현장측정장비 0건 — 권한관리→분석장비 탭에서 '🎯 현장장비 스크랩' 필요)");
                }
            }

            // ── G. 측정목적 — 레코드별 다이얼로그 선택값 사용 ──
            {
                string pv = (purposeValues != null && ri < purposeValues.Count
                    ? purposeValues[ri] : "CF").Replace("'", "\\'");
                await CdpEvalAsync(socket, $@"(function(){{
                    var sel = document.getElementById('add_meas_purpose');
                    if (!sel) return 'NO';
                    var opt = sel.querySelector('option[value=""{pv}""]');
                    if (!opt) opt = sel.querySelector('option[value=""CF""]');
                    if (opt) {{ opt.selected = true; }}
                    sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'OK:' + (opt ? opt.value : 'none');
                }})()", cts.Token);
            }


            // ── 저장 버튼 클릭 → 모달 닫힘 대기 (마지막 레코드 포함 전체) ──
            string saveResult = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
                var btn = document.getElementById('insertFieldPlanBtn');
                if (btn) { btn.click(); return 'SAVE'; }
                return 'NO_BTN';
            })()", cts.Token));
            Log($"  saveResult='{saveResult}'");

            // 모달 닫힘 대기 (add_meas_cont_no 사라질 때까지 최대 5초)
            for (int w = 0; w < 5000; w += 400)
            {
                await Task.Delay(400, cts.Token);
                string chk = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
                    var el = document.getElementById('add_meas_cont_no');
                    return (el && el.offsetParent !== null) ? 'OPEN' : 'CLOSED';
                })()", cts.Token));
                if (chk == "CLOSED") break;
            }
            Log($"  ✅ record[{ri}] 완료");

            // 다음 루프 또는 완료 전 잠깐 대기
            await Task.Delay(800, cts.Token);
        } // end for each record

        await socket.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None);
        return true;
    }

    private static async Task<string> CdpEvalAsync(ClientWebSocket ws, string expression, CancellationToken ct)
    {
        string msg = System.Text.Json.JsonSerializer.Serialize(new
        {
            id = Interlocked.Increment(ref _cdpEvalId),
            method = "Runtime.evaluate",
            @params = new { expression, returnByValue = true }
        });
        var buf = Encoding.UTF8.GetBytes(msg);
        await ws.SendAsync(buf, WebSocketMessageType.Text, true, ct);
        var recv = new byte[65536];
        var result = await ws.ReceiveAsync(recv, ct);
        return Encoding.UTF8.GetString(recv, 0, result.Count);
    }
    private static int _cdpEvalId;

    private static string ExtractCdpValue(string cdpResponse)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(cdpResponse);
            return doc.RootElement.GetProperty("result").GetProperty("result").GetProperty("value").GetString() ?? "";
        }
        catch { return ""; }
    }

    private static HashSet<string> FixedColsForAnalyte() =>
        new(StringComparer.OrdinalIgnoreCase)
        {            // 내부 키 컬럼
            "_id","rowid",            "약칭","시료명","접수번호","의뢰일","업체명","대표자",
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
            CanResize = false, Background = AppRes("PanelBg"),
            SystemDecorations = Avalonia.Controls.SystemDecorations.Full,
        };
        var btnOk = new Button
        {
            Content = "확인", Width = 80, Height = 28, FontSize = AppTheme.FontBase, FontFamily = Font,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
            BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        btnOk.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = message, FontSize = AppTheme.FontBase, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
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
