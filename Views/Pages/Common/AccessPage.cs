using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services.SERVICE1;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 접근관리 페이지
///   Show1 — 직원 트리 (AgentTreePage와 동일 구조)
///   Show2 — 선택된 직원의 메뉴별 접근 토글 목록
/// </summary>
public class AccessPage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly string[] DeptItems = { "전체", "수질분석센터", "처리시설", "일반업무", "기타" };

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public Control Show1 { get; }
    public Control Show2 { get; }

    // ── 내부 상태 ────────────────────────────────────────────────────────────
    private Agent?    _selected;
    private string    _deptFilter   = "전체";
    private string    _searchText   = "";
    private readonly TreeView         _treeView;
    private readonly ContentControl   _show2Content;
    private DispatcherTimer?          _debounce;

    public AccessPage()
    {
        AccessService.EnsureTable();

        _treeView    = BuildTreeView();
        _show2Content = new ContentControl
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment   = VerticalAlignment.Stretch,
        };

        Show1 = BuildShow1Panel();
        Show2 = new Border
        {
            Padding  = new Thickness(6),
            Child    = _show2Content,
        };

        LoadTree();
    }

    // =========================================================================
    // Show1 — 직원 트리
    // =========================================================================
    private Control BuildShow1Panel()
    {
        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        // 헤더
        var header = new Border
        {
            Background   = AppRes("PanelBg"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child        = new TextBlock
            {
                Text       = "🔐 접근관리 — 직원 목록",
                FontSize   = AppTheme.FontBase,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted"),
                FontFamily = Font,
            }
        };
        Grid.SetRow(header, 0);

        // 검색
        var searchBox = new TextBox
        {
            Watermark       = "🔍 직원 검색...",
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            Height          = 32,
            Margin          = new Thickness(5, 4, 5, 2),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(8, 5),
            BorderThickness = new Thickness(1),
            BorderBrush     = AppRes("InputBorder"),
            Background      = AppRes("PanelBg"),
        };
        searchBox.TextChanged += (_, _) =>
        {
            _debounce?.Stop();
            _debounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _debounce.Tick += (_, _) =>
            {
                _debounce!.Stop();
                _searchText = searchBox.Text?.Trim().ToLower() ?? "";
                ApplyFilters();
            };
            _debounce.Start();
        };
        Grid.SetRow(searchBox, 1);

        // 부서 필터
        var deptBox = new ComboBox
        {
            ItemsSource         = DeptItems,
            SelectedIndex       = 0,
            FontSize            = AppTheme.FontMD,
            FontFamily          = Font,
            Height              = 30,
            Margin              = new Thickness(5, 0, 5, 3),
            Padding             = new Thickness(8, 4),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            BorderThickness     = new Thickness(1),
            BorderBrush         = AppRes("InputBorder"),
            Background          = AppRes("PanelBg"),
        };
        deptBox.SelectionChanged += (_, _) =>
        {
            _deptFilter = deptBox.SelectedItem?.ToString() ?? "전체";
            ApplyFilters();
        };
        Grid.SetRow(deptBox, 2);

        Grid.SetRow(_treeView, 3);

        grid.Children.Add(header);
        grid.Children.Add(searchBox);
        grid.Children.Add(deptBox);
        grid.Children.Add(_treeView);

        return grid;
    }

    private TreeView BuildTreeView()
    {
        var tv = new TreeView { Margin = new Thickness(5) };
        tv.SelectionChanged += TreeView_SelectionChanged;
        return tv;
    }

    private void LoadTree()
    {
        _treeView.Items.Clear();
        _selected = null;
        _show2Content.Content = BuildEmptyHint();

        try
        {
            var agents = AgentService.GetAllItems().OrderBy(a => a.입사일).ToList();
            foreach (var a in agents)
                _treeView.Items.Add(CreateTreeItem(a));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AccessPage] LoadTree 실패: {ex.Message}");
        }
    }

    private TreeViewItem CreateTreeItem(Agent agent)
    {
        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 8,
            VerticalAlignment = VerticalAlignment.Center,
        };

        row.Children.Add(MakePhotoCircle(agent, 26));

        // 직급 배지
        if (!string.IsNullOrWhiteSpace(agent.직급))
        {
            var (bbg, bfg) = BadgeColorHelper.GetBadgeColor(agent.직급);
            row.Children.Add(new Border
            {
                Background    = new SolidColorBrush(Color.Parse(bbg)),
                CornerRadius  = new CornerRadius(3),
                Padding       = new Thickness(5, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Child         = new TextBlock
                {
                    Text      = agent.직급,
                    FontSize  = AppTheme.FontSM,
                    Foreground = new SolidColorBrush(Color.Parse(bfg)),
                    FontFamily = Font,
                }
            });
        }

        row.Children.Add(new TextBlock
        {
            Text      = agent.성명,
            FontSize  = AppTheme.FontLG,
            FontFamily = Font,
            Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        // 부서 표시
        if (!string.IsNullOrWhiteSpace(agent.부서))
            row.Children.Add(new TextBlock
            {
                Text      = $"({agent.부서})",
                FontSize  = AppTheme.FontSM,
                FontFamily = Font,
                Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
            });

        return new TreeViewItem { Header = row, Tag = agent };
    }

    private void ApplyFilters()
    {
        foreach (var tvi in _treeView.Items.OfType<TreeViewItem>())
        {
            if (tvi.Tag is not Agent a) continue;
            bool nameOk = string.IsNullOrEmpty(_searchText)
                || (a.성명?.ToLower().Contains(_searchText) ?? false)
                || (a.사번?.ToLower().Contains(_searchText) ?? false);
            bool deptOk = _deptFilter == "전체" || a.부서 == _deptFilter;
            tvi.IsVisible = nameOk && deptOk;
        }
    }

    private void TreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        Agent? agent = null;
        if (e.AddedItems.Count > 0 && e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is Agent a)
            agent = a;
        if (agent == null) return;
        _selected = agent;
        _show2Content.Content = BuildAccessPanel(agent);
    }

    // =========================================================================
    // Show2 — 접근 권한 토글 패널
    // =========================================================================
    private Control BuildAccessPanel(Agent agent)
    {
        var accessMap = AccessService.GetAccessMap(agent.사번);

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(4) };

        // 헤더
        root.Children.Add(new TextBlock
        {
            Text       = $"🔐 {agent.성명} ({agent.사번}) — 메뉴 접근 권한",
            FontSize   = AppTheme.FontXL,
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font,
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 6),
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderDefault,
            Margin     = new Thickness(0, 0, 0, 8),
        });

        // 안내
        root.Children.Add(new TextBlock
        {
            Text       = "비활성화된 메뉴는 해당 직원에게 숨김 처리됩니다.",
            FontSize   = AppTheme.FontBase,
            FontFamily = Font,
            Foreground = AppRes("FgMuted"),
            Margin     = new Thickness(0, 0, 0, 12),
        });

        // 카테고리별 그룹
        var categories = AccessService.AllMenus
            .GroupBy(m => m.Category)
            .ToList();

        foreach (var cat in categories)
        {
            // 카테고리 헤더
            root.Children.Add(new Border
            {
                Background   = AppTheme.BgSecondary,
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(8, 4),
                Margin       = new Thickness(0, 4, 0, 2),
                Child        = new TextBlock
                {
                    Text       = cat.Key,
                    FontSize   = AppTheme.FontMD,
                    FontWeight = FontWeight.SemiBold,
                    FontFamily = Font,
                    Foreground = AppTheme.FgInfo,
                }
            });

            foreach (var (key, label, _) in cat)
            {
                bool allowed = accessMap.TryGetValue(key, out var v) ? v : true;
                root.Children.Add(BuildToggleRow(agent.사번, key, label, allowed));
            }
        }

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Content = root,
        };
        return scroll;
    }

    private Control BuildToggleRow(string 사번, string key, string label, bool allowed)
    {
        var row = new Grid
        {
            Margin = new Thickness(0, 1, 0, 1),
        };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var lbl = new TextBlock
        {
            Text              = label,
            FontSize          = AppTheme.FontLG,
            FontFamily        = Font,
            Foreground        = allowed ? AppRes("AppFg") : AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            Padding           = new Thickness(8, 5),
        };
        Grid.SetColumn(lbl, 0);

        // 커스텀 토글 버튼 (ToggleSwitch 대신 Border+애니메이션 불필요 → 간단 버튼)
        var toggle = BuildToggleSwitch(allowed, isOn =>
        {
            AccessService.SetAccess(사번, key, isOn);
            lbl.Foreground = isOn ? AppRes("AppFg") : AppRes("FgMuted");
        });
        Grid.SetColumn(toggle, 1);

        row.Children.Add(lbl);
        row.Children.Add(toggle);

        // 구분선
        var wrap = new StackPanel { Spacing = 0 };
        wrap.Children.Add(row);
        wrap.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderSubtle,
            Margin     = new Thickness(8, 0),
        });
        return wrap;
    }

    private static Control BuildToggleSwitch(bool initialOn, Action<bool> onChange)
    {
        bool isOn = initialOn;

        var knob = new Border
        {
            Width               = 18,
            Height              = 18,
            CornerRadius        = new CornerRadius(9),
            Background          = Brushes.White,
            Margin              = isOn ? new Thickness(22, 0, 0, 0) : new Thickness(2, 0, 0, 0),
            VerticalAlignment   = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Left,
        };

        var track = new Border
        {
            Width           = 44,
            Height          = 24,
            CornerRadius    = new CornerRadius(12),
            Background      = isOn
                ? new SolidColorBrush(Color.Parse("#2e7d32"))
                : new SolidColorBrush(Color.Parse("#555566")),
            Padding         = new Thickness(0),
            Margin          = new Thickness(8, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Cursor          = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            Child           = knob,
        };

        track.PointerPressed += (_, _) =>
        {
            isOn = !isOn;
            track.Background = isOn
                ? new SolidColorBrush(Color.Parse("#2e7d32"))
                : new SolidColorBrush(Color.Parse("#555566"));
            knob.Margin = isOn ? new Thickness(22, 0, 0, 0) : new Thickness(2, 0, 0, 0);
            onChange(isOn);
        };

        return track;
    }

    private static Control MakePhotoCircle(Agent agent, double size)
    {
        if (!string.IsNullOrEmpty(agent.PhotoPath))
        {
            AgentService.EnsurePhotoLocal(agent.사번, agent.PhotoPath);
            var fullPath = System.IO.Path.IsPathRooted(agent.PhotoPath)
                ? agent.PhotoPath
                : System.IO.Path.Combine(AgentService.GetPhotoDirectory(), agent.PhotoPath);
            if (System.IO.File.Exists(fullPath))
            {
                try
                {
                    using var stream = System.IO.File.OpenRead(fullPath);
                    var bmp = new Bitmap(stream);
                    return new Ellipse
                    {
                        Width  = size,
                        Height = size,
                        Fill   = new ImageBrush(bmp) { Stretch = Stretch.UniformToFill },
                    };
                }
                catch { }
            }
        }

        // fallback: 이니셜 원
        var initial = agent.성명.Length > 0 ? agent.성명[0].ToString() : "?";
        var key     = !string.IsNullOrWhiteSpace(agent.직급) ? agent.직급 : agent.성명;
        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(key);
        return new Border
        {
            Width        = size,
            Height       = size,
            CornerRadius = new CornerRadius(size / 2),
            Background   = new SolidColorBrush(Color.Parse(ibg)),
            ClipToBounds = true,
            Child        = new TextBlock
            {
                Text                = initial,
                FontSize            = size * 0.42,
                Foreground          = new SolidColorBrush(Color.Parse(ifg)),
                FontFamily          = Font,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
    }

    private static Control BuildEmptyHint() =>
        new TextBlock
        {
            Text              = "← 직원을 선택하면\n메뉴 접근 권한을 설정할 수 있습니다.",
            FontSize          = AppTheme.FontLG,
            FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground        = new SolidColorBrush(Color.Parse("#666677")),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment   = VerticalAlignment.Center,
            TextAlignment       = Avalonia.Media.TextAlignment.Center,
            Margin              = new Thickness(20),
        };
}
