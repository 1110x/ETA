using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 폐수 업체명 정리 — 폐수배출업소 공식명 ↔ DATA 테이블 명칭 매핑 후 일괄 변경
/// </summary>
public sealed class WasteNameReconcilePage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 3개 패널 (MainPage → Show1 / Show2 / Show4) ──────────────────────────
    public Control LeftPanel   { get; }
    public Control CenterPanel { get; }
    public Control RightPanel  { get; }

    // ── 상태 ─────────────────────────────────────────────────────────────────
    private string?      _officialName;
    private string?      _officialKey;    // 관리번호 — 동명 업체 구분용
    private readonly List<string> _aliases = new();

    // ── UI 참조 ───────────────────────────────────────────────────────────────
    private string   _activeGroup = "여수";
    private TreeView? _officialTree;
    private Button?   _btnYeosoo, _btnYulchon, _btnSepung;
    private TextBox?  _searchBox;
    private List<WasteCompany> _officialItems = new();
    private TextBlock? _officialLabel;
    private StackPanel? _aliasList;
    private TextBlock?  _statusLabel;
    private TreeView?   _dataTree;

    public WasteNameReconcilePage()
    {
        LeftPanel   = BuildOfficialPanel();
        CenterPanel = BuildMappingPanel();
        RightPanel  = BuildDataNamePanel();
    }

    public void Reload()
    {
        LoadOfficialTree();
        RefreshDataTree();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // 좌측: 폐수배출업소 공식명 트리
    // ═══════════════════════════════════════════════════════════════════════════
    private Control BuildOfficialPanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };

        // 탭 헤더
        var tabBar = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        tabBar.Children.Add(new TextBlock
        {
            Text = "공식 업체명", FontSize = 11, FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#8899bb"), FontFamily = Font,
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0),
        });
        _btnYeosoo  = MakeTabBtn("🌊 여수", "여수", true);
        _btnYulchon = MakeTabBtn("🏗 율촌", "율촌", false);
        _btnSepung  = MakeTabBtn("🏭 세풍", "세풍", false);
        tabBar.Children.Add(_btnYeosoo);
        tabBar.Children.Add(_btnYulchon);
        tabBar.Children.Add(_btnSepung);

        var header = new Border
        {
            Background = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(8, 4),
            Child = tabBar,
        };

        // 검색창
        _searchBox = new TextBox
        {
            Watermark = "🔍 업체명 검색",
            FontSize = 11, FontFamily = Font,
            Margin = new Thickness(5, 4, 5, 0),
            Padding = new Thickness(6, 4),
            CornerRadius = new CornerRadius(4),
        };
        _searchBox.TextChanged += (_, _) => ApplyOfficialFilter(_searchBox.Text ?? "");
        Grid.SetRow(_searchBox, 1);

        _officialTree = new TreeView { Margin = new Thickness(5) };
        _officialTree.SelectionChanged += OfficialTree_SelectionChanged;
        Grid.SetRow(_officialTree, 2);

        root.Children.Add(header);
        root.Children.Add(_searchBox);
        root.Children.Add(_officialTree);

        LoadOfficialTree();
        return root;
    }

    private Button MakeTabBtn(string content, string tag, bool active)
    {
        var btn = new Button
        {
            Content = content, Tag = tag,
            FontSize = 10, FontFamily = Font,
            Padding = new Thickness(8, 3), CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
        };
        ApplyTabStyle(btn, active);
        btn.Click += (_, _) =>
        {
            _activeGroup = tag;
            ApplyTabStyle(_btnYeosoo!,  _activeGroup == "여수");
            ApplyTabStyle(_btnYulchon!, _activeGroup == "율촌");
            ApplyTabStyle(_btnSepung!,  _activeGroup == "세풍");
            LoadOfficialTree();
        };
        return btn;
    }

    private static void ApplyTabStyle(Button btn, bool active)
    {
        var (bg, fg, border) = (btn.Tag as string, active) switch
        {
            ("여수", true) => ("#1e3a5a", "#88aacc", "#336699"),
            ("율촌", true) => ("#1a3a1a", "#aaccaa", "#336633"),
            ("세풍", true) => ("#3a2a1a", "#ccaa88", "#996633"),
            _              => ("#222222", "#666666", "#444444"),
        };
        btn.Background  = Brush.Parse(bg);
        btn.Foreground  = Brush.Parse(fg);
        btn.BorderBrush = Brush.Parse(border);
    }

    private void LoadOfficialTree()
    {
        try
        {
            _officialItems = WasteCompanyService.GetAllItems()
                .Where(c => GetGroup(c) == _activeGroup)
                .OrderBy(c => c.업체명, StringComparer.CurrentCulture)
                .ToList();
        }
        catch { _officialItems = new(); }
        ApplyOfficialFilter(_searchBox?.Text ?? "");
    }

    private void ApplyOfficialFilter(string keyword)
    {
        if (_officialTree == null) return;
        _officialTree.Items.Clear();
        var filtered = string.IsNullOrWhiteSpace(keyword)
            ? _officialItems
            : _officialItems.Where(c =>
                c.업체명.Contains(keyword, StringComparison.CurrentCultureIgnoreCase) ||
                c.관리번호.Contains(keyword, StringComparison.CurrentCultureIgnoreCase)).ToList();
        foreach (var c in filtered)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            sp.Children.Add(new TextBlock
            {
                Text = c.업체명, FontSize = 12, FontFamily = Font,
                Foreground = AppRes("AppFg"),
                VerticalAlignment = VerticalAlignment.Center,
            });
            sp.Children.Add(new TextBlock
            {
                Text = c.관리번호, FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse("#666666"),
                VerticalAlignment = VerticalAlignment.Center,
            });
            if (!string.IsNullOrEmpty(c.비용부담금_업체명))
                sp.Children.Add(new TextBlock
                {
                    Text = $"[{c.비용부담금_업체명}]", FontSize = 9, FontFamily = Font,
                    Foreground = Brush.Parse("#5599ee"),
                    VerticalAlignment = VerticalAlignment.Center,
                });
            _officialTree.Items.Add(new TreeViewItem { Tag = c, Header = sp });
        }
    }

    private void OfficialTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        WasteCompany? c = null;
        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is WasteCompany wc) c = wc;
            else if (e.AddedItems[0] is WasteCompany wc2) c = wc2;
        }
        if (c == null) return;

        _officialName = c.업체명;
        _officialKey  = c.관리번호;
        _aliases.Clear();
        if (_statusLabel != null) _statusLabel.Text = "";
        RefreshMappingPanel();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // 중앙: 매핑 편집기
    // ═══════════════════════════════════════════════════════════════════════════
    private Control BuildMappingPanel()
    {
        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        outer.Children.Add(new Border
        {
            Background = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text = "🔗  명칭 매핑",
                FontSize = 11, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
            }
        });

        var body = new StackPanel { Spacing = 8, Margin = new Thickness(8, 10) };

        // 공식명 영역
        body.Children.Add(new TextBlock
        {
            Text = "📌  공식명  (A컬럼 — 변경 목표)",
            FontSize = 10, FontFamily = Font, Foreground = Brush.Parse("#7788aa"),
        });
        _officialLabel = new TextBlock
        {
            Text = "← 왼쪽 트리에서 공식명 선택",
            FontSize = 14, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#ccddff"),
            Margin = new Thickness(4, 2, 0, 8),
        };
        body.Children.Add(_officialLabel);

        body.Children.Add(new Border { Height = 1, Background = Brush.Parse("#333344") });

        // DATA 명칭 목록 영역
        body.Children.Add(new TextBlock
        {
            Text = "🔀  변경 대상 DATA 명칭  (오른쪽 트리에서 클릭 → 추가)",
            FontSize = 10, FontFamily = Font, Foreground = Brush.Parse("#7788aa"),
            Margin = new Thickness(0, 4, 0, 0),
        });
        _aliasList = new StackPanel { Spacing = 3, Margin = new Thickness(0, 4, 0, 0) };
        body.Children.Add(_aliasList);

        // 실행 버튼
        var execBtn = new Button
        {
            Content = "▶  DATA 일괄 변경 실행",
            FontSize = 11, FontFamily = Font,
            Background = Brush.Parse("#1a3a2a"),
            Foreground = Brush.Parse("#66ee88"),
            BorderBrush = Brush.Parse("#336644"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 6),
            Margin = new Thickness(0, 12, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        execBtn.Click += ExecuteRename;
        body.Children.Add(execBtn);

        _statusLabel = new TextBlock
        {
            FontSize = 10, FontFamily = Font,
            Foreground = Brush.Parse("#666666"),
            Margin = new Thickness(0, 4, 0, 0),
        };
        body.Children.Add(_statusLabel);

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = body,
        };
        Grid.SetRow(scroll, 1);
        outer.Children.Add(scroll);

        RefreshMappingPanel();
        return outer;
    }

    private void RefreshMappingPanel()
    {
        if (_officialLabel == null || _aliasList == null) return;

        _officialLabel.Text = string.IsNullOrEmpty(_officialName)
            ? "← 왼쪽 트리에서 공식명 선택"
            : _officialName;

        _aliasList.Children.Clear();

        if (_aliases.Count == 0)
        {
            _aliasList.Children.Add(new TextBlock
            {
                Text = "오른쪽 트리에서 DATA 명칭을 클릭하여 추가하세요",
                FontSize = 10, FontFamily = Font,
                Foreground = Brush.Parse("#555555"),
            });
            return;
        }

        foreach (var alias in _aliases.ToList())
        {
            var captured = alias;
            var rowGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            var label = new TextBlock
            {
                Text = captured, FontSize = 11, FontFamily = Font,
                Foreground = Brush.Parse("#aaffaa"),
                VerticalAlignment = VerticalAlignment.Center,
            };
            var removeBtn = new Button
            {
                Content = "✕", FontSize = 9,
                Background = Brush.Parse("#3a1a1a"),
                Foreground = Brush.Parse("#ff8888"),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(6, 2), CornerRadius = new CornerRadius(3),
            };
            removeBtn.Click += (_, _) => { _aliases.Remove(captured); RefreshMappingPanel(); };
            Grid.SetColumn(removeBtn, 1);
            rowGrid.Children.Add(label);
            rowGrid.Children.Add(removeBtn);

            _aliasList.Children.Add(new Border
            {
                Background = Brush.Parse("#1e2a1e"),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8, 4),
                Margin = new Thickness(0, 1),
                Child = rowGrid,
            });
        }
    }

    private void ExecuteRename(object? sender, RoutedEventArgs e)
    {
        if (_statusLabel == null) return;
        if (string.IsNullOrEmpty(_officialName))
        {
            _statusLabel.Foreground = Brush.Parse("#ee8844");
            _statusLabel.Text = "공식명을 먼저 선택하세요.";
            return;
        }
        if (_aliases.Count == 0)
        {
            _statusLabel.Foreground = Brush.Parse("#ee8844");
            _statusLabel.Text = "변경할 DATA 명칭을 추가하세요.";
            return;
        }

        int total = 0;
        var errors = new List<string>();
        foreach (var alias in _aliases)
        {
            try { total += WasteDataService.RenameCompany(alias, _officialName); }
            catch (Exception ex) { errors.Add($"[{alias}] {ex.Message}"); }
        }

        if (errors.Count > 0)
        {
            var logPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "..", "..", "..", "Logs", "WasteRenameError.log");
            File.AppendAllText(logPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {_officialName}\n" +
                string.Join("\n", errors) + "\n\n");
        }

        _statusLabel.Foreground = errors.Count == 0
            ? Brush.Parse("#88ee88") : Brush.Parse("#ee4444");
        _statusLabel.Text = errors.Count == 0
            ? $"✅  {total}건 변경 완료"
            : $"⚠  {total}건 변경, 오류 {errors.Count}건: {errors[0]}";

        _aliases.Clear();
        RefreshMappingPanel();
        Reload();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // 우측: DATA 테이블 업체명 트리
    // ═══════════════════════════════════════════════════════════════════════════
    private TextBlock? _dataStatusLabel;

    private Control BuildDataNamePanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };

        root.Children.Add(new Border
        {
            Background = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text = "📊  DATA 테이블 업체명",
                FontSize = 11, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = Brush.Parse("#8899bb"),
            }
        });

        _dataStatusLabel = new TextBlock
        {
            FontSize = 10, FontFamily = Font,
            Foreground = Brush.Parse("#666666"),
            Margin = new Thickness(8, 3),
        };
        Grid.SetRow(_dataStatusLabel, 1);
        root.Children.Add(_dataStatusLabel);

        _dataTree = new TreeView { Margin = new Thickness(5) };
        _dataTree.SelectionChanged += DataTree_SelectionChanged;
        Grid.SetRow(_dataTree, 2);
        root.Children.Add(_dataTree);

        RefreshDataTree();
        return root;
    }

    private void RefreshDataTree()
    {
        if (_dataTree == null) return;
        _dataTree.Items.Clear();
        try
        {
            var dataNames = WasteDataService.GetDataCompanyNames();

            // 공식명 전체 목록 (그룹 무관)
            var officialNames = new HashSet<string>(
                WasteCompanyService.GetAllItems().Select(c => c.업체명),
                StringComparer.Ordinal);

            // 실무명칭(비용부담금_업체명)으로 이미 저장된 명칭 목록
            var costNames = WasteCompanyService.GetAllCostNames();

            // 공식명에 없는 것만 표시 (미매칭)
            var unmatched = dataNames.Where(n => !officialNames.Contains(n))
                                     .OrderBy(n => n, StringComparer.CurrentCulture)
                                     .ToList();

            if (_dataStatusLabel != null)
            {
                _dataStatusLabel.Foreground = Brush.Parse("#666666");
                _dataStatusLabel.Text = $"미매칭 {unmatched.Count}개 / 전체 {dataNames.Count}개";
            }
            foreach (var name in unmatched)
            {
                bool isMapped = costNames.Contains(name);
                var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                sp.Children.Add(new TextBlock
                {
                    Text = name, FontSize = 11, FontFamily = Font,
                    Foreground = isMapped ? Brush.Parse("#5599ee") : AppRes("AppFg"),
                });
                if (isMapped)
                    sp.Children.Add(new TextBlock
                    {
                        Text = "✓", FontSize = 10, FontFamily = Font,
                        Foreground = Brush.Parse("#5599ee"),
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                _dataTree.Items.Add(new TreeViewItem { Tag = name, Header = sp });
            }
        }
        catch (Exception ex)
        {
            if (_dataStatusLabel != null)
            {
                _dataStatusLabel.Foreground = Brush.Parse("#ee4444");
                _dataStatusLabel.Text = $"오류: {ex.Message}";
            }
        }
    }

    private void DataTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        string? name = null;
        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is string s) name = s;
            else if (e.AddedItems[0] is string s2) name = s2;
        }
        if (string.IsNullOrEmpty(name)) return;

        if (string.IsNullOrEmpty(_officialName))
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = Brush.Parse("#ee8844");
                _statusLabel.Text = "먼저 왼쪽에서 공식명을 선택하세요.";
            }
            return;
        }

        if (!_aliases.Contains(name))
        {
            _aliases.Add(name);
            RefreshMappingPanel();
        }
    }

    // ── 헬퍼 ──────────────────────────────────────────────────────────────────
    private static string GetGroup(WasteCompany c)
    {
        var p = c.프로젝트명 ?? "";
        if (p.Contains("여수")) return "여수";
        if (p.Contains("세풍")) return "세풍";
        if (p.Contains("율촌")) return "율촌";
        return "기타";
    }

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
}
