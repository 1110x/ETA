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
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// 수질분석센터 시료명 정리 — 약칭 선택 → 시료명 별칭 일괄 변경
/// </summary>
public sealed class WaterQualityNameReconcilePage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    public Control LeftPanel   { get; }   // Show1: 약칭 트리
    public Control CenterPanel { get; }   // Show2: 매핑 편집기
    public Control BottomPanel { get; }   // Show3: 의뢰일 그리드
    public Control RightPanel  { get; }   // Show4: 목표명 입력 + 시료명 목록

    private string? _abbr;               // 선택된 약칭
    private readonly List<string> _aliases = new();

    private TreeView?   _abbrTree;
    private TextBox?    _targetBox;      // Show4 목표 시료명 직접 입력
    private TreeView?   _sampleTree;
    private TextBlock?  _sampleStatus;
    private TextBlock?  _officialLabel;  // Show2 - 선택된 약칭 표시
    private StackPanel? _aliasList;
    private TextBlock?  _statusLabel;
    private StackPanel? _bottomContent;   // Show3 컨텐츠 (의뢰일 그리드)

    public WaterQualityNameReconcilePage()
    {
        ETA.Services.Common.SampleDischargeStandardService.EnsureTable();
        LeftPanel   = BuildAbbrPanel();
        CenterPanel = BuildMappingPanel();
        BottomPanel = BuildBottomPanel();
        RightPanel  = BuildSamplePanel();
    }

    public void Reload()
    {
        LoadAbbrTree();
        RefreshSampleTree();
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show1: 약칭 트리
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildAbbrPanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "🏢  약칭 목록",
                FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppTheme.FgMuted,
            }
        });

        _abbrTree = new TreeView { Margin = new Thickness(5) };
        _abbrTree.SelectionChanged += AbbrTree_SelectionChanged;
        Grid.SetRow(_abbrTree, 1);
        root.Children.Add(_abbrTree);

        LoadAbbrTree();
        return root;
    }

    private void LoadAbbrTree()
    {
        if (_abbrTree == null) return;
        _abbrTree.Items.Clear();
        try
        {
            var companyMap = AnalysisRequestService.GetAbbrCompanyMap();
            foreach (var abbr in AnalysisRequestService.GetDistinctAbbreviations())
            {
                var (bgColor, fgColor) = BadgeColorHelper.GetBadgeColor(abbr);
                var badge = new Border
                {
                    Background = new SolidColorBrush(Color.Parse(bgColor)),
                    BorderBrush = new SolidColorBrush(Color.Parse(fgColor)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(10),
                    Padding = new Thickness(6, 1, 8, 1),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = abbr,
                        FontSize = AppTheme.FontSM,
                        FontWeight = FontWeight.Medium,
                        FontFamily = Font,
                        Foreground = new SolidColorBrush(Color.Parse(fgColor)),
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                };

                var company = companyMap.TryGetValue(abbr, out var c) ? c : "";
                var headerPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 6,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                headerPanel.Children.Add(badge);
                headerPanel.Children.Add(new TextBlock
                {
                    Text = company,
                    FontSize = AppTheme.FontMD, FontFamily = Font,
                    FontWeight = FontWeight.Regular,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                });

                _abbrTree.Items.Add(new TreeViewItem
                {
                    Tag    = abbr,
                    Header = headerPanel,
                });
            }
        }
        catch { }
    }

    private void AbbrTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        string? abbr = null;
        if (e.AddedItems.Count > 0 && e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is string s)
            abbr = s;
        if (string.IsNullOrEmpty(abbr)) return;

        _abbr = abbr;
        _aliases.Clear();
        if (_statusLabel  != null) _statusLabel.Text  = "";
        if (_officialLabel != null) _officialLabel.Text = $"약칭: {_abbr}";
        RefreshMappingPanel();
        RefreshSampleTree();
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show2: 매핑 편집기
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildMappingPanel()
    {
        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        outer.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "🔗  시료명 매핑",
                FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppTheme.FgMuted,
            }
        });

        var body = new StackPanel { Spacing = 8, Margin = new Thickness(8, 10) };

        body.Children.Add(new TextBlock
        {
            Text       = "📌  선택된 약칭",
            FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = AppTheme.FgMuted,
        });
        _officialLabel = new TextBlock
        {
            Text       = "← 왼쪽 트리에서 약칭 선택",
            FontSize   = AppTheme.FontXL, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgSecondary,
            Margin     = new Thickness(4, 2, 0, 4),
        };
        body.Children.Add(_officialLabel);

        body.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle });

        // 목표 시료명 텍스트박스 (앞으로 사용될 이름)
        body.Children.Add(new TextBlock
        {
            Text       = "✏  변경될 시료명 (직접 입력)",
            FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = AppTheme.FgMuted,
            Margin     = new Thickness(0, 8, 0, 2),
        });
        _targetBox = new TextBox
        {
            FontSize        = AppTheme.FontMD, FontFamily = Font,
            Watermark       = "앞으로 사용할 시료명... (오른쪽에서 드래그 가능)",
            Background      = Brush.Parse("#1e2030"),
            Foreground      = AppTheme.FgPrimary,
            BorderBrush     = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 5),
            Margin          = new Thickness(0, 0, 0, 4),
        };
        // 드롭 대상 설정
        Avalonia.Input.DragDrop.SetAllowDrop(_targetBox, true);
        _targetBox.AddHandler(Avalonia.Input.DragDrop.DragOverEvent, (_, e) =>
        {
            e.DragEffects = e.Data.Contains(Avalonia.Input.DataFormats.Text)
                ? Avalonia.Input.DragDropEffects.Copy
                : Avalonia.Input.DragDropEffects.None;
        });
        _targetBox.AddHandler(Avalonia.Input.DragDrop.DropEvent, (_, e) =>
        {
            var txt = e.Data.Get(Avalonia.Input.DataFormats.Text) as string;
            if (!string.IsNullOrEmpty(txt))
                _targetBox!.Text = txt;
        });
        body.Children.Add(_targetBox);

        body.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle });

        body.Children.Add(new TextBlock
        {
            Text       = "🔀  변경 대상 시료명  (오른쪽 목록에서 클릭 → 추가)",
            FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = AppTheme.FgMuted,
            Margin     = new Thickness(0, 8, 0, 0),
        });
        _aliasList = new StackPanel { Spacing = 3, Margin = new Thickness(0, 4, 0, 0) };
        body.Children.Add(_aliasList);

        var execBtn = new Button
        {
            Content             = "▶  시료명 일괄 변경 실행",
            FontSize            = AppTheme.FontBase, FontFamily = Font,
            Background          = AppTheme.BgActiveGreen,
            Foreground          = AppTheme.FgSuccess,
            BorderBrush         = AppTheme.BorderActive,
            BorderThickness     = new Thickness(1),
            CornerRadius        = new CornerRadius(4),
            Padding             = new Thickness(12, 6),
            Margin              = new Thickness(0, 12, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        execBtn.Click += ExecuteRename;
        body.Children.Add(execBtn);

        _statusLabel = new TextBlock
        {
            FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = AppTheme.FgDimmed,
            Margin     = new Thickness(0, 4, 0, 0),
            TextWrapping = TextWrapping.Wrap,
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
        if (_aliasList == null) return;
        _aliasList.Children.Clear();

        if (_aliases.Count == 0)
        {
            _aliasList.Children.Add(new TextBlock
            {
                Text       = "오른쪽 목록에서 시료명을 클릭하여 추가하세요",
                FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
                Foreground = AppTheme.FgDimmed,
            });
            return;
        }

        foreach (var alias in _aliases.ToList())
        {
            var captured = alias;
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            var lbl = new TextBlock
            {
                Text              = captured, FontSize = AppTheme.FontBase, FontFamily = Font,
                FontWeight        = FontWeight.Regular,
                Foreground        = AppTheme.FgSuccess,
                VerticalAlignment = VerticalAlignment.Center,
            };
            var removeBtn = new Button
            {
                Content         = "✕", FontSize = AppTheme.FontXS,
                Background      = AppTheme.BgDanger,
                Foreground      = AppTheme.FgDanger,
                BorderThickness = new Thickness(0),
                Padding         = new Thickness(6, 2), CornerRadius = new CornerRadius(3),
            };
            removeBtn.Click += (_, _) => { _aliases.Remove(captured); RefreshMappingPanel(); };
            Grid.SetColumn(removeBtn, 1);
            row.Children.Add(lbl);
            row.Children.Add(removeBtn);
            _aliasList.Children.Add(new Border
            {
                Background   = AppTheme.BgActiveGreen,
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(8, 4),
                Margin       = new Thickness(0, 1),
                Child        = row,
            });
        }

        RefreshBottomPanel();
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show3: 선택된 시료들의 의뢰일 그리드 (4컬럼)
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildBottomPanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "📅  변경 대상 시료 의뢰일",
                FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppTheme.FgMuted,
            }
        });

        _bottomContent = new StackPanel { Spacing = 8, Margin = new Thickness(10, 6) };
        var scroll = new ScrollViewer
        {
            Content = _bottomContent,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);
        return root;
    }

    /// <summary>_aliases 리스트를 기반으로 각 시료명의 의뢰일을 4컬럼 그리드로 표시</summary>
    private void RefreshBottomPanel(HashSet<string>? renamed = null)
    {
        if (_bottomContent == null) return;
        _bottomContent.Children.Clear();

        if (_aliases.Count == 0)
        {
            _bottomContent.Children.Add(new TextBlock
            {
                Text       = "변경 대상 시료명을 선택하세요",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
            });
            return;
        }

        foreach (var name in _aliases)
        {
            var isRenamed = renamed != null && renamed.Contains(name);
            var dates = AnalysisRequestService.GetRequestDatesBySample(_abbr ?? "", name);

            // 시료명 헤더
            var headerText = new TextBlock
            {
                Text = isRenamed ? $"✅ {name}  ({dates.Count}건, 변경됨)" : $"{name}  ({dates.Count}건)",
                FontSize = AppTheme.FontBase, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = isRenamed ? AppTheme.FgSuccess : AppRes("AppFg"),
                Margin = new Thickness(0, 4, 0, 2),
            };
            _bottomContent.Children.Add(headerText);

            if (dates.Count == 0)
            {
                _bottomContent.Children.Add(new TextBlock
                {
                    Text = "  의뢰일 없음",
                    FontSize = AppTheme.FontSM, FontFamily = Font,
                    Foreground = AppTheme.FgDimmed,
                });
                continue;
            }

            // 4컬럼 그리드
            var grid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,*,*,*"),
                ColumnSpacing = 4,
                RowSpacing = 2,
            };
            int rowsCount = (dates.Count + 3) / 4;
            for (int i = 0; i < rowsCount; i++)
                grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
            for (int i = 0; i < dates.Count; i++)
            {
                int r = i / 4;
                int c = i % 4;
                var cell = new Border
                {
                    Background = isRenamed ? AppTheme.BgActiveGreen : Brush.Parse("#1e2030"),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(6, 2),
                    Child = new TextBlock
                    {
                        Text = dates[i],
                        FontSize = AppTheme.FontSM, FontFamily = Font,
                        Foreground = isRenamed ? AppTheme.FgSuccess : AppTheme.FgMuted,
                        HorizontalAlignment = HorizontalAlignment.Center,
                    },
                };
                Grid.SetRow(cell, r);
                Grid.SetColumn(cell, c);
                grid.Children.Add(cell);
            }
            _bottomContent.Children.Add(grid);
        }
    }

    private void ExecuteRename(object? sender, RoutedEventArgs e)
    {
        if (_statusLabel == null) return;
        if (string.IsNullOrEmpty(_abbr))
        {
            _statusLabel.Foreground = AppTheme.FgWarn;
            _statusLabel.Text = "약칭을 먼저 선택하세요.";
            return;
        }
        var targetName = _targetBox?.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(targetName))
        {
            _statusLabel.Foreground = AppTheme.FgWarn;
            _statusLabel.Text = "오른쪽 텍스트박스에 목표 시료명을 입력하세요.";
            return;
        }
        if (_aliases.Count == 0)
        {
            _statusLabel.Foreground = AppTheme.FgWarn;
            _statusLabel.Text = "변경할 시료명을 추가하세요.";
            return;
        }

        int total = 0;
        var errors = new List<string>();
        var renamed = new HashSet<string>();
        foreach (var alias in _aliases)
        {
            try
            {
                total += AnalysisRequestService.RenameSampleName(_abbr, alias, targetName);
                renamed.Add(alias);
                // 각 시료 변경 완료 후 즉시 리프레시 (실시간 표시)
                RefreshBottomPanel(renamed);
            }
            catch (Exception ex) { errors.Add($"[{alias}] {ex.Message}"); }
        }

        if (errors.Count > 0)
        {
            if (App.EnableLogging)
            {
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "..", "..", "..", "WaterQualityRenameError.log");
                File.AppendAllText(logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {_abbr} → {targetName}\n" +
                    string.Join("\n", errors) + "\n\n");
            }
        }

        _statusLabel.Foreground = errors.Count == 0
            ? AppTheme.FgSuccess : AppTheme.FgDanger;
        _statusLabel.Text = errors.Count == 0
            ? $"✅  {total}건 변경 완료"
            : $"⚠  {total}건 변경, 오류 {errors.Count}건: {errors[0]}";

        _aliases.Clear();
        RefreshMappingPanel();
        RefreshSampleTree();
    }

    // ═══════════════════════════════════════════════════════════════════
    // Show4: 목표명 텍스트박스 + 시료명 목록
    // ═══════════════════════════════════════════════════════════════════
    private Control BuildSamplePanel()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };

        // 헤더
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "📋  시료명 목록 (중복제거)",
                FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppTheme.FgMuted,
            }
        });

        // 상태 레이블
        _sampleStatus = new TextBlock
        {
            FontSize   = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Regular,
            Foreground = AppTheme.FgDimmed,
            Margin     = new Thickness(8, 3),
        };
        Grid.SetRow(_sampleStatus, 1);
        root.Children.Add(_sampleStatus);

        // 시료명 트리
        _sampleTree = new TreeView { Margin = new Thickness(5) };
        _sampleTree.SelectionChanged += SampleTree_SelectionChanged;
        Grid.SetRow(_sampleTree, 2);
        root.Children.Add(_sampleTree);

        return root;
    }

    private void RefreshSampleTree()
    {
        if (_sampleTree == null) return;
        _sampleTree.Items.Clear();
        if (string.IsNullOrEmpty(_abbr))
        {
            if (_sampleStatus != null) _sampleStatus.Text = "약칭을 선택하세요";
            return;
        }
        try
        {
            var names = AnalysisRequestService.GetDistinctSampleNames(_abbr)
                            .OrderBy(n => n, StringComparer.CurrentCulture)
                            .ToList();
            if (_sampleStatus != null)
            {
                _sampleStatus.Foreground = AppTheme.FgDimmed;
                _sampleStatus.Text = $"{_abbr} — {names.Count}개";
            }
            // 방류기준 후보 + 현재 약칭의 매핑 로드
            var dischargeNames = new List<string> { "" };  // "" = 미지정
            try { dischargeNames.AddRange(AnalysisRequestService.GetDischargeStandardNames()); }
            catch { }
            var currentMap = ETA.Services.Common.SampleDischargeStandardService
                .GetByAbbr(_abbr)
                .ToDictionary(x => x.시료명, x => x.방류기준, StringComparer.OrdinalIgnoreCase);

            foreach (var name in names)
            {
                var capturedName = name;
                var nameText = new TextBlock
                {
                    Text       = capturedName, FontSize = AppTheme.FontBase, FontFamily = Font,
                    FontWeight = FontWeight.Regular,
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                Grid.SetColumn(nameText, 0);

                var cmb = new ComboBox
                {
                    Width = 160,
                    FontSize = AppTheme.FontSM, FontFamily = Font,
                    ItemsSource = dischargeNames,
                    SelectedItem = currentMap.TryGetValue(capturedName, out var cur) ? cur : "",
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                };
                Grid.SetColumn(cmb, 1);
                cmb.SelectionChanged += (_, _) =>
                {
                    var sel = (cmb.SelectedItem as string) ?? "";
                    if (string.IsNullOrWhiteSpace(sel))
                        ETA.Services.Common.SampleDischargeStandardService.Remove(_abbr, capturedName);
                    else
                        ETA.Services.Common.SampleDischargeStandardService.AddOrUpdate(_abbr, capturedName, sel);
                };
                // ComboBox 클릭이 TreeViewItem 의 드래그/선택으로 전파되지 않도록
                cmb.PointerPressed += (_, e) => e.Handled = true;

                var headerGrid = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("*,Auto"),
                };
                headerGrid.Children.Add(nameText);
                headerGrid.Children.Add(cmb);

                var tvi = new TreeViewItem
                {
                    Tag    = capturedName,
                    Header = headerGrid,
                };
                // 드래그 시작 — 이름 영역 클릭에서만
                nameText.PointerPressed += async (s, e) =>
                {
                    if (e.GetCurrentPoint(nameText).Properties.IsLeftButtonPressed)
                    {
                        var data = new Avalonia.Input.DataObject();
                        data.Set(Avalonia.Input.DataFormats.Text, capturedName);
                        await Avalonia.Input.DragDrop.DoDragDrop(e, data,
                            Avalonia.Input.DragDropEffects.Copy);
                    }
                };
                _sampleTree.Items.Add(tvi);
            }
        }
        catch (Exception ex)
        {
            if (_sampleStatus != null)
            {
                _sampleStatus.Foreground = AppTheme.FgDanger;
                _sampleStatus.Text = $"오류: {ex.Message}";
            }
        }
    }

    private void SampleTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        string? name = null;
        if (e.AddedItems.Count > 0 && e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is string s)
            name = s;
        if (string.IsNullOrEmpty(name)) return;

        if (string.IsNullOrEmpty(_abbr))
        {
            if (_statusLabel != null)
            {
                _statusLabel.Foreground = AppTheme.FgWarn;
                _statusLabel.Text = "먼저 왼쪽에서 약칭을 선택하세요.";
            }
            return;
        }

        if (!_aliases.Contains(name))
        {
            _aliases.Add(name);
            RefreshMappingPanel();
        }
    }

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
}
