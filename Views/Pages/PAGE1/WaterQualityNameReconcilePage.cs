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

    public WaterQualityNameReconcilePage()
    {
        LeftPanel   = BuildAbbrPanel();
        CenterPanel = BuildMappingPanel();
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
            foreach (var abbr in AnalysisRequestService.GetDistinctAbbreviations())
            {
                _abbrTree.Items.Add(new TreeViewItem
                {
                    Tag    = abbr,
                    Header = new TextBlock
                    {
                        Text       = abbr, FontSize = AppTheme.FontMD, FontFamily = Font,
                        FontWeight = FontWeight.Regular,
                        Foreground = AppRes("AppFg"),
                    }
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
            Watermark       = "앞으로 사용할 시료명...",
            Background      = Brush.Parse("#1e2030"),
            Foreground      = AppTheme.FgPrimary,
            BorderBrush     = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 5),
            Margin          = new Thickness(0, 0, 0, 4),
        };
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
        foreach (var alias in _aliases)
        {
            try { total += AnalysisRequestService.RenameSampleName(_abbr, alias, targetName); }
            catch (Exception ex) { errors.Add($"[{alias}] {ex.Message}"); }
        }

        if (errors.Count > 0)
        {
            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "..", "..", "..", "WaterQualityRenameError.log");
            File.AppendAllText(logPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {_abbr} → {targetName}\n" +
                string.Join("\n", errors) + "\n\n");
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
            foreach (var name in names)
            {
                _sampleTree.Items.Add(new TreeViewItem
                {
                    Tag    = name,
                    Header = new TextBlock
                    {
                        Text       = name, FontSize = AppTheme.FontBase, FontFamily = Font,
                        FontWeight = FontWeight.Regular,
                        Foreground = AppRes("AppFg"),
                    }
                });
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
