using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 사용자메뉴얼 페이지.
///   Show1 : 메뉴 트리 (카테고리 → 메뉴)
///   Show2 : 선택된 메뉴의 MD 문서 렌더 + 경로 편집
/// 메뉴 ↔ MD 매핑은 `사용자메뉴얼` DB 테이블 (UserManualService) 에 보관.
/// </summary>
public class UserManualPage
{
    private static readonly FontFamily Font     = new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly FontFamily MonoFont = new("Menlo,Consolas,monospace");

    public Control TreeControl    { get; }
    public Control ContentControl { get; }

    private readonly TreeView _tree;
    private readonly StackPanel _mdPanel;
    private readonly TextBlock _titleBlock;
    private readonly TextBox _pathBox;
    private readonly Button _btnSavePath;
    private readonly Button _btnReload;
    private readonly Button _btnOpenInEditor;
    private UserManualService.ManualEntry? _current;

    public UserManualPage()
    {
        UserManualService.EnsureTable();
        UserManualService.SeedFromAccessMenus();
        UserManualService.EnsureStubFiles();

        // ── Show1: 카테고리별 메뉴 트리 ──────────────────────────────────────
        _tree = new TreeView { Margin = new Thickness(4) };
        _tree.SelectionChanged += OnTreeSelection;
        LoadTree();

        var treeRoot = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        treeRoot.Children.Add(new TextBlock
        {
            Text = "📚  사용자메뉴얼",
            FontFamily = Font, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            Margin = new Thickness(8, 8, 0, 6),
        });
        Grid.SetRow(_tree, 1);
        treeRoot.Children.Add(_tree);
        TreeControl = treeRoot;

        // ── Show2: 제목 + 경로편집 + MD 본문 ─────────────────────────────────
        _titleBlock = new TextBlock
        {
            Text = "메뉴를 선택하세요", FontFamily = Font, FontSize = AppTheme.FontXL,
            FontWeight = FontWeight.SemiBold, Foreground = AppTheme.FgPrimary,
            Margin = new Thickness(0, 0, 0, 6),
        };

        _pathBox = new TextBox
        {
            FontFamily = MonoFont, FontSize = AppTheme.FontSM,
            Watermark = "Docs/Manuals/{메뉴키}.md",
            Height = 28, Padding = new Thickness(6, 2),
        };

        _btnSavePath = MakeBtn("💾 경로저장", AppTheme.BgActiveBlue, AppTheme.FgInfo);
        _btnSavePath.Click += (_, _) => SaveCurrentPath();

        _btnReload = MakeBtn("🔄 새로고침", AppTheme.BorderSubtle, AppTheme.FgMuted);
        _btnReload.Click += (_, _) => RenderCurrentMd();

        _btnOpenInEditor = MakeBtn("📝 편집기로 열기", AppTheme.BorderSubtle, AppTheme.FgInfo);
        _btnOpenInEditor.Click += (_, _) => OpenCurrentMdInEditor();

        var pathRow = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto,Auto"),
            ColumnSpacing = 4,
            Margin = new Thickness(0, 0, 0, 8),
        };
        Grid.SetColumn(_pathBox, 0);     pathRow.Children.Add(_pathBox);
        Grid.SetColumn(_btnSavePath, 1); pathRow.Children.Add(_btnSavePath);
        Grid.SetColumn(_btnReload, 2);   pathRow.Children.Add(_btnReload);
        Grid.SetColumn(_btnOpenInEditor, 3); pathRow.Children.Add(_btnOpenInEditor);

        _mdPanel = new StackPanel { Spacing = 4 };

        var contentRoot = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,*"),
            Margin = new Thickness(12),
        };
        Grid.SetRow(_titleBlock, 0); contentRoot.Children.Add(_titleBlock);
        Grid.SetRow(pathRow, 1);     contentRoot.Children.Add(pathRow);

        var scroll = new ScrollViewer
        {
            Content = _mdPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        Grid.SetRow(scroll, 2);
        contentRoot.Children.Add(scroll);
        ContentControl = contentRoot;
    }

    private static Button MakeBtn(string txt, IBrush bg, IBrush fg) => new()
    {
        Content = txt, FontFamily = Font, FontSize = AppTheme.FontSM,
        Padding = new Thickness(10, 4),
        Background = bg, Foreground = fg,
        BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderMuted,
        CornerRadius = new CornerRadius(4),
        Cursor = new Cursor(StandardCursorType.Hand),
    };

    private void LoadTree()
    {
        _tree.Items.Clear();
        var entries = UserManualService.GetAll();
        var grouped = entries
            .GroupBy(e => e.Category)
            .OrderBy(g => g.Key);

        foreach (var grp in grouped)
        {
            var catNode = new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text = $"📂 {grp.Key}", FontFamily = Font, FontSize = AppTheme.FontMD,
                    FontWeight = FontWeight.SemiBold, Foreground = AppTheme.FgInfo,
                },
                IsExpanded = true,
            };
            foreach (var entry in grp)
            {
                var leaf = new TreeViewItem
                {
                    Header = new TextBlock
                    {
                        Text = $"📄 {entry.Label}", FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgPrimary,
                    },
                    Tag = entry,
                };
                catNode.Items.Add(leaf);
            }
            _tree.Items.Add(catNode);
        }
    }

    private void OnTreeSelection(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0 || e.AddedItems[0] is not TreeViewItem tvi) return;
        if (tvi.Tag is not UserManualService.ManualEntry entry) return;
        _current = entry;
        _titleBlock.Text = $"📖  {entry.Category} · {entry.Label}";
        _pathBox.Text    = entry.MdPath;
        RenderCurrentMd();
    }

    private void RenderCurrentMd()
    {
        _mdPanel.Children.Clear();
        if (_current == null)
        {
            _mdPanel.Children.Add(new TextBlock
            {
                Text = "좌측에서 메뉴를 선택하세요.",
                FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppTheme.FgMuted,
            });
            return;
        }

        var md = UserManualService.ReadMd(_current.MdPath, _current.Label);
        foreach (var ctrl in RenderMarkdown(md))
            _mdPanel.Children.Add(ctrl);
    }

    private void SaveCurrentPath()
    {
        if (_current == null) return;
        var newPath = _pathBox.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(newPath)) return;
        UserManualService.SetMdPath(_current.Key, newPath);
        _current = UserManualService.GetByKey(_current.Key);
        RenderCurrentMd();
    }

    private void OpenCurrentMdInEditor()
    {
        if (_current == null) return;
        try
        {
            var resolved = Path.IsPathRooted(_current.MdPath)
                ? _current.MdPath
                : Path.Combine(AppPaths.RootPath, _current.MdPath);
            if (!File.Exists(resolved)) return;
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName        = resolved,
                UseShellExecute = true,
            });
        }
        catch (Exception ex) { Console.WriteLine($"[OpenMd] {ex.Message}"); }
    }

    // =========================================================================
    // 간이 마크다운 → Avalonia Control 렌더링
    //   지원: # ## ###  -  > ```  **bold**  `inline`
    // =========================================================================
    private static IEnumerable<Control> RenderMarkdown(string md)
    {
        var lines = md.Replace("\r\n", "\n").Split('\n');
        bool inCode = false;
        var codeBuf = new System.Text.StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = rawLine;

            if (line.TrimStart().StartsWith("```"))
            {
                if (inCode)
                {
                    yield return new Border
                    {
                        Background = AppTheme.BgSecondary,
                        BorderBrush = AppTheme.BorderMuted,
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(4),
                        Padding = new Thickness(8, 6),
                        Margin = new Thickness(0, 4),
                        Child = new SelectableTextBlock
                        {
                            Text = codeBuf.ToString().TrimEnd(),
                            FontFamily = MonoFont, FontSize = AppTheme.FontSM,
                            Foreground = AppTheme.FgPrimary,
                        },
                    };
                    codeBuf.Clear();
                    inCode = false;
                }
                else inCode = true;
                continue;
            }
            if (inCode) { codeBuf.AppendLine(line); continue; }

            if (string.IsNullOrWhiteSpace(line)) { yield return new Border { Height = 6 }; continue; }

            if (line.StartsWith("### "))
            {
                yield return MakeHeading(line[4..], AppTheme.FontLG, AppTheme.FgInfo);
                continue;
            }
            if (line.StartsWith("## "))
            {
                yield return MakeHeading(line[3..], AppTheme.FontXL, AppTheme.FgPrimary);
                continue;
            }
            if (line.StartsWith("# "))
            {
                yield return MakeHeading(line[2..], 18, AppTheme.FgPrimary, topMargin: 10);
                continue;
            }
            if (line.StartsWith("> "))
            {
                yield return new Border
                {
                    BorderBrush = AppTheme.BorderInfo, BorderThickness = new Thickness(3, 0, 0, 0),
                    Padding = new Thickness(10, 2, 0, 2),
                    Margin = new Thickness(0, 2),
                    Child = MakeInline(line[2..], italic: true, fg: AppTheme.FgMuted),
                };
                continue;
            }
            if (line.StartsWith("- ") || line.StartsWith("* "))
            {
                var row = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(12, 1, 0, 1),
                };
                row.Children.Add(new TextBlock
                {
                    Text = "•", FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppTheme.FgMuted,
                    Margin = new Thickness(0, 0, 6, 0),
                });
                row.Children.Add(MakeInline(line[2..], fg: AppTheme.FgPrimary));
                yield return row;
                continue;
            }

            yield return MakeInline(line, fg: AppTheme.FgPrimary);
        }
    }

    private static TextBlock MakeHeading(string text, double fontSize, IBrush fg, double topMargin = 6)
    {
        return new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = fontSize,
            FontWeight = FontWeight.SemiBold, Foreground = fg,
            Margin = new Thickness(0, topMargin, 0, 4),
            TextWrapping = TextWrapping.Wrap,
        };
    }

    private static SelectableTextBlock MakeInline(string text, bool italic = false, IBrush? fg = null)
    {
        var tb = new SelectableTextBlock
        {
            FontFamily = Font, FontSize = AppTheme.FontMD,
            Foreground = fg ?? AppTheme.FgPrimary,
            FontStyle  = italic ? FontStyle.Italic : FontStyle.Normal,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 1),
        };
        // **bold** 와 `inline` 분해
        var inlines = new InlineCollection();
        ParseInline(text, inlines);
        tb.Inlines = inlines;
        return tb;
    }

    private static void ParseInline(string text, InlineCollection target)
    {
        int i = 0;
        var buf = new System.Text.StringBuilder();
        void Flush()
        {
            if (buf.Length > 0)
            {
                target.Add(new Run(buf.ToString()));
                buf.Clear();
            }
        }
        while (i < text.Length)
        {
            if (i + 1 < text.Length && text[i] == '*' && text[i + 1] == '*')
            {
                int end = text.IndexOf("**", i + 2, StringComparison.Ordinal);
                if (end > 0)
                {
                    Flush();
                    target.Add(new Run(text[(i + 2)..end]) { FontWeight = FontWeight.Bold });
                    i = end + 2; continue;
                }
            }
            if (text[i] == '`')
            {
                int end = text.IndexOf('`', i + 1);
                if (end > 0)
                {
                    Flush();
                    target.Add(new Run(text[(i + 1)..end])
                    {
                        FontFamily = MonoFont,
                        Background = AppTheme.BgSecondary,
                        Foreground = AppTheme.FgInfo,
                    });
                    i = end + 1; continue;
                }
            }
            buf.Append(text[i]);
            i++;
        }
        Flush();
    }
}
