using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 설정 → 양식폴더. 각 양식(엑셀 템플릿)의 파일 경로를 개별 지정·복원할 수 있다.
/// 변경 즉시 Data/template_paths.json 에 저장되며, 각 서비스는
/// TemplateConfiguration.Resolve(key) 로 최종 경로를 얻는다.
/// </summary>
public sealed class TemplateSettingsPage
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    private readonly Dictionary<string, TextBlock> _pathBlocks = new();
    private readonly Dictionary<string, TextBlock> _statusBlocks = new();
    private readonly StackPanel _root;

    public Control View => _root;

    public TemplateSettingsPage()
    {
        _root = BuildRoot();
    }

    private StackPanel BuildRoot()
    {
        var outer = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        var titleRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            VerticalAlignment = VerticalAlignment.Center,
        };
        titleRow.Children.Add(new TextBlock
        {
            Text = "📁  양식(템플릿) 파일 경로 설정",
            FontFamily = Font,
            FontSize = AppTheme.FontXL,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
        });

        var btnOpenFolder = MakeButton("📂 양식폴더 열기", "#1a2a3a", "#7aaae8");
        btnOpenFolder.Click += (_, _) => OpenTemplatesFolder();

        var btnResetAll = MakeButton("↺ 전체 기본값 복원", "#3a1a1a", "#ee8888");
        btnResetAll.Click += (_, _) => ResetAll();

        titleRow.Children.Add(new Border { Width = 1 });
        titleRow.Children.Add(btnOpenFolder);
        titleRow.Children.Add(btnResetAll);
        outer.Children.Add(titleRow);

        outer.Children.Add(new TextBlock
        {
            Text = "각 양식의 ‘변경’ 버튼으로 다른 엑셀 파일을 지정할 수 있습니다. 변경 즉시 저장됩니다. "
                 + "경로를 비워두면(기본값 복원) 프로젝트 내 Data/Templates/… 경로를 사용합니다.",
            FontFamily = Font,
            FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
        });

        // 슬롯 목록
        var list = new StackPanel { Spacing = 2 };
        int i = 0;
        foreach (var slot in TemplateConfiguration.Slots)
        {
            list.Children.Add(BuildSlotRow(slot, i % 2 == 0));
            i++;
        }

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = list,
            Height = 560,
        };
        outer.Children.Add(scroll);

        return outer;
    }

    private Border BuildSlotRow(TemplateConfiguration.Slot slot, bool even)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("2*,4*,Auto,Auto,Auto"),
            Margin = new Thickness(4, 2),
        };

        // 라벨 + 사용 서비스
        var labelPanel = new StackPanel { Spacing = 1, VerticalAlignment = VerticalAlignment.Center };
        labelPanel.Children.Add(new TextBlock
        {
            Text = slot.Label,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        });
        labelPanel.Children.Add(new TextBlock
        {
            Text = slot.UsedBy,
            FontFamily = Font,
            FontSize = AppTheme.FontXS,
            Foreground = AppTheme.FgMuted,
        });
        Grid.SetColumn(labelPanel, 0);
        grid.Children.Add(labelPanel);

        // 현재 경로
        var resolved = TemplateConfiguration.Resolve(slot.Key);
        var pathBlock = new TextBlock
        {
            Text = resolved,
            FontFamily = Font,
            FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(8, 0),
        };
        ToolTip.SetTip(pathBlock, resolved);
        Grid.SetColumn(pathBlock, 1);
        grid.Children.Add(pathBlock);
        _pathBlocks[slot.Key] = pathBlock;

        // 상태 (✓/✗/폴더)
        var statusBlock = new TextBlock
        {
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0),
        };
        Grid.SetColumn(statusBlock, 2);
        grid.Children.Add(statusBlock);
        _statusBlocks[slot.Key] = statusBlock;
        UpdateStatus(slot.Key);

        // 변경 버튼
        var changeBtn = MakeButton("변경", "#2a3a4a", "#99bbdd");
        changeBtn.Click += async (_, _) => await ChangePath(slot);
        Grid.SetColumn(changeBtn, 3);
        grid.Children.Add(changeBtn);

        // 기본값 버튼
        var resetBtn = MakeButton("기본값", "#3a2a1a", "#ddaa88");
        resetBtn.Click += (_, _) => ResetOne(slot);
        Grid.SetColumn(resetBtn, 4);
        grid.Children.Add(resetBtn);

        return new Border
        {
            Background = even ? AppTheme.BgSecondary : AppTheme.BgCard,
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 4),
            Child = grid,
        };
    }

    private void UpdateStatus(string key)
    {
        if (!_statusBlocks.TryGetValue(key, out var block)) return;
        var path = TemplateConfiguration.Resolve(key);
        bool exists = Directory.Exists(path) || File.Exists(path);
        bool isDir = Directory.Exists(path);
        if (exists)
        {
            block.Text = isDir ? "📂" : "✓";
            block.Foreground = AppTheme.FgSuccess;
        }
        else
        {
            block.Text = "✗";
            block.Foreground = AppTheme.FgDanger;
        }
    }

    private async Task ChangePath(TemplateConfiguration.Slot slot)
    {
        var top = TopLevel.GetTopLevel(_root);
        if (top == null) return;

        var exts = slot.Key == "TestRecordBookFolder"
            ? null  // 폴더형 슬롯
            : new[] { new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx", "*.xlsm" } } };

        if (slot.Key == "TestRecordBookFolder")
        {
            var folders = await top.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                AllowMultiple = false,
                Title = $"{slot.Label} 폴더 선택",
            });
            if (folders == null || folders.Count == 0) return;
            var chosen = folders[0].Path.LocalPath;
            PersistOverride(slot.Key, chosen);
        }
        else
        {
            var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                AllowMultiple = false,
                Title = $"{slot.Label} 파일 선택",
                FileTypeFilter = exts,
            });
            if (files == null || files.Count == 0) return;
            var chosen = files[0].Path.LocalPath;
            PersistOverride(slot.Key, chosen);
        }
    }

    private void PersistOverride(string key, string absolutePath)
    {
        var map = TemplateConfiguration.Snapshot();
        map[key] = absolutePath;
        TemplateConfiguration.Save(map);
        RefreshRow(key);
    }

    private void ResetOne(TemplateConfiguration.Slot slot)
    {
        var map = TemplateConfiguration.Snapshot();
        map.Remove(slot.Key);
        TemplateConfiguration.Save(map);
        RefreshRow(slot.Key);
    }

    private void ResetAll()
    {
        TemplateConfiguration.Save(new Dictionary<string, string>());
        foreach (var slot in TemplateConfiguration.Slots) RefreshRow(slot.Key);
    }

    private void RefreshRow(string key)
    {
        if (_pathBlocks.TryGetValue(key, out var block))
        {
            var path = TemplateConfiguration.Resolve(key);
            block.Text = path;
            ToolTip.SetTip(block, path);
        }
        UpdateStatus(key);
    }

    private static void OpenTemplatesFolder()
    {
        try
        {
            var path = Path.Combine(TemplateConfiguration.RepoRoot, "Data", "Templates");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true,
            });
        }
        catch { }
    }

    private static Button MakeButton(string text, string bg, string fg) => new()
    {
        Content = text,
        FontFamily = Font,
        FontSize = AppTheme.FontBase,
        Background = new SolidColorBrush(Color.Parse(bg)),
        Foreground = new SolidColorBrush(Color.Parse(fg)),
        BorderThickness = new Thickness(0),
        Padding = new Thickness(10, 4),
        CornerRadius = new CornerRadius(4),
        VerticalAlignment = VerticalAlignment.Center,
    };
}
