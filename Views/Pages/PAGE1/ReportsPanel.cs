using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

/// <summary>Content4 — Data/Reports 폴더 파일 목록. 체크박스 + 우클릭 삭제.</summary>
public class ReportsPanel : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static string ReportsDir =>
        Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "Data", "Reports"));

    private readonly StackPanel _fileList  = new() { Spacing = 2 };
    private readonly TextBlock  _countText = new()
    {
        FontSize = AppTheme.FontBase, FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        Foreground = AppTheme.FgDimmed,
        Margin = new Avalonia.Thickness(2, 0, 0, 4),
    };

    // 체크박스 → 파일 매핑
    private readonly Dictionary<CheckBox, FileInfo> _checkMap = new();
    // 전체선택 체크박스 참조 (삭제 후 해제용)
    private CheckBox? _chkAll;

    // 폴더 변경 감시 — 시험성적서 출력 직후 자동 갱신
    private FileSystemWatcher? _watcher;
    private DispatcherTimer? _refreshTimer;

    public ReportsPanel()
    {
        BuildUI();
        LoadFiles();
        StartWatching();
        DetachedFromVisualTree += (_, _) => StopWatching();
    }

    private void StartWatching()
    {
        try
        {
            Directory.CreateDirectory(ReportsDir);

            // 다중 이벤트(쓰기 중 Created→Changed×N) 디바운스용 타이머
            _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _refreshTimer.Tick += (_, _) =>
            {
                _refreshTimer!.Stop();
                LoadFiles();
            };

            _watcher = new FileSystemWatcher(ReportsDir)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
                IncludeSubdirectories = false,
                EnableRaisingEvents = true,
            };
            _watcher.Created += OnDirChanged;
            _watcher.Deleted += OnDirChanged;
            _watcher.Renamed += OnDirChanged;
            _watcher.Changed += OnDirChanged;
        }
        catch { }
    }

    private void StopWatching()
    {
        try
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Created -= OnDirChanged;
                _watcher.Deleted -= OnDirChanged;
                _watcher.Renamed -= OnDirChanged;
                _watcher.Changed -= OnDirChanged;
                _watcher.Dispose();
                _watcher = null;
            }
            _refreshTimer?.Stop();
        }
        catch { }
    }

    private void OnDirChanged(object? sender, FileSystemEventArgs e)
    {
        // FSW 이벤트는 백그라운드 스레드 → UI 스레드로 마샬링 후 디바운스 재시작
        Dispatcher.UIThread.Post(() =>
        {
            _refreshTimer?.Stop();
            _refreshTimer?.Start();
        });
    }

    // ── UI 구성 ───────────────────────────────────────────────────────────
    private void BuildUI()
    {
        // ── 헤더 ──────────────────────────────────────────────────────────
        var header = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        header.Children.Add(new TextBlock
        {
            Text = "📂  출력 보관함",
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 0,
        });
        var refreshBtn = MakeBtn("🔄", 30, "#2a2a3a", "#aaa");
        refreshBtn.Click += (_, _) => LoadFiles();
        refreshBtn[Grid.ColumnProperty] = 1;
        header.Children.Add(refreshBtn);

        // ── 툴바: 전체선택 / 선택삭제 / 폴더열기 ─────────────────────────
        var toolbar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            Margin = new Avalonia.Thickness(0, 0, 0, 4),
        };

        _chkAll = new CheckBox
        {
            Content = "전체", FontSize = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        };
        _chkAll.IsCheckedChanged += (_, _) =>
        {
            bool check = _chkAll.IsChecked == true;
            foreach (var cb in _checkMap.Keys)
                cb.IsChecked = check;
        };
        var chkAll = _chkAll;

        var deleteSelBtn = MakeBtn("🗑 선택 삭제", 0, "#5a2020", "#f0aeae");
        deleteSelBtn.Padding = new Avalonia.Thickness(10, 0);
        deleteSelBtn.Click += async (_, _) => await DeleteCheckedAsync();

        var folderBtn = MakeBtn("폴더 열기", 0, "#2a2a3a", "#aaa");
        folderBtn.Padding = new Avalonia.Thickness(10, 0);
        folderBtn.Click += (_, _) => OpenFolder();

        toolbar.Children.Add(chkAll);
        toolbar.Children.Add(deleteSelBtn);
        toolbar.Children.Add(folderBtn);

        // ── 컬럼 헤더 ─────────────────────────────────────────────────────
        var colHeader = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#16162a")),
            Padding = new Avalonia.Thickness(8, 3),
            Margin = new Avalonia.Thickness(0, 0, 0, 2),
            Child = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("24,Auto,*,60,55"),
                Children =
                {
                    ColHead("",       0),
                    ColHead("",       1),
                    ColHead("파일명", 2),
                    ColHead("날짜",   3),
                    ColHead("크기",   4),
                }
            }
        };

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content = _fileList,
        };

        var divider = new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#333")), Margin = new Avalonia.Thickness(0, 4, 0, 4) };

        Content = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,*"),
            Children =
            {
                Row(header,    0),
                Row(divider,   1),
                Row(_countText,2),
                Row(toolbar,   3),
                Row(colHeader, 4),
                Row(scroll,    5),
            }
        };
    }

    // ── 파일 목록 로드 ────────────────────────────────────────────────────
    public void LoadFiles()
    {
        _fileList.Children.Clear();
        _checkMap.Clear();

        try
        {
            Directory.CreateDirectory(ReportsDir);
            var files = Directory.GetFiles(ReportsDir)
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.LastWriteTime)
                .ToList();

            _countText.Text = $"{files.Count}개 파일";

            if (files.Count == 0)
            {
                _fileList.Children.Add(new TextBlock
                {
                    Text = "출력된 파일이 없습니다.",
                    FontSize = AppTheme.FontBase, FontFamily = Font,
                    Foreground = AppTheme.FgDimmed,
                    Margin = new Avalonia.Thickness(4, 8),
                });
                return;
            }
            foreach (var f in files)
                _fileList.Children.Add(MakeFileRow(f));
        }
        catch (Exception ex)
        {
        }
    }

    // ── 파일 행 빌드 ──────────────────────────────────────────────────────
    private Border MakeFileRow(FileInfo file)
    {
        var icon = file.Extension.ToLower() switch
        {
            ".xlsx" or ".xls" => "📊",
            ".pdf"            => "📄",
            ".csv"            => "📋",
            _                 => "📎",
        };

        // 체크박스
        var chk = new CheckBox
        {
            IsChecked = false,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(0, 0, 2, 0),
            [Grid.ColumnProperty] = 0,
        };
        _checkMap[chk] = file;

        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("24,Auto,*,60,55"),
        };
        row.Children.Add(chk);

        // 아이콘
        row.Children.Add(new TextBlock
        {
            Text = icon, FontSize = AppTheme.FontLG,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(0, 0, 6, 0),
            [Grid.ColumnProperty] = 1,
        });

        // 파일명
        row.Children.Add(new TextBlock
        {
            Text = file.Name, FontSize = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgSecondary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 2,
        });

        // 날짜
        row.Children.Add(new TextBlock
        {
            Text = file.LastWriteTime.ToString("MM-dd HH:mm"),
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 3,
        });

        // 크기
        row.Children.Add(new TextBlock
        {
            Text = FormatSize(file.Length),
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = new SolidColorBrush(Color.Parse("#444466")),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Right,
            [Grid.ColumnProperty] = 4,
        });

        var border = new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new Avalonia.CornerRadius(3),
            Padding      = new Avalonia.Thickness(8, 5),
            Margin       = new Avalonia.Thickness(0, 0, 0, 2),
            Child        = row,
        };

        // 호버
        border.PointerEntered += (_, _) =>
            border.Background = AppTheme.BgSecondary;
        border.PointerExited += (_, _) =>
            border.Background = new SolidColorBrush(
                Color.Parse(chk.IsChecked == true ? "#1e2e1e" : "#1e1e2e"));

        // 체크 시 행 색상 변경
        chk.IsCheckedChanged += (_, _) =>
            border.Background = new SolidColorBrush(
                Color.Parse(chk.IsChecked == true ? "#1e2e1e" : "#1e1e2e"));

        // 더블클릭 → 파일 열기
        border.DoubleTapped += (_, _) => OpenFile(file.FullName);

        // 우클릭 → 컨텍스트 메뉴
        border.PointerPressed += (_, e) =>
        {
            var pt = e.GetCurrentPoint(border).Properties;
            if (pt.IsRightButtonPressed)
            {
                e.Handled = true;
                ShowContextMenu(border, file, chk);
            }
        };

        return border;
    }

    // ── 체크된 파일 일괄 삭제 ─────────────────────────────────────────────
    private async Task DeleteCheckedAsync()
    {
        var targets = _checkMap
            .Where(kv => kv.Key.IsChecked == true)
            .Select(kv => kv.Value)
            .ToList();

        if (targets.Count == 0) return;

        bool confirmed = await ShowDeleteConfirmAsync(
            targets.Count == 1
                ? $"'{targets[0].Name}'"
                : $"선택된 {targets.Count}개 파일");

        if (!confirmed) return;

        int ok = 0;
        foreach (var f in targets)
        {
            try { f.Delete(); ok++; }
            catch (Exception ex)
            {
            }
        }


        // 삭제 후 전체 체크박스 해제
        if (_chkAll != null) _chkAll.IsChecked = false;

        LoadFiles();  // 목록 새로고침 (개별 체크박스도 재생성되므로 자동 해제)
    }

    // ── 우클릭 컨텍스트 메뉴 ─────────────────────────────────────────────
    private void ShowContextMenu(Control anchor, FileInfo file, CheckBox chk)
    {
        var menu = new ContextMenu();

        var openItem = new MenuItem { Header = "📂  열기", FontFamily = Font, FontSize = AppTheme.FontMD };
        openItem.Click += (_, _) => OpenFile(file.FullName);

        var exploreItem = new MenuItem { Header = "🔍  탐색기에서 보기", FontFamily = Font, FontSize = AppTheme.FontMD };
        exploreItem.Click += (_, _) =>
        {
            try { Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"/select,\"{file.FullName}\"", UseShellExecute = false }); }
            catch { }
        };

        var checkItem = new MenuItem
        {
            Header = chk.IsChecked == true ? "☑  체크 해제" : "☐  체크",
            FontFamily = Font, FontSize = AppTheme.FontMD,
        };
        checkItem.Click += (_, _) => chk.IsChecked = !(chk.IsChecked == true);

        var deleteItem = new MenuItem
        {
            Header = "🗑  삭제", FontFamily = Font, FontSize = AppTheme.FontMD,
            Foreground = AppTheme.FgDanger,
        };
        deleteItem.Click += async (_, _) =>
        {
            bool ok = await ShowDeleteConfirmAsync($"'{file.Name}'");
            if (!ok) return;
            try { file.Delete(); LoadFiles(); }
            catch (Exception ex) { }
        };

        menu.Items.Add(openItem);
        menu.Items.Add(exploreItem);
        menu.Items.Add(checkItem);
        menu.Items.Add(new Separator());
        menu.Items.Add(deleteItem);
        menu.Open(anchor);
    }

    // ── 삭제 확인 다이얼로그 ─────────────────────────────────────────────
    private async Task<bool> ShowDeleteConfirmAsync(string target)
    {
        var dlg = new Window
        {
            Title = "삭제 확인", Width = 340, Height = 150, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgPrimary,
        };
        bool result = false;
        var yes = new Button { Content = "삭제", Width = 80, Background = new SolidColorBrush(Color.Parse("#8b2020")), Foreground = Brushes.White, BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4) };
        var no  = new Button { Content = "취소", Width = 80, Background = new SolidColorBrush(Color.Parse("#333")),    Foreground = Brushes.White, BorderThickness = new Avalonia.Thickness(0), CornerRadius = new Avalonia.CornerRadius(4) };
        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };
        dlg.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = $"{target}\n삭제하시겠습니까?", FontFamily = Font, FontSize = AppTheme.FontMD, Foreground = AppRes("AppFg"), TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8, Children = { no, yes } }
            }
        };
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
        return result;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────
    private static void OpenFile(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { }
    }

    private static void OpenFolder()
    {
        try { Directory.CreateDirectory(ReportsDir); Process.Start(new ProcessStartInfo(ReportsDir) { UseShellExecute = true }); }
        catch { }
    }

    private static TextBlock ColHead(string text, int col) => new()
    {
        Text = text, FontSize = AppTheme.FontSM, FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        Foreground = AppTheme.FgDimmed,
        VerticalAlignment = VerticalAlignment.Center,
        [Grid.ColumnProperty] = col,
    };

    private static Button MakeBtn(string text, double width, string bg, string fg)
    {
        var btn = new Button
        {
            Content = text, FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase, Height = 26,
            Background = new SolidColorBrush(Color.Parse(bg)),
            Foreground = new SolidColorBrush(Color.Parse(fg)),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(3),
            VerticalContentAlignment = VerticalAlignment.Center,
        };
        if (width > 0) btn.Width = width;
        return btn;
    }

    private static string FormatSize(long bytes) => bytes switch
    {
        < 1024        => $"{bytes}B",
        < 1024 * 1024 => $"{bytes / 1024.0:F1}KB",
        _             => $"{bytes / 1024.0 / 1024.0:F1}MB",
    };

    private static Control Row(Control ctrl, int row)
    {
        Grid.SetRow(ctrl, row); return ctrl;
    }
}
