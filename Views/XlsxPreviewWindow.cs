using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ETA.Views;

/// <summary>
/// 범용 xlsx 뷰어 — 경로를 넘기면 시트 드롭다운 + DataGrid 로 표시.
///   new XlsxPreviewWindow(path).Show();
/// </summary>
public class XlsxPreviewWindow : Window
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    private readonly string _path;
    private XLWorkbook? _wb;
    private readonly ComboBox _sheetCombo;
    private readonly DataGrid _grid;
    private readonly TextBlock _statusTb;
    private readonly TextBlock _metaTb;

    public XlsxPreviewWindow(string path)
    {
        _path = path;

        Title                 = $"📊 xlsx — {Path.GetFileName(path)}";
        Width                 = 1100;
        Height                = 680;
        MinWidth              = 600;
        MinHeight             = 400;
        CanResize             = true;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        SystemDecorations     = SystemDecorations.Full;
        Background            = AppTheme.BgPrimary;
        FontFamily            = Font;

        _sheetCombo = new ComboBox
        {
            FontSize = AppTheme.FontBase, FontFamily = Font,
            Height = 30, Padding = new Thickness(8, 2), Margin = new Thickness(0, 0, 6, 0),
            MinWidth = 280,
            Background = AppTheme.BgCard, BorderBrush = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
        };
        _sheetCombo.SelectionChanged += (_, _) => LoadCurrentSheet();

        _statusTb = new TextBlock
        {
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(10, 0),
        };
        _metaTb = new TextBlock
        {
            FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        };

        _grid = new DataGrid
        {
            FontSize              = AppTheme.FontSM,
            FontFamily            = Font,
            CanUserReorderColumns = true,
            CanUserResizeColumns  = true,
            CanUserSortColumns    = true,
            GridLinesVisibility   = DataGridGridLinesVisibility.All,
            HeadersVisibility     = DataGridHeadersVisibility.All,
            IsReadOnly            = true,
            Background            = AppTheme.BgCard,
            Foreground            = AppTheme.FgPrimary,
            BorderBrush           = AppTheme.BorderDefault,
            BorderThickness       = new Thickness(1),
        };

        var topBar = new DockPanel { LastChildFill = true, Margin = new Thickness(8, 6) };
        var left = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        left.Children.Add(new TextBlock
        {
            Text = "시트:", FontSize = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0),
        });
        left.Children.Add(_sheetCombo);
        left.Children.Add(_metaTb);
        DockPanel.SetDock(left, Dock.Left);
        topBar.Children.Add(left);
        topBar.Children.Add(_statusTb);

        var root = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(topBar, Dock.Top);
        root.Children.Add(topBar);
        root.Children.Add(new Border { Padding = new Thickness(8, 0, 8, 8), Child = _grid });

        Content = root;

        Opened += (_, _) => LoadWorkbook();
        Closed += (_, _) => { try { _wb?.Dispose(); } catch { } };
    }

    private void LoadWorkbook()
    {
        try
        {
            if (!File.Exists(_path))
            {
                _statusTb.Text = $"⚠ 파일 없음: {_path}";
                _statusTb.Foreground = Brushes.OrangeRed;
                return;
            }
            using var fs = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _wb = new XLWorkbook(fs);
            _sheetCombo.Items.Clear();
            foreach (var ws in _wb.Worksheets)
                _sheetCombo.Items.Add(ws.Name);
            if (_sheetCombo.Items.Count > 0) _sheetCombo.SelectedIndex = 0;

            var fi = new FileInfo(_path);
            _statusTb.Text = $"시트 {_wb.Worksheets.Count}개  ·  {fi.Length / 1024.0 / 1024.0:F2} MB  ·  {fi.LastWriteTime:yyyy-MM-dd HH:mm}";
            _statusTb.Foreground = AppTheme.FgMuted;
        }
        catch (Exception ex)
        {
            _statusTb.Text = $"⚠ 로드 실패: {ex.Message}";
            _statusTb.Foreground = Brushes.OrangeRed;
        }
    }

    private void LoadCurrentSheet()
    {
        if (_wb == null) return;
        var name = _sheetCombo.SelectedItem as string;
        if (string.IsNullOrEmpty(name)) return;

        try
        {
            var ws = _wb.Worksheet(name);
            var used = ws.RangeUsed();
            if (used == null)
            {
                _grid.Columns.Clear();
                _grid.ItemsSource = new List<Dictionary<string, object?>>();
                _metaTb.Text = "  (빈 시트)";
                return;
            }

            int firstRow = used.FirstRow().RowNumber();
            int lastRow  = used.LastRow().RowNumber();
            int firstCol = used.FirstColumn().ColumnNumber();
            int lastCol  = used.LastColumn().ColumnNumber();

            // 헤더 = 첫 행, 컬럼명 중복/빈 값은 A/B/C... 대체
            var headerRow = ws.Row(firstRow);
            var headers = new List<string>();
            var seen = new HashSet<string>();
            for (int c = firstCol; c <= lastCol; c++)
            {
                var raw = headerRow.Cell(c).GetString().Trim();
                if (string.IsNullOrEmpty(raw)) raw = XLHelper.GetColumnLetterFromNumber(c);
                var h = raw;
                int dup = 2;
                while (seen.Contains(h)) h = $"{raw}_{dup++}";
                seen.Add(h);
                headers.Add(h);
            }

            _grid.Columns.Clear();
            for (int i = 0; i < headers.Count; i++)
            {
                int colIdx = i;
                // Avalonia 11.3 의 인덱서 binding parser 가 "[0]" 형식을 거부하므로
                // FuncDataTemplate 으로 우회 — 데이터 행(string?[]) 의 colIdx 셀을 직접 표시.
                _grid.Columns.Add(new DataGridTemplateColumn
                {
                    Header  = headers[i],
                    Width   = new DataGridLength(120),
                    CellTemplate = new Avalonia.Controls.Templates.FuncDataTemplate<string?[]>(
                        (row, _) => new TextBlock
                        {
                            Text = row != null && colIdx < row.Length ? (row[colIdx] ?? "") : "",
                            FontFamily = Font,
                            FontSize = AppTheme.FontSM,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(4, 2),
                        }),
                });
            }

            // 행 데이터
            var rows = new ObservableCollection<string?[]>();
            int dataStart = firstRow + 1;
            for (int r = dataStart; r <= lastRow; r++)
            {
                var cells = new string?[headers.Count];
                for (int c = firstCol, idx = 0; c <= lastCol; c++, idx++)
                {
                    var cell = ws.Cell(r, c);
                    cells[idx] = cell.IsEmpty() ? null : cell.GetString();
                }
                rows.Add(cells);
            }
            _grid.ItemsSource = rows;
            _metaTb.Text = $"  · {rows.Count:N0}행 × {headers.Count}열";
        }
        catch (Exception ex)
        {
            _statusTb.Text = $"⚠ 시트 로드 실패: {ex.Message}";
            _statusTb.Foreground = Brushes.OrangeRed;
        }
    }
}
