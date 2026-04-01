using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services.SERVICE2;
using ETA.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ETA.Views.Pages.PAGE2;

public partial class ProcessingFacilityPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }
    // ── 외부 이벤트 ──────────────────────────────────────────────────────
    public event Action<Control?>? ResultGridChanged;

    // ── 상태 ─────────────────────────────────────────────────────────────
    private string? _selectedFacility;
    private List<FacilityResultRow> _currentRows = new();

    // TextBox 참조: (마스터Id, 항목명) → TextBox
    private readonly Dictionary<(int, string), TextBox> _cells = new();

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // 컬럼 정의: (헤더, 항목키, 너비)
    private static readonly (string Header, string Key, int Width)[] Cols =
    [
        ("BOD",    "BOD",    62),
        ("TOC",    "TOC",    62),
        ("SS",     "SS",     62),
        ("T-N",    "TN",     62),
        ("T-P",    "TP",     62),
        ("총대장균", "총대장균군", 70),
        ("COD",    "COD",    62),
        ("염소이온", "염소이온",  72),
        ("영양염류", "영양염류",  72),
        ("함수율",  "함수율",   62),
        ("중금속",  "중금속",   62),
        ("비고",   "비고",    110),
    ];

    public ProcessingFacilityPage()
    {
        InitializeComponent();
        dpDate.SelectedDate = DateTime.Today;
        LoadFacilityButtons();
    }

    // =========================================================================
    // 시설 버튼 동적 생성
    // =========================================================================
    private void LoadFacilityButtons()
    {
        try
        {
            var names = FacilityResultService.GetFacilityNames();
            foreach (var name in names)
            {
                var btn = new Button
                {
                    Content = name,
                    Tag     = name,
                    Margin  = new Thickness(2),
                    Classes = { "FacBtn" },
                };
                btn.Click += FacilityBtn_Click;
                FacilityButtonPanel.Children.Add(btn);
            }
        }
        catch (Exception ex)
        {
            tbStatus.Foreground = Brush.Parse("#ee4444");
            tbStatus.Text = $"시설 목록 오류: {ex.Message}";
        }
    }

    // =========================================================================
    // 시설 버튼 클릭
    // =========================================================================
    private void FacilityBtn_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;

        foreach (var child in FacilityButtonPanel.Children)
            if (child is Button b) b.Classes.Remove("selected");

        btn.Classes.Add("selected");
        _selectedFacility = btn.Tag as string;
        tbStatus.Text = "";
    }

    // =========================================================================
    // 조회 버튼
    // =========================================================================
    private void BtnLoad_Click(object? sender, RoutedEventArgs e)
    {
        if (_selectedFacility == null)
        {
            tbStatus.Foreground = Brush.Parse("#ee8844");
            tbStatus.Text = "시설을 먼저 선택하세요";
            return;
        }
        if (dpDate.SelectedDate == null)
        {
            tbStatus.Foreground = Brush.Parse("#ee8844");
            tbStatus.Text = "날짜를 선택하세요";
            return;
        }

        string date = dpDate.SelectedDate!.Value.ToString("yyyy-MM-dd");

        try
        {
            _currentRows = FacilityResultService.GetRows(_selectedFacility, date);
            _cells.Clear();

            var grid = BuildResultGrid(_currentRows);
            ResultGridChanged?.Invoke(grid);

            tbStatus.Foreground = Brush.Parse("#888888");
            tbStatus.Text = $"{_selectedFacility} / {date}\n{_currentRows.Count}개 시료포인트";
        }
        catch (Exception ex)
        {
            tbStatus.Foreground = Brush.Parse("#ee4444");
            tbStatus.Text = $"조회 오류: {ex.Message}";
            Debug.WriteLine($"[ProcessingFacility] 조회 오류: {ex}");
        }
    }

    // =========================================================================
    // 저장 (MainPage BT1에서 호출)
    // =========================================================================
    public void Save()
    {
        if (_selectedFacility == null || dpDate.SelectedDate == null || _currentRows.Count == 0)
            return;

        string date = dpDate.SelectedDate!.Value.ToString("yyyy-MM-dd");

        // TextBox 값을 모델에 반영
        foreach (var r in _currentRows)
        {
            r.BOD      = GetCell(r.마스터Id, "BOD");
            r.TOC      = GetCell(r.마스터Id, "TOC");
            r.SS       = GetCell(r.마스터Id, "SS");
            r.TN       = GetCell(r.마스터Id, "TN");
            r.TP       = GetCell(r.마스터Id, "TP");
            r.총대장균군 = GetCell(r.마스터Id, "총대장균군");
            r.COD      = GetCell(r.마스터Id, "COD");
            r.염소이온  = GetCell(r.마스터Id, "염소이온");
            r.영양염류  = GetCell(r.마스터Id, "영양염류");
            r.함수율   = GetCell(r.마스터Id, "함수율");
            r.중금속   = GetCell(r.마스터Id, "중금속");
            r.비고     = GetCell(r.마스터Id, "비고");
        }

        try
        {
            FacilityResultService.SaveRows(_selectedFacility, date, _currentRows,
                MainPage.CurrentEmployeeId);
            tbStatus.Foreground = Brush.Parse("#88ee88");
            tbStatus.Text = $"저장 완료 ({DateTime.Now:HH:mm:ss})";

            // 저장 후 재조회하여 Id 갱신
            _currentRows = FacilityResultService.GetRows(_selectedFacility, date);
        }
        catch (Exception ex)
        {
            tbStatus.Foreground = Brush.Parse("#ee4444");
            tbStatus.Text = $"저장 오류: {ex.Message}";
            Debug.WriteLine($"[ProcessingFacility] 저장 오류: {ex}");
        }
    }

    private string GetCell(int masterId, string key)
    {
        return _cells.TryGetValue((masterId, key), out var tb) ? tb.Text ?? "" : "";
    }

    // =========================================================================
    // 결과 입력 그리드 빌드
    // =========================================================================
    private Control BuildResultGrid(List<FacilityResultRow> rows)
    {
        // 시료명(160) + Cols + 기준비고(100)
        var colWidths = new List<int> { 160 };
        foreach (var (_, _, w) in Cols) colWidths.Add(w);
        colWidths.Add(100);

        var colDefs = string.Join(",", colWidths);

        var outerGrid = new Grid();
        outerGrid.ColumnDefinitions = ColumnDefinitions.Parse(colDefs);

        // 헤더 + 데이터행 수만큼 RowDefinition
        for (int i = 0; i <= rows.Count; i++)
            outerGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        // ── 헤더 행 ─────────────────────────────────────────────────────
        AddHeaderCell(outerGrid, 0, 0, "시료명");
        for (int c = 0; c < Cols.Length; c++)
            AddHeaderCell(outerGrid, 0, c + 1, Cols[c].Header);
        AddHeaderCell(outerGrid, 0, Cols.Length + 1, "기준");

        // ── 데이터 행 ────────────────────────────────────────────────────
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            bool isEven = r % 2 == 0;
            var rowBg = (AppRes(isEven ? "GridRowBg" : "GridRowAltBg") as SolidColorBrush)?.Color
                        ?? (isEven ? Color.Parse("#35343c") : Color.Parse("#2d2d35"));

            // 시료명 셀
            var nameCell = MakeCell(row.시료명, isEven, readOnly: true);
            Grid.SetRow(nameCell, r + 1);
            Grid.SetColumn(nameCell, 0);
            outerGrid.Children.Add(nameCell);

            // 항목 셀
            for (int c = 0; c < Cols.Length; c++)
            {
                var (_, key, _) = Cols[c];
                bool active = IsActive(row, key);
                string value = GetValue(row, key);

                Control cell;
                if (active)
                {
                    var tb = new TextBox
                    {
                        Text             = value,
                        FontFamily       = Font,
                        FontSize         = 12,
                        Background       = new SolidColorBrush(rowBg),
                        Foreground       = Brush.Parse("#e8e8e8"),
                        BorderThickness  = new Thickness(0, 0, 1, 1),
                        BorderBrush      = Brush.Parse("#404050"),
                        Padding          = new Thickness(4, 2),
                        VerticalAlignment = VerticalAlignment.Stretch,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        MinHeight        = 26,
                    };
                    _cells[(row.마스터Id, key)] = tb;
                    cell = tb;
                }
                else
                {
                    cell = new Border
                    {
                        Background      = Brush.Parse("#222228"),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        BorderBrush     = Brush.Parse("#333340"),
                        MinHeight       = 26,
                    };
                }

                Grid.SetRow(cell, r + 1);
                Grid.SetColumn(cell, c + 1);
                outerGrid.Children.Add(cell);
            }

            // 기준비고 셀 (마스터에서)
            var noteCell = MakeCell(row.비고마스터, isEven, readOnly: true,
                fg: "#666666");
            Grid.SetRow(noteCell, r + 1);
            Grid.SetColumn(noteCell, Cols.Length + 1);
            outerGrid.Children.Add(noteCell);
        }

        return new ScrollViewer
        {
            Content = outerGrid,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // ── 헤더 셀 ──────────────────────────────────────────────────────────
    private static void AddHeaderCell(Grid grid, int row, int col, string text)
    {
        var cell = new Border
        {
            Background      = Brush.Parse("#2a2a32"),
            BorderThickness = new Thickness(0, 0, 1, 1),
            BorderBrush     = Brush.Parse("#404050"),
            Padding         = new Thickness(4, 4),
            Child = new TextBlock
            {
                Text                = text,
                FontFamily          = Font,
                FontSize            = 11,
                FontWeight          = FontWeight.SemiBold,
                Foreground          = Brush.Parse("#cccccc"),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
        Grid.SetRow(cell, row);
        Grid.SetColumn(cell, col);
        grid.Children.Add(cell);
    }

    // ── 읽기전용 텍스트 셀 ────────────────────────────────────────────────
    private static Border MakeCell(string text, bool even,
        bool readOnly = false, string fg = "#e0e0e0")
    {
        var bg = even ? "#35343c" : "#2d2d35";
        return new Border
        {
            Background      = Brush.Parse(bg),
            BorderThickness = new Thickness(0, 0, 1, 1),
            BorderBrush     = Brush.Parse("#404050"),
            Padding         = new Thickness(4, 4),
            MinHeight       = 26,
            Child = new TextBlock
            {
                Text                = text,
                FontFamily          = Font,
                FontSize            = 12,
                Foreground          = Brush.Parse(fg),
                VerticalAlignment   = VerticalAlignment.Center,
                TextWrapping        = TextWrapping.NoWrap,
            }
        };
    }

    // ── 활성 여부 ─────────────────────────────────────────────────────────
    private static bool IsActive(FacilityResultRow r, string key) => key switch
    {
        "BOD"    => r.BOD활성,
        "TOC"    => r.TOC활성,
        "SS"     => r.SS활성,
        "TN"     => r.TN활성,
        "TP"     => r.TP활성,
        "총대장균군" => r.총대장균군활성,
        "COD"    => r.COD활성,
        "염소이온" => r.염소이온활성,
        "영양염류" => r.영양염류활성,
        "함수율"  => r.함수율활성,
        "중금속"  => r.중금속활성,
        "비고"   => true,
        _        => false,
    };

    // ── 현재 값 ───────────────────────────────────────────────────────────
    private static string GetValue(FacilityResultRow r, string key) => key switch
    {
        "BOD"    => r.BOD,
        "TOC"    => r.TOC,
        "SS"     => r.SS,
        "TN"     => r.TN,
        "TP"     => r.TP,
        "총대장균군" => r.총대장균군,
        "COD"    => r.COD,
        "염소이온" => r.염소이온,
        "영양염류" => r.영양염류,
        "함수율"  => r.함수율,
        "중금속"  => r.중금속,
        "비고"   => r.비고,
        _        => "",
    };
}
