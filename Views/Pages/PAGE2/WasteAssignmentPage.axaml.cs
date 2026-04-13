using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;
using ETA.Services.Common;
using ETA.Services.SERVICE2;
using ETA.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE2;

public partial class WasteAssignmentPage : UserControl
{
    // ── 외부 이벤트 ──────────────────────────────────────────────────────
    public event Action<Control?>? ResultGridChanged;

    // ── 상태 ─────────────────────────────────────────────────────────────
    private List<FacilityWorkItem>  _facilityItems  = new();
    private List<WasteRequestItem>  _wasteItems     = new();
    private string _currentDate = DateTime.Today.ToString("yyyy-MM-dd");

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // 상태별 색
    private static readonly Color ColMidam    = Color.Parse("#dd3333"); // 미담 = 빨강
    private static readonly Color ColDameum   = Color.Parse("#dd8833"); // 담음 = 주황
    private static readonly Color ColComplete = Color.Parse("#3388dd"); // 완료 = 파랑

    public WasteAssignmentPage()
    {
        InitializeComponent();
        dpDate.SelectedDate = DateTime.Today;
    }

    // =========================================================================
    // 조회
    // =========================================================================
    private void BtnLoad_Click(object? sender, RoutedEventArgs e)
    {
        if (dpDate.SelectedDate == null)
        {
            SetStatus("날짜를 선택하세요", "#ee8844");
            return;
        }
        _currentDate = dpDate.SelectedDate!.Value.ToString("yyyy-MM-dd");
        Load();
    }

    private void Load()
    {
        try
        {
            _facilityItems = WasteRequestService.GetFacilityItems(_currentDate);
            _wasteItems    = WasteRequestService.GetItemsByDate(_currentDate);
            var grid = BuildAssignmentGrid();
            ResultGridChanged?.Invoke(grid);

            int total = _facilityItems.Count + _wasteItems.Count;
            SetStatus($"{_currentDate}\n처리시설 {_facilityItems.Count}건 / 폐수의뢰 {_wasteItems.Count}건", "#888888");
        }
        catch (Exception ex)
        {
            SetStatus($"조회 오류: {ex.Message}", "#ee4444");
        }
    }

    // =========================================================================
    // 의뢰 생성
    // =========================================================================
    private void CbGroup_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        string 구분 = GetSelectedGroup();
        if (string.IsNullOrEmpty(구분)) return;

        try
        {
            var companies = WasteRequestService.GetCompanies(구분);
            cbCompany.Items.Clear();
            foreach (var (업체명, 관리번호) in companies)
            {
                cbCompany.Items.Add(new ComboBoxItem
                {
                    Content = 업체명,
                    Tag     = 관리번호,
                });
            }
            if (cbCompany.Items.Count > 0) cbCompany.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            SetStatus($"업체 목록 오류: {ex.Message}", "#ee4444");
        }
    }

    private void BtnCreate_Click(object? sender, RoutedEventArgs e)
    {
        if (dpDate.SelectedDate == null)
        {
            SetStatus("날짜를 선택하세요", "#ee8844");
            return;
        }
        string 구분 = GetSelectedGroup();
        if (string.IsNullOrEmpty(구분))
        {
            SetStatus("구분을 선택하세요", "#ee8844");
            return;
        }
        if (cbCompany.SelectedItem is not ComboBoxItem cItem)
        {
            SetStatus("업체를 선택하세요", "#ee8844");
            return;
        }
        var 항목들 = GetSelectedAnalytes();
        if (항목들.Count == 0)
        {
            SetStatus("항목을 하나 이상 선택하세요", "#ee8844");
            return;
        }

        DateTime date   = dpDate.SelectedDate!.Value;
        string dateStr  = date.ToString("yyyy-MM-dd");
        string 업체명   = cItem.Content?.ToString() ?? "";
        string 관리번호 = cItem.Tag?.ToString() ?? "";

        try
        {
            string 의뢰번호 = WasteRequestService.GenerateRequestNumber(date, 구분, 관리번호);
            var req = new WasteRequest
            {
                의뢰번호 = 의뢰번호,
                구분     = 구분,
                채취일자 = dateStr,
                업체명   = 업체명,
                관리번호 = 관리번호,
                등록자   = MainPage.CurrentEmployeeId.ToString(),
            };
            WasteRequestService.CreateRequest(req, 항목들);
            SetStatus($"의뢰 생성 완료\n{의뢰번호}", "#88ee88");
            _currentDate = dateStr;
            Load();
        }
        catch (Exception ex)
        {
            SetStatus($"의뢰 생성 오류: {ex.Message}", "#ee4444");
        }
    }

    // =========================================================================
    // 결과 그리드 빌드
    // =========================================================================
    private Control BuildAssignmentGrid()
    {
        var outer = new StackPanel { Spacing = 0 };

        // ── 처리시설 기본 항목 섹션 ──────────────────────────────────────
        if (_facilityItems.Count > 0)
        {
            outer.Children.Add(BuildSectionHeader("처리시설 기본 항목"));
            outer.Children.Add(BuildFacilityHeader());

            for (int i = 0; i < _facilityItems.Count; i++)
            {
                var item = _facilityItems[i];
                outer.Children.Add(BuildFacilityRow(item, i));
            }
        }

        // ── 폐수배출업소 의뢰 섹션 ──────────────────────────────────────
        if (_wasteItems.Count > 0)
        {
            outer.Children.Add(BuildSectionHeader("폐수배출업소 의뢰"));
            outer.Children.Add(BuildWasteHeader());

            // 의뢰번호별 그룹
            var groups = _wasteItems.GroupBy(x => x.의뢰번호).ToList();
            int rowIdx = 0;
            foreach (var grp in groups)
            {
                foreach (var item in grp)
                {
                    outer.Children.Add(BuildWasteRow(item, rowIdx));
                    rowIdx++;
                }
            }
        }

        if (_facilityItems.Count == 0 && _wasteItems.Count == 0)
        {
            outer.Children.Add(new TextBlock
            {
                Text            = "해당 날짜의 작업 항목이 없습니다",
                FontFamily      = Font,
                Foreground      = AppTheme.FgDimmed,
                Margin          = new Thickness(12, 16),
                HorizontalAlignment = HorizontalAlignment.Center,
            }.BindMD());
        }

        return new ScrollViewer
        {
            Content = outer,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // ── 섹션 헤더 ────────────────────────────────────────────────────────
    private static Border BuildSectionHeader(string title)
    {
        return new Border
        {
            Background      = AppTheme.BgPrimary,
            BorderThickness = new Thickness(0, 0, 0, 1),
            BorderBrush     = AppTheme.BorderAccent,
            Padding         = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text = title, FontFamily = Font,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse("#8899cc"),
            }.BindBase()
        };
    }

    // ── 처리시설 헤더 행 ─────────────────────────────────────────────────
    private static Grid BuildFacilityHeader()
    {
        // 상태(40) | 시설명(120) | 시료명(160) | 항목(300) | 주기(110) | 배정자(80)
        var g = MakeRowGrid("40,120,160,300,110,80");
        AddHdr(g, 0, "상태");
        AddHdr(g, 1, "시설명");
        AddHdr(g, 2, "시료명");
        AddHdr(g, 3, "항목");
        AddHdr(g, 4, "주기");
        AddHdr(g, 5, "배정자");
        return g;
    }

    // ── 처리시설 데이터 행 ───────────────────────────────────────────────
    private Grid BuildFacilityRow(FacilityWorkItem item, int rowIdx)
    {
        bool even = rowIdx % 2 == 0;
        var g = MakeRowGrid("40,120,160,300,110,80");
        var bg = even ? "#35343c" : "#2d2d35";

        Button? statusBtn = null;
        statusBtn = MakeStatusButton(item.상태, () =>
        {
            string next = item.상태 == "미담" ? "담음" : item.상태 == "담음" ? "완료" : "미담";
            string user = MainPage.CurrentEmployeeId;
            if (next == "담음")        WasteRequestService.AssignFacilityItem(item.Id, user);
            else if (next == "완료")   WasteRequestService.CompleteFacilityItem(item.Id, user);
            else                        WasteRequestService.UnassignFacilityItem(item.Id);
            item.상태 = next;
            item.배정자 = next == "미담" ? "" : user;
            if (statusBtn != null) UpdateStatusButton(statusBtn, next);
        });

        Grid.SetColumn(statusBtn, 0);
        g.Children.Add(statusBtn);
        AddCell(g, 1, item.시설명, bg);
        AddCell(g, 2, item.시료명, bg);
        AddCell(g, 3, item.항목목록, bg);
        AddCell(g, 4, item.비고마스터, bg, fg: "#888888");
        AddCell(g, 5, item.배정자, bg, fg: "#aaaaaa");
        return g;
    }

    // ── 폐수의뢰 헤더 행 ─────────────────────────────────────────────────
    private static Grid BuildWasteHeader()
    {
        // 상태(40) | 의뢰번호(110) | 구분(50) | 업체명(160) | 항목(80) | 배정자(80)
        var g = MakeRowGrid("40,110,50,160,80,80");
        AddHdr(g, 0, "상태");
        AddHdr(g, 1, "의뢰번호");
        AddHdr(g, 2, "구분");
        AddHdr(g, 3, "업체명");
        AddHdr(g, 4, "항목");
        AddHdr(g, 5, "배정자");
        return g;
    }

    // ── 폐수의뢰 데이터 행 ──────────────────────────────────────────────
    private Grid BuildWasteRow(WasteRequestItem item, int rowIdx)
    {
        bool even = rowIdx % 2 == 0;
        var g = MakeRowGrid("40,110,50,160,80,80");
        var bg = even ? "#35343c" : "#2d2d35";

        Button? statusBtn = null;
        statusBtn = MakeStatusButton(item.상태, () =>
        {
            string next = item.상태 == "미담" ? "담음" : item.상태 == "담음" ? "완료" : "미담";
            string user = MainPage.CurrentEmployeeId;
            if (next == "담음")       WasteRequestService.AssignItem(item.Id, user);
            else if (next == "완료")  WasteRequestService.CompleteItem(item.Id, user);
            else                       WasteRequestService.UnassignItem(item.Id);
            item.상태 = next;
            item.배정자 = next == "미담" ? "" : user;
            if (statusBtn != null) UpdateStatusButton(statusBtn, next);
        });

        Grid.SetColumn(statusBtn, 0);
        g.Children.Add(statusBtn);
        AddCell(g, 1, item.의뢰번호, bg);
        AddCell(g, 2, item.구분, bg, fg: "#88aacc");
        AddCell(g, 3, item.업체명, bg);
        AddCell(g, 4, item.항목, bg);
        AddCell(g, 5, item.배정자, bg, fg: "#aaaaaa");
        return g;
    }

    // =========================================================================
    // 상태 버튼 헬퍼
    // =========================================================================
    private static Button MakeStatusButton(string 상태, Action onClick)
    {
        var btn = new Button
        {
            Background      = Brushes.Transparent,
            BorderThickness = new Thickness(0, 0, 1, 1),
            BorderBrush     = AppTheme.BorderMuted,
            MinHeight       = 28,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment   = VerticalAlignment.Center,
            Padding         = new Thickness(8, 0),
            Content         = MakeCircle(상태),
        };
        btn.Click += (_, _) => onClick();
        return btn;
    }

    private static Ellipse MakeCircle(string 상태)
    {
        var color = 상태 switch
        {
            "담음" => ColDameum,
            "완료" => ColComplete,
            _      => ColMidam,
        };
        return new Ellipse
        {
            Width  = 14,
            Height = 14,
            Fill   = new SolidColorBrush(color),
        };
    }

    private static void UpdateStatusButton(Button btn, string 상태)
    {
        btn.Content = MakeCircle(상태);
    }

    // =========================================================================
    // 그리드 / 셀 헬퍼
    // =========================================================================
    private static Grid MakeRowGrid(string colDefs)
    {
        var g = new Grid();
        g.ColumnDefinitions = ColumnDefinitions.Parse(colDefs);
        g.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        return g;
    }

    private static void AddHdr(Grid g, int col, string text)
    {
        var cell = new Border
        {
            Background      = AppTheme.BgCard,
            BorderThickness = new Thickness(0, 0, 1, 1),
            BorderBrush     = AppTheme.BorderMuted,
            Padding         = new Thickness(4, 4),
            Child = new TextBlock
            {
                Text = text, FontFamily = Font,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgSecondary,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }.BindBase()
        };
        Grid.SetColumn(cell, col);
        g.Children.Add(cell);
    }

    private static void AddCell(Grid g, int col, string text, string bg,
        string fg = "#e0e0e0")
    {
        var cell = new Border
        {
            Background      = Brush.Parse(bg),
            BorderThickness = new Thickness(0, 0, 1, 1),
            BorderBrush     = AppTheme.BorderMuted,
            Padding         = new Thickness(6, 4),
            MinHeight       = 28,
            Child = new TextBlock
            {
                Text = text, FontFamily = Font,
                Foreground = Brush.Parse(fg),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping      = TextWrapping.NoWrap,
            }.BindBase()
        };
        Grid.SetColumn(cell, col);
        g.Children.Add(cell);
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private string GetSelectedGroup()
    {
        if (cbGroup.SelectedItem is ComboBoxItem item)
            return item.Content?.ToString() ?? "";
        return "";
    }

    private List<string> GetSelectedAnalytes()
    {
        var list = new List<string>();
        if (chkBOD.IsChecked     == true) list.Add("BOD");
        if (chkTOC.IsChecked     == true) list.Add("TOC");
        if (chkSS.IsChecked      == true) list.Add("SS");
        if (chkTN.IsChecked      == true) list.Add("T-N");
        if (chkTP.IsChecked      == true) list.Add("T-P");
        if (chkPhenols.IsChecked == true) list.Add("Phenols");
        if (chkHexane.IsChecked  == true) list.Add("N-Hexane");
        if (chkCOD.IsChecked     == true) list.Add("COD");
        return list;
    }

    private void SetStatus(string msg, string color)
    {
        tbStatus.Foreground = Brush.Parse(color);
        tbStatus.Text       = msg;
    }
}
