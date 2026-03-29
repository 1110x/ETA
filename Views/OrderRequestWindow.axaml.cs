using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views;

/// <summary>
/// 의뢰서 시료명 선택 팝업
/// - 이동 가능, 최소/최대/닫기
/// - 체크박스 리스트로 다중 선택
/// - 하단 시료명 추가 기능
/// </summary>
public class OrderRequestWindow : Window
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    private readonly QuotationIssue  _issue;
    private readonly HashSet<string> _analysisItems;

    private string?      _matchedColumn          = null;
    private string?      _matchedMeasurerCompany = null;
    private StackPanel   _cbPanel       = new();
    private TextBox      _txbNew        = new();
    private TextBlock    _txbInfo       = new();
    private TextBlock    _txbStatus     = new();
    private Button       _btnNext       = new();
    private ToggleSwitch _tglMeasurer   = new();

    /// <summary>확인 후 넘길 선택된 시료명 목록</summary>
    public List<string> SelectedSamples { get; private set; } = new();
    public bool Confirmed { get; private set; } = false;

    public OrderRequestWindow(QuotationIssue issue, HashSet<string> analysisItems)
    {
        _issue         = issue;
        _analysisItems = analysisItems;

        Title                 = "의뢰서 작성 — 시료명 선택";
        Width                 = 440;
        Height                = 560;
        MinWidth              = 360;
        MinHeight             = 400;
        CanResize             = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        // 기본 시스템 데코레이션(이동+최소/최대/닫기) 유지
        SystemDecorations     = SystemDecorations.Full;
        Background            = Brush.Parse("#1a1a2a");

        Content = BuildUI();
        LoadSampleNames();
    }

    // ── UI ───────────────────────────────────────────────────────────────
    private Control BuildUI()
    {
        // 헤더 정보
        var headerInfo = new StackPanel
        {
            Spacing = 3, Margin = new Thickness(0, 0, 0, 8),
            Children =
            {
                new TextBlock
                {
                    Text = "📋  의뢰서 작성",
                    FontSize = 14, FontWeight = FontWeight.Bold,
                    FontFamily = Font, Foreground = Brush.Parse("#e0e0e0"),
                },
                new TextBlock
                {
                    Text = $"업체: {_issue.업체명}  |  견적번호: {_issue.견적번호}",
                    FontSize = 10, FontFamily = Font,
                    Foreground = Brush.Parse("#777799"),
                },
            }
        };

        // 측정인 토글
        _tglMeasurer = new ToggleSwitch
        {
            IsChecked   = true,
            OnContent   = "측정인 채취지점",
            OffContent  = "기존 목록",
            FontSize    = 10, FontFamily = Font,
            Foreground  = Brush.Parse("#aaaacc"),
            Margin      = new Thickness(0, 0, 0, 6),
        };
        _tglMeasurer.IsCheckedChanged += (_, _) => ReloadList();

        // 매칭 정보
        _txbInfo = new TextBlock
        {
            FontSize = 10, FontFamily = Font,
            Foreground = Brush.Parse("#88aacc"),
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
        };

        // 전체선택/해제 버튼
        var btnAll = new Button
        {
            Content = "전체 선택", Height = 22, FontSize = 10,
            FontFamily = Font, Background = Brush.Parse("#2a2a3a"),
            Foreground = Brush.Parse("#aaa"), BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(3), Padding = new Thickness(8, 0),
            Margin = new Thickness(0, 0, 6, 0),
        };
        var btnNone = new Button
        {
            Content = "전체 해제", Height = 22, FontSize = 10,
            FontFamily = Font, Background = Brush.Parse("#2a2a3a"),
            Foreground = Brush.Parse("#aaa"), BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(3), Padding = new Thickness(8, 0),
        };
        btnAll.Click  += (_, _) => SetAllChecked(true);
        btnNone.Click += (_, _) => SetAllChecked(false);

        var selBtns = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 0,
            Margin = new Thickness(0, 0, 0, 4),
            Children = { btnAll, btnNone }
        };

        // 체크박스 패널 (스크롤)
        _cbPanel = new StackPanel { Spacing = 0 };
        var scroll = new ScrollViewer
        {
            Content = _cbPanel,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Height = 260,
            Background = Brush.Parse("#13131f"),
        };

        // 시료명 추가
        var addLabel = new TextBlock
        {
            Text = "시료명 추가", FontSize = 10, FontFamily = Font,
            Foreground = Brush.Parse("#666"), Margin = new Thickness(0, 10, 0, 4),
        };
        _txbNew = new TextBox
        {
            Watermark = "새 시료명 입력", Height = 28, FontSize = 11,
            FontFamily = Font, Background = Brush.Parse("#1e1e2e"),
            Foreground = Brush.Parse("#ddd"), BorderBrush = Brush.Parse("#444"),
            Padding = new Thickness(8, 0),
        };
        var btnAdd = new Button
        {
            Content = "➕ 추가", Height = 28, FontSize = 11,
            FontFamily = Font, Background = Brush.Parse("#2a3a2a"),
            Foreground = Brush.Parse("#88cc88"), BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(3), Padding = new Thickness(12, 0),
        };
        btnAdd.Click += BtnAdd_Click;

        var addRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(0,0,0,4) };
        addRow.Children.Add(_txbNew);
        Grid.SetColumn(btnAdd, 1);
        addRow.Children.Add(btnAdd);

        // 상태
        _txbStatus = new TextBlock
        {
            FontSize = 10, FontFamily = Font,
            IsVisible = false, Margin = new Thickness(0, 2, 0, 0),
        };

        // 다음 버튼
        _btnNext = new Button
        {
            Content = "다음 →  편집 화면으로",
            Height = 34, FontSize = 12, FontFamily = Font,
            Background = Brush.Parse("#1a4a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            IsEnabled = false, Margin = new Thickness(0, 10, 0, 0),
        };
        _btnNext.Click += BtnNext_Click;

        return new Border
        {
            Padding = new Thickness(16),
            Child   = new StackPanel
            {
                Spacing = 0,
                Children =
                {
                    headerInfo,
                    new Border { Height = 1, Background = Brush.Parse("#333"), Margin = new Thickness(0,0,0,8) },
                    _tglMeasurer,
                    _txbInfo,
                    selBtns,
                    scroll,
                    addLabel,
                    addRow,
                    _txbStatus,
                    _btnNext,
                }
            }
        };
    }

    // ── 시료명 로드 ───────────────────────────────────────────────────────
    private void LoadSampleNames() => ReloadList();

    private void ReloadList()
    {
        if (_tglMeasurer.IsChecked == true)
            LoadMeasurerList();
        else
            LoadOrderRequestList();
    }

    private void LoadMeasurerList()
    {
        var points = FindMeasurerPoints(_issue.업체명);
        if (points.Count > 0)
        {
            _txbInfo.Text       = $"✅ 측정인 채취지점 {points.Count}개  ({_matchedMeasurerCompany})";
            _txbInfo.Foreground = Brush.Parse("#88cc88");
            LoadCheckboxes(points);
        }
        else
        {
            _txbInfo.Text       = $"⚠️ 측정인 DB에서 '{_issue.업체명}' 현장을 찾지 못했습니다.";
            _txbInfo.Foreground = Brush.Parse("#f0c040");
            _cbPanel.Children.Clear();
            UpdateNextButton();
        }
    }

    private void LoadOrderRequestList()
    {
        _matchedColumn = OrderRequestService.FindColumnByCompany(_issue.업체명);
        if (_matchedColumn == null)
        {
            _txbInfo.Text       = $"ℹ️ '{_issue.업체명}' 컬럼이 없습니다. 하단에서 시료명을 추가하면 자동 생성됩니다.";
            _txbInfo.Foreground = Brush.Parse("#aaaacc");
            _cbPanel.Children.Clear();
            UpdateNextButton();
            return;
        }

        bool exact = string.Equals(_matchedColumn.Trim(), _issue.업체명.Trim(),
                                   StringComparison.OrdinalIgnoreCase);
        _txbInfo.Text       = exact
            ? $"✅ 매칭: {_matchedColumn}"
            : $"🔍 유사매칭: {_matchedColumn}  (입력: {_issue.업체명})";
        _txbInfo.Foreground = exact ? Brush.Parse("#88cc88") : Brush.Parse("#f0c040");
        RefreshCheckboxes();
    }

    // 업체명 정규화: ㈜/(주)/㈔/(사) 통일, 공백·괄호·하이픈 제거, 소문자화
    private static string NormCompany(string name)
    {
        var s = name
            .Replace("㈜", "주").Replace("(주)", "주").Replace("（주）", "주")
            .Replace("㈔", "사").Replace("(사)", "사")
            .Replace("주식회사", "주").Replace("유한회사", "유")
            .Replace(" ", "").Replace("-", "").Replace("·", "")
            .Replace("(", "").Replace(")", "").Replace("（", "").Replace("）", "");
        return s.ToLower();
    }

    // 측정인 DB에서 업체명 퍼지 매칭 후 채취지점 반환
    private List<string> FindMeasurerPoints(string companyName)
    {
        var companies = MeasurerService.GetCompanies();

        // 1. 완전일치
        var match = companies.FirstOrDefault(c =>
            string.Equals(c.Trim(), companyName.Trim(), StringComparison.OrdinalIgnoreCase));

        // 2. 정규화 후 일치
        if (match == null)
        {
            var normQuery = NormCompany(companyName);
            match = companies.FirstOrDefault(c =>
                string.Equals(NormCompany(c), normQuery, StringComparison.OrdinalIgnoreCase));
        }

        // 3. 정규화 포함 관계
        if (match == null)
        {
            var normQuery = NormCompany(companyName);
            match = companies.FirstOrDefault(c =>
            {
                var nc = NormCompany(c);
                return nc.Contains(normQuery, StringComparison.OrdinalIgnoreCase) ||
                       normQuery.Contains(nc, StringComparison.OrdinalIgnoreCase);
            });
        }

        // 4. 원본 포함 관계 (기존 로직 유지)
        match ??= companies.FirstOrDefault(c =>
            c.Contains(companyName, StringComparison.OrdinalIgnoreCase) ||
            companyName.Contains(c, StringComparison.OrdinalIgnoreCase));

        if (match == null) return new List<string>();
        _matchedMeasurerCompany = match;
        return MeasurerService.GetSamplingPoints(match);
    }

    private void LoadCheckboxes(List<string> names)
    {
        _cbPanel.Children.Clear();
        bool odd = false;
        foreach (var name in names)
            _cbPanel.Children.Add(MakeCbRow(name, ref odd));
        UpdateNextButton();
    }

    private void RefreshCheckboxes()
    {
        if (_matchedColumn == null) return;
        _cbPanel.Children.Clear();

        var names = OrderRequestService.GetSampleNames(_matchedColumn);
        bool odd  = false;
        foreach (var name in names)
            _cbPanel.Children.Add(MakeCbRow(name, ref odd));

        UpdateNextButton();
    }

    private Border MakeCbRow(string name, ref bool odd)
    {
        var cb = new CheckBox
        {
            Content    = name,
            IsChecked  = false,
            FontSize   = 11, FontFamily = Font,
            Foreground = Brush.Parse("#cccccc"),
            Padding    = new Thickness(6, 0),
        };
        cb.IsCheckedChanged += (_, _) => UpdateNextButton();

        var row = new Border
        {
            Background      = Brush.Parse(odd ? "#16161e" : "#1a1a26"),
            Padding         = new Thickness(8, 5),
            Child           = cb,
        };
        odd = !odd;
        return row;
    }

    private void SetAllChecked(bool check)
    {
        foreach (var child in _cbPanel.Children.OfType<Border>())
            if (child.Child is CheckBox cb) cb.IsChecked = check;
        UpdateNextButton();
    }

    private void UpdateNextButton()
    {
        int cnt = _cbPanel.Children.OfType<Border>()
            .Count(b => b.Child is CheckBox cb && cb.IsChecked == true);
        _btnNext.IsEnabled  = cnt > 0;
        _btnNext.Content    = cnt > 0
            ? $"다음 →  편집 화면으로  ({cnt}건 선택)"
            : "다음 →  편집 화면으로";
    }

    // ── 시료명 추가 ───────────────────────────────────────────────────────
    private void BtnAdd_Click(object? sender, RoutedEventArgs e)
    {
        var name = _txbNew.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(name)) return;

        // 업체 컬럼이 없으면 자동 생성
        if (_matchedColumn == null)
        {
            var colName = _issue.업체명?.Trim() ?? "";
            if (string.IsNullOrEmpty(colName))
            {
                ShowStatus("⚠️ 업체명이 비어있어 컬럼을 생성할 수 없습니다.", error: true);
                return;
            }
            if (!OrderRequestService.CreateCompanyColumn(colName))
            {
                ShowStatus("❌ 업체 컬럼 생성 실패", error: true);
                return;
            }
            _matchedColumn = colName;
            _txbInfo.Text       = $"✅ '{colName}' 컬럼 자동 생성";
            _txbInfo.Foreground = Brush.Parse("#88cc88");
        }

        bool ok = OrderRequestService.AddSampleName(_matchedColumn, name);
        if (ok)
        {
            _txbNew.Text = "";
            RefreshCheckboxes();
            ShowStatus($"✅ '{name}' 추가 완료");
        }
        else ShowStatus("❌ 추가 실패", error: true);
    }

    // ── 다음 (편집 화면으로) ──────────────────────────────────────────────
    private void BtnNext_Click(object? sender, RoutedEventArgs e)
    {
        SelectedSamples = _cbPanel.Children.OfType<Border>()
            .Where(b => b.Child is CheckBox cb && cb.IsChecked == true)
            .Select(b => ((CheckBox)b.Child!).Content?.ToString() ?? "")
            .Where(s => !string.IsNullOrEmpty(s))
            .ToList();

        Confirmed = true;
        Close();
    }

    private void ShowStatus(string msg, bool error = false)
    {
        _txbStatus.Text       = msg;
        _txbStatus.Foreground = error ? Brush.Parse("#f08080") : Brush.Parse("#88ee88");
        _txbStatus.IsVisible  = true;
    }
}
