using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Services.SERVICE2;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 배출업소 시험성적서 페이지
/// — Show1: 날짜 트리뷰 (월 > 일)
/// — Show2: 선택 날짜의 결과 그리드 (업체명 | SN | BOD | TOC | SS | TN | TP | N-Hexan | Phenols)
/// — 추세 비교: 이전 3회 평균 대비 ▲/▼ 표시
/// </summary>
public class WasteTestReportPage : UserControl
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 외부 이벤트 ──────────────────────────────────────────────────────
    public event Action<Control?>? ResultGridChanged;

    // ── 상태 ─────────────────────────────────────────────────────────────
    private TreeView _dateTree = null!;
    private string? _selectedDate;
    private bool _showTrend = true;
    private TextBlock _statusText = null!;
    private CheckBox _chk여수 = null!, _chk율촌 = null!, _chk세풍 = null!;
    private CheckBox _trendCheck = null!;
    private NumericUpDown _nudUp = null!, _nudDown = null!, _nudCount = null!;

    private List<string> SelectedGroups
    {
        get
        {
            var g = new List<string>();
            if (_chk여수.IsChecked == true) g.Add("여수");
            if (_chk율촌.IsChecked == true) g.Add("율촌");
            if (_chk세풍.IsChecked == true) g.Add("세풍");
            return g;
        }
    }

    private string GroupLabel => SelectedGroups.Count == 0 ? "전체" : string.Join("+", SelectedGroups);

    // ── 생성자 ───────────────────────────────────────────────────────────
    public WasteTestReportPage()
    {
        Content = BuildDatePanel();
    }

    // =====================================================================
    // Show1: 날짜 트리뷰
    // =====================================================================
    private Control BuildDatePanel()
    {
        var root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,Auto,Auto,*,Auto") };

        // ── 헤더 ─────────────────────────────────────────────────────────
        var header = new TextBlock
        {
            Text = "배출업소 시험성적서",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = AppTheme.FgPrimary,
            Margin = new Thickness(8, 6),
        };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ── 구분 버튼 + 추세 체크 ────────────────────────────────────────
        var toolbar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 4, Margin = new Thickness(8, 2),
        };

        _chk여수 = MakeGroupCheck("여수", true);
        _chk율촌 = MakeGroupCheck("율촌", false);
        _chk세풍 = MakeGroupCheck("세풍", false);
        toolbar.Children.Add(_chk여수);
        toolbar.Children.Add(_chk율촌);
        toolbar.Children.Add(_chk세풍);

        toolbar.Children.Add(new Border { Width = 12 }); // spacer

        _trendCheck = new CheckBox
        {
            Content = "추세 비교",
            IsChecked = true,
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
        };
        _trendCheck.IsCheckedChanged += (_, _) =>
        {
            _showTrend = _trendCheck.IsChecked == true;
            if (_selectedDate != null) ShowResults(_selectedDate);
        };
        toolbar.Children.Add(_trendCheck);

        Grid.SetRow(toolbar, 1);
        root.Children.Add(toolbar);

        // ── 추세 설정 행 ─────────────────────────────────────────────────
        var trendRow = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(8, 2),
        };
        trendRow.Children.Add(MakeLabel("상승"));
        _nudUp = MakeNud(50, 0, 500, 10);
        trendRow.Children.Add(_nudUp);
        trendRow.Children.Add(MakeLabel("%"));

        trendRow.Children.Add(new Border { Width = 8 });
        trendRow.Children.Add(MakeLabel("하락"));
        _nudDown = MakeNud(50, 0, 100, 10);
        trendRow.Children.Add(_nudDown);
        trendRow.Children.Add(MakeLabel("%"));

        trendRow.Children.Add(new Border { Width = 8 });
        trendRow.Children.Add(MakeLabel("표본"));
        _nudCount = MakeNud(10, 1, 50, 1);
        trendRow.Children.Add(_nudCount);
        trendRow.Children.Add(MakeLabel("회"));

        Grid.SetRow(trendRow, 2);
        root.Children.Add(trendRow);

        // ── 트리뷰 ──────────────────────────────────────────────────────
        _dateTree = new TreeView
        {
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Margin = new Thickness(4, 4),
        };
        _dateTree.SelectionChanged += DateTree_SelectionChanged;

        var scroll = new ScrollViewer { Content = _dateTree };
        Grid.SetRow(scroll, 3);
        root.Children.Add(scroll);

        // ── 하단 상태 ────────────────────────────────────────────────────
        _statusText = new TextBlock
        {
            Text = "", FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgMuted, Margin = new Thickness(8, 4),
        };
        Grid.SetRow(_statusText, 4);
        root.Children.Add(_statusText);

        return root;
    }

    private static TextBlock MakeLabel(string text) => new()
    {
        Text = text, FontSize = AppTheme.FontSM, FontFamily = Font,
        Foreground = AppTheme.FgMuted,
        VerticalAlignment = VerticalAlignment.Center,
        Margin = new Thickness(2, 0),
    };

    private NumericUpDown MakeNud(double value, double min, double max, double inc)
    {
        var nud = new NumericUpDown
        {
            Value = (decimal)value, Minimum = (decimal)min, Maximum = (decimal)max,
            Increment = (decimal)inc,
            Width = 70, Height = 24,
            FontSize = AppTheme.FontSM, FontFamily = Font,
            FormatString = "F0",
            VerticalAlignment = VerticalAlignment.Center,
        };
        nud.ValueChanged += (_, _) =>
        {
            if (_selectedDate != null) ShowResults(_selectedDate);
        };
        return nud;
    }

    private CheckBox MakeGroupCheck(string label, bool isChecked)
    {
        var chk = new CheckBox
        {
            Content = label,
            IsChecked = isChecked,
            FontSize = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 4, 0),
        };
        chk.IsCheckedChanged += (_, _) => LoadData();
        return chk;
    }

    // ── 데이터 로드 ─────────────────────────────────────────────────────
    public void LoadData()
    {
        _dateTree.Items.Clear();
        _selectedDate = null;
        ResultGridChanged?.Invoke(null);

        try
        {
            var groups = SelectedGroups;
            var dates = WasteTestReportService.GetDates(groups);
            _statusText.Text = $"{GroupLabel} {dates.Count}건";

            // 월별 그룹핑
            var byMonth = new Dictionary<string, List<string>>();
            foreach (var d in dates)
            {
                string m = d.Length >= 7 ? d[..7] : d;
                if (!byMonth.ContainsKey(m)) byMonth[m] = new();
                byMonth[m].Add(d);
            }

            foreach (var (month, dayList) in byMonth)
            {
                var monthNode = new TreeViewItem
                {
                    Header = FormatMonth(month),
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppTheme.FgSecondary,
                    Tag = month,
                };

                foreach (var day in dayList)
                {
                    var dayNode = new TreeViewItem
                    {
                        Header = FormatDay(day),
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = AppTheme.FgPrimary,
                        Tag = day,
                    };
                    monthNode.Items.Add(dayNode);
                }

                // 최근 월은 펼침
                if (_dateTree.Items.Count == 0)
                    monthNode.IsExpanded = true;

                _dateTree.Items.Add(monthNode);
            }
        }
        catch (Exception ex)
        {
            _statusText.Text = $"오류: {ex.Message}";
        }
    }

    private void DateTree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_dateTree.SelectedItem is not TreeViewItem item) return;
        if (item.Tag is not string date || date.Length < 10) return; // 월 노드 무시
        _selectedDate = date;
        ShowResults(date);
    }

    // =====================================================================
    // Show2: 결과 그리드
    // =====================================================================
    private void ShowResults(string date)
    {
        try
        {
            var groups = SelectedGroups;
            var rows = WasteTestReportService.GetByDate(date, groups);
            Dictionary<string, Dictionary<string, string>>? trends = null;
            if (_showTrend && rows.Count > 0)
            {
                double upPct   = (double)(_nudUp.Value ?? 50m) / 100.0;
                double downPct = 1.0 - (double)(_nudDown.Value ?? 50m) / 100.0;
                int count      = (int)(_nudCount.Value ?? 10m);
                trends = WasteTestReportService.GetTrends(date, groups, count, upPct, downPct);
            }

            var grid = BuildResultGrid(date, rows, trends);
            ResultGridChanged?.Invoke(grid);
            _statusText.Text = $"{GroupLabel} {date} — {rows.Count}건";
        }
        catch (Exception ex)
        {
            _statusText.Text = $"오류: {ex.Message}";
        }
    }

    private Control BuildResultGrid(string date, List<WasteTestRow> rows,
        Dictionary<string, Dictionary<string, string>>? trends)
    {
        var root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,*") };

        // ── 헤더 바 ─────────────────────────────────────────────────────
        var headerBar = new Border
        {
            Background = AppTheme.BgSecondary,
            Padding = new Thickness(8, 6),
            Child = new TextBlock
            {
                Text = $"{GroupLabel} 배출업소 시험성적서  [{date}]  ({rows.Count}건)",
                FontSize = AppTheme.FontMD, FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppTheme.FgPrimary,
            },
        };
        Grid.SetRow(headerBar, 0);
        root.Children.Add(headerBar);

        // ── 데이터 그리드 ────────────────────────────────────────────────
        bool multiGroup = SelectedGroups.Count != 1;
        var colList = new List<(string Header, string Key, double Width)>
        {
            ("No",       "No",      35),
        };
        if (multiGroup) colList.Add(("구분", "구분", 45));
        colList.AddRange(new (string, string, double)[]
        {
            ("업체명",   "업체명",  200),
            ("S/N",      "SN",       60),
            ("BOD",      "BOD",      70),
            ("TOC",      "TOC",      70),
            ("SS",       "SS",       70),
            ("T-N",      "TN",       70),
            ("T-P",      "TP",       70),
            ("N-Hexan",  "NHexan",   70),
            ("Phenols",  "Phenols",  70),
            ("비고",     "비고",    100),
        });
        var cols = colList.ToArray();

        var dataGrid = new Grid
        {
            RowDefinitions = RowDefinitions.Parse(
                "Auto," + string.Join(",", Enumerable.Repeat("Auto", rows.Count))),
            ColumnDefinitions = new ColumnDefinitions(
                string.Join(",", cols.Select(c => $"{c.Width}"))),
        };

        // ── 헤더 행 ─────────────────────────────────────────────────────
        for (int ci = 0; ci < cols.Length; ci++)
        {
            var hdr = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#2a3a5a")),
                BorderBrush = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(4, 4),
                Child = new TextBlock
                {
                    Text = cols[ci].Header,
                    FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
                    FontFamily = Font, Foreground = AppTheme.FgPrimary,
                    HorizontalAlignment = HorizontalAlignment.Center,
                },
            };
            Grid.SetRow(hdr, 0);
            Grid.SetColumn(hdr, ci);
            dataGrid.Children.Add(hdr);
        }

        // ── 데이터 행 ────────────────────────────────────────────────────
        for (int ri = 0; ri < rows.Count; ri++)
        {
            var row = rows[ri];
            Dictionary<string, string>? rowTrend = null;
            trends?.TryGetValue(row.업체명, out rowTrend);

            var bgColor = ri % 2 == 0 ? "#1e1e2a" : "#22222e";

            for (int ci = 0; ci < cols.Length; ci++)
            {
                var (_, key, _) = cols[ci];
                string text = key switch
                {
                    "No"    => (ri + 1).ToString(),
                    "구분"   => row.구분,
                    "업체명" => row.업체명,
                    "SN"     => row.SN,
                    "BOD"    => FormatValue(row.BOD, "BOD"),
                    "TOC"    => FormatValue(row.TOC, "TOC"),
                    "SS"     => FormatValue(row.SS, "SS"),
                    "TN"     => FormatValue(row.TN, "TN"),
                    "TP"     => FormatValue(row.TP, "TP"),
                    "NHexan" => FormatValue(row.NHexan, "NHexan"),
                    "Phenols"=> FormatValue(row.Phenols, "Phenols"),
                    "비고"   => row.비고,
                    _ => "",
                };

                // 추세 적용
                IBrush fg = AppTheme.FgPrimary;
                if (rowTrend != null && key is "BOD" or "TOC" or "SS" or "TN" or "TP" or "NHexan" or "Phenols")
                {
                    if (rowTrend.TryGetValue(key, out var trendText))
                    {
                        text = trendText;
                        fg = trendText.StartsWith("▲")
                            ? new SolidColorBrush(Color.Parse("#ff6666"))
                            : new SolidColorBrush(Color.Parse("#6699ff"));
                    }
                }

                var cell = new Border
                {
                    Background = new SolidColorBrush(Color.Parse(bgColor)),
                    BorderBrush = AppTheme.BorderSubtle,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(4, 3),
                    Child = new TextBlock
                    {
                        Text = text,
                        FontSize = AppTheme.FontSM,
                        FontFamily = Font,
                        Foreground = fg,
                        HorizontalAlignment = key is "업체명" or "비고"
                            ? HorizontalAlignment.Left
                            : HorizontalAlignment.Center,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    },
                };
                Grid.SetRow(cell, ri + 1);
                Grid.SetColumn(cell, ci);
                dataGrid.Children.Add(cell);
            }
        }

        var scroll = new ScrollViewer
        {
            Content = dataGrid,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        return root;
    }

    // ── 유틸 ─────────────────────────────────────────────────────────────
    /// <summary>항목별 자릿수 포맷 적용 (BOD/TOC/SS/N-Hexan: F1, TN/TP/Phenols: F3)</summary>
    private static string FormatValue(string raw, string item)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        if (!double.TryParse(raw, out var v)) return raw; // 숫자 아니면 원본 반환
        return v.ToString(WasteTestReportService.GetFormat(item));
    }

    private static string FormatMonth(string m) =>
        m.Length >= 7 && int.TryParse(m[5..7], out var mm) ? $"{m[..4]}년 {mm}월" : m;

    private static string FormatDay(string d)
    {
        if (d.Length >= 10 && DateTime.TryParse(d, out var dt))
        {
            var dow = dt.DayOfWeek switch
            {
                DayOfWeek.Monday    => "월",
                DayOfWeek.Tuesday   => "화",
                DayOfWeek.Wednesday => "수",
                DayOfWeek.Thursday  => "목",
                DayOfWeek.Friday    => "금",
                DayOfWeek.Saturday  => "토",
                DayOfWeek.Sunday    => "일",
                _ => "",
            };
            return $"{dt:MM/dd} ({dow})";
        }
        return d;
    }
}
