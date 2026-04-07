using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using ETA.Services.Common;
using ETA.Services.SERVICE2;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 배출업소 시험성적서 페이지
/// — Show1: 날짜 트리뷰 (노드에 구분별 건수 표시)
/// — Show2: 고정 레이아웃 (헤더 + 추세설정 + 데이터), 노드 변경 시 데이터만 갱신
/// — 추세: ▲(30%+) ▲▲(70%+) ↑(100%+) 단계별 표시
/// — 행 호버 하이라이트
/// </summary>
public class WasteTestReportPage : UserControl
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly IBrush HoverBg = new SolidColorBrush(Color.Parse("#333355"));

    // ── 외부 이벤트 ──────────────────────────────────────────────────────
    public event Action<Control?>? ResultGridChanged;

    // ── Show1 상태 ───────────────────────────────────────────────────────
    private TreeView _dateTree = null!;
    private string? _selectedDate;
    private bool _showTrend = true;
    private TextBlock _statusText = null!;
    private Button _btn여수 = null!, _btn율촌 = null!, _btn세풍 = null!;
    private Button _btnTrend = null!;

    // ── Show2 고정 레이아웃 ──────────────────────────────────────────────
    private Grid? _show2Root;
    private TextBlock _headerText = null!;
    private ScrollViewer _dataScroll = null!;
    private TextBox _txtT1 = null!, _txtT2 = null!, _txtT3 = null!, _txtCount = null!;

    // ── 캐시 ────────────────────────────────────────────────────────────
    private Dictionary<string, Dictionary<string, int>>? _dateGroupCounts;
    private Dictionary<string, string>? _companyAbbr;  // 업체명→약칭

    private static readonly IBrush ToggleOnBg  = new SolidColorBrush(Color.Parse("#2e4a6e"));
    private static readonly IBrush ToggleOffBg = new SolidColorBrush(Color.Parse("#2a2a3a"));

    private static bool IsToggleOn(Button btn) => btn.Tag is true;

    private List<string> SelectedGroups
    {
        get
        {
            var g = new List<string>();
            if (IsToggleOn(_btn여수)) g.Add("여수");
            if (IsToggleOn(_btn율촌)) g.Add("율촌");
            if (IsToggleOn(_btn세풍)) g.Add("세풍");
            return g;
        }
    }

    private string GroupLabel => SelectedGroups.Count == 0 ? "전체" : string.Join("+", SelectedGroups);

    // ── 생성자 ───────────────────────────────────────────────────────────
    public WasteTestReportPage()
    {
        InitTrendControls();
        Content = BuildDatePanel();
    }

    private void InitTrendControls()
    {
        _txtT1    = MakeTrendTextBox("30");
        _txtT2    = MakeTrendTextBox("70");
        _txtT3    = MakeTrendTextBox("100");
        _txtCount = MakeTrendTextBox("10");
    }

    private TextBox MakeTrendTextBox(string defaultValue)
    {
        var tb = new TextBox
        {
            Text = defaultValue,
            Width = 50, Height = 26,
            FontSize = AppTheme.FontMD, FontFamily = Font,
            TextAlignment = TextAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            Padding = new Thickness(2, 0),
        };
        tb.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter && _selectedDate != null) ShowResults(_selectedDate);
        };
        tb.LostFocus += (_, _) =>
        {
            if (_selectedDate != null) ShowResults(_selectedDate);
        };
        return tb;
    }

    private static double GetTextValue(TextBox tb, double fallback)
        => double.TryParse(tb.Text?.Trim(), out var v) ? v : fallback;

    // =====================================================================
    // Show1: 날짜 트리뷰
    // =====================================================================
    private Control BuildDatePanel()
    {
        var root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,Auto,*,Auto") };

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

        // ── 구분 토글 + 추세 토글 ─────────────────────────────────────────
        var toolbar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 4, Margin = new Thickness(8, 4),
        };

        _btn여수 = MakeGroupToggle("여수", true);
        _btn율촌 = MakeGroupToggle("율촌", true);
        _btn세풍 = MakeGroupToggle("세풍", true);
        toolbar.Children.Add(_btn여수);
        toolbar.Children.Add(_btn율촌);
        toolbar.Children.Add(_btn세풍);

        toolbar.Children.Add(new Border { Width = 8 });

        _btnTrend = MakeToggleButton("추세 비교", true);
        _btnTrend.Click += (_, _) =>
        {
            _btnTrend.Tag = !IsToggleOn(_btnTrend);
            ApplyToggleStyle(_btnTrend);
            _showTrend = IsToggleOn(_btnTrend);
            if (_selectedDate != null) ShowResults(_selectedDate);
        };
        toolbar.Children.Add(_btnTrend);

        Grid.SetRow(toolbar, 1);
        root.Children.Add(toolbar);

        // ── 트리뷰 ──────────────────────────────────────────────────────
        _dateTree = new TreeView
        {
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Margin = new Thickness(4, 4),
        };
        _dateTree.SelectionChanged += DateTree_SelectionChanged;

        var scroll = new ScrollViewer { Content = _dateTree };
        Grid.SetRow(scroll, 2);
        root.Children.Add(scroll);

        // ── 하단 상태 ────────────────────────────────────────────────────
        _statusText = new TextBlock
        {
            Text = "", FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgMuted, Margin = new Thickness(8, 4),
        };
        Grid.SetRow(_statusText, 3);
        root.Children.Add(_statusText);

        return root;
    }

    private static TextBlock MakeLabel(string text) => new()
    {
        Text = text, FontSize = AppTheme.FontMD, FontFamily = Font,
        Foreground = AppTheme.FgMuted,
        VerticalAlignment = VerticalAlignment.Center,
        Margin = new Thickness(2, 0),
    };

    private Button MakeToggleButton(string label, bool isOn)
    {
        var btn = new Button
        {
            Content = label,
            MinWidth = 40, Height = 26,
            FontSize = AppTheme.FontSM, FontFamily = Font,
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(4),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            Padding = new Thickness(8, 0),
            Tag = isOn,
        };
        ApplyToggleStyle(btn);
        return btn;
    }

    private Button MakeGroupToggle(string label, bool isOn)
    {
        var btn = MakeToggleButton(label, isOn);
        btn.Click += (_, _) =>
        {
            btn.Tag = !IsToggleOn(btn);
            ApplyToggleStyle(btn);
            LoadData();
        };
        return btn;
    }

    private static void ApplyToggleStyle(Button btn)
    {
        bool on = IsToggleOn(btn);
        btn.Background = on ? ToggleOnBg : ToggleOffBg;
        btn.Foreground = on ? AppTheme.FgPrimary : AppTheme.FgMuted;
    }

    // =====================================================================
    // Show2: 고정 레이아웃 (한번만 생성, 이후 데이터만 교체)
    // =====================================================================
    private void EnsureShow2()
    {
        if (_show2Root == null)
        {
            _show2Root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,Auto,*") };

            // ── Row 0: 헤더 바 ───────────────────────────────────────────
            _headerText = new TextBlock
            {
                Text = "",
                FontSize = AppTheme.FontMD, FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppTheme.FgPrimary,
            };
            var headerBar = new Border
            {
                Background = AppTheme.BgSecondary,
                Padding = new Thickness(8, 6),
                Child = _headerText,
            };
            Grid.SetRow(headerBar, 0);
            _show2Root.Children.Add(headerBar);

            // ── Row 1: 추세 설정 행 ──────────────────────────────────────
            var trendRow = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(8, 4),
            };
            trendRow.Children.Add(MakeLabel("단계"));
            trendRow.Children.Add(_txtT1);
            trendRow.Children.Add(MakeLabel("/"));
            trendRow.Children.Add(_txtT2);
            trendRow.Children.Add(MakeLabel("/"));
            trendRow.Children.Add(_txtT3);
            trendRow.Children.Add(MakeLabel("%"));
            trendRow.Children.Add(new Border { Width = 8 });
            trendRow.Children.Add(MakeLabel("표본"));
            trendRow.Children.Add(_txtCount);
            trendRow.Children.Add(MakeLabel("회"));

            Grid.SetRow(trendRow, 1);
            _show2Root.Children.Add(trendRow);

            // ── Row 2: 데이터 스크롤 영역 ────────────────────────────────
            _dataScroll = new ScrollViewer
            {
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            };
            Grid.SetRow(_dataScroll, 2);
            _show2Root.Children.Add(_dataScroll);
        }

        // 부모에서 떨어졌으면 다시 연결
        if (_show2Root.Parent == null)
            ResultGridChanged?.Invoke(_show2Root);
    }

    // ── 데이터 로드 ─────────────────────────────────────────────────────
    public void LoadData()
    {
        _dateTree.Items.Clear();
        _selectedDate = null;

        if (_show2Root == null)
            ResultGridChanged?.Invoke(null);
        else
        {
            _headerText.Text = "";
            _dataScroll.Content = null;
        }

        try
        {
            // 캐시 로드
            _dateGroupCounts = WasteTestReportService.GetDateGroupCounts();
            _companyAbbr ??= WasteTestReportService.GetCompanyAbbreviations();

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
                        Header = FormatDayWithCounts(day),
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = AppTheme.FgPrimary,
                        Tag = day,
                    };
                    monthNode.Items.Add(dayNode);
                }

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
        if (item.Tag is not string date || date.Length < 10) return;
        _selectedDate = date;
        ShowResults(date);
    }

    // =====================================================================
    // Show2: 데이터 갱신 (레이아웃 유지, 데이터만 교체)
    // =====================================================================
    private void ShowResults(string date)
    {
        try
        {
            EnsureShow2();

            var groups = SelectedGroups;
            var rows = WasteTestReportService.GetByDate(date, groups);
            Dictionary<string, Dictionary<string, TrendInfo>>? trends = null;
            if (_showTrend && rows.Count > 0)
            {
                int count = (int)GetTextValue(_txtCount, 10);
                trends = WasteTestReportService.GetTrends(date, groups, count);
            }

            _headerText.Text = $"{GroupLabel} 배출업소 시험성적서  [{date}]  ({rows.Count}건)";
            _dataScroll.Content = BuildDataGrid(rows, trends);
            _dataScroll.Offset = default; // 스크롤 맨 위로
            _statusText.Text = $"{GroupLabel} {date} — {rows.Count}건";
        }
        catch (Exception ex)
        {
            _statusText.Text = $"오류: {ex.Message}";
        }
    }

    private Control BuildDataGrid(List<WasteTestRow> rows,
        Dictionary<string, Dictionary<string, TrendInfo>>? trends)
    {
        double t1 = GetTextValue(_txtT1, 30);
        double t2 = GetTextValue(_txtT2, 70);
        double t3 = GetTextValue(_txtT3, 100);

        bool multiGroup = SelectedGroups.Count != 1;
        var colList = new List<(string Header, string Key, double Width)>
        {
            ("No", "No", 30),
        };
        if (multiGroup) colList.Add(("구분", "구분", 45));
        colList.AddRange(new (string, string, double)[]
        {
            ("업체명",   "업체명",  250),
            ("S/N",      "SN",       150),
            ("BOD",      "BOD",     130),
            ("TOC",      "TOC",     130),
            ("SS",       "SS",      130),
            ("T-N",      "TN",      130),
            ("T-P",      "TP",      130),
            ("N-Hexan",  "NHexan",  130),
            ("Phenols",  "Phenols", 130),
        });
        var cols = colList.ToArray();

        var colDef = string.Join(",", cols.Select(c => $"{c.Width}"));

        // ── 전체를 StackPanel(행 단위)로 구성 → 호버 하이라이트 가능
        var stack = new StackPanel { Orientation = Orientation.Vertical };

        // ── 헤더 행 ─────────────────────────────────────────────────────
        var headerRow = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(colDef),
        };
        for (int ci = 0; ci < cols.Length; ci++)
        {
            var hdr = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#2a3a5a")),
                BorderBrush = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(4, 5),
                Child = new TextBlock
                {
                    Text = cols[ci].Header,
                    FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
                    FontFamily = Font, Foreground = AppTheme.FgPrimary,
                    HorizontalAlignment = HorizontalAlignment.Center,
                },
            };
            Grid.SetColumn(hdr, ci);
            headerRow.Children.Add(hdr);
        }
        stack.Children.Add(headerRow);

        // ── 데이터 행 ────────────────────────────────────────────────────
        for (int ri = 0; ri < rows.Count; ri++)
        {
            var row = rows[ri];
            Dictionary<string, TrendInfo>? rowTrend = null;
            trends?.TryGetValue(row.업체명, out rowTrend);

            var bgColor = ri % 2 == 0 ? "#1e1e2a" : "#22222e";
            var bgBrush = new SolidColorBrush(Color.Parse(bgColor));

            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDef),
                Background = bgBrush,
            };

            for (int ci = 0; ci < cols.Length; ci++)
            {
                var (_, key, _) = cols[ci];

                Control cellContent;

                if (key == "업체명")
                {
                    cellContent = BuildCompanyCell(row.업체명);
                }
                else if (key is "BOD" or "TOC" or "SS" or "TN" or "TP" or "NHexan" or "Phenols")
                {
                    string rawVal = key switch
                    {
                        "BOD"     => row.BOD,
                        "TOC"     => row.TOC,
                        "SS"      => row.SS,
                        "TN"      => row.TN,
                        "TP"      => row.TP,
                        "NHexan"  => row.NHexan,
                        "Phenols" => row.Phenols,
                        _ => "",
                    };

                    if (rowTrend != null && rowTrend.TryGetValue(key, out var ti))
                        cellContent = BuildTrendCell(ti, key, t1, t2, t3);
                    else
                        cellContent = new TextBlock
                        {
                            Text = FormatValue(rawVal, key),
                            FontSize = AppTheme.FontMD, FontFamily = Font,
                            Foreground = AppTheme.FgPrimary,
                            HorizontalAlignment = HorizontalAlignment.Center,
                        };
                }
                else
                {
                    string text = key switch
                    {
                        "No"   => (ri + 1).ToString(),
                        "구분"  => row.구분,
                        "SN"    => row.SN,
                        "비고"  => row.비고,
                        _ => "",
                    };
                    cellContent = new TextBlock
                    {
                        Text = text,
                        FontSize = AppTheme.FontMD, FontFamily = Font,
                        Foreground = AppTheme.FgPrimary,
                        HorizontalAlignment = key == "비고"
                            ? HorizontalAlignment.Left : HorizontalAlignment.Center,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    };
                }

                var cell = new Border
                {
                    BorderBrush = AppTheme.BorderSubtle,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(4, 4),
                    Child = cellContent,
                };
                Grid.SetColumn(cell, ci);
                rowGrid.Children.Add(cell);
            }

            // ── 호버 하이라이트 ──────────────────────────────────────────
            var savedBg = bgBrush;
            rowGrid.PointerEntered += (_, _) => rowGrid.Background = HoverBg;
            rowGrid.PointerExited  += (_, _) => rowGrid.Background = savedBg;

            stack.Children.Add(rowGrid);
        }

        return stack;
    }

    // ── 업체명 셀 (약칭 배지 + 업체명) ─────────────────────────────────
    private Control BuildCompanyCell(string 업체명)
    {
        string? abbr = null;
        _companyAbbr?.TryGetValue(업체명, out abbr);

        var nameBlock = new TextBlock
        {
            Text = 업체명,
            FontSize = AppTheme.FontMD, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        if (string.IsNullOrEmpty(abbr))
            return nameBlock;

        var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);
        var badge = new Border
        {
            Background = new SolidColorBrush(Color.Parse(bg)),
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1),
            Margin = new Thickness(0, 0, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = abbr,
                FontSize = AppTheme.FontXS, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse(fg)),
            },
        };

        return new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Children = { badge, nameBlock },
        };
    }

    // ── 추세 셀 (3단계: ▲ / 포개진▲▲ / ⬆) ─────────────────────────────
    private static Control BuildTrendCell(TrendInfo ti, string key,
        double t1, double t2, double t3)
    {
        string fmt = WasteTestReportService.GetFormat(key);
        double absPct = Math.Abs(ti.PctChange);
        bool isUp = ti.PctChange > 0;

        // 변화 미미 → 값만 표시
        if (absPct < t1)
            return new TextBlock
            {
                Text = ti.Value.ToString(fmt),
                FontSize = AppTheme.FontMD, FontFamily = Font,
                Foreground = AppTheme.FgPrimary,
                HorizontalAlignment = HorizontalAlignment.Center,
            };

        int tier = absPct >= t3 ? 3 : absPct >= t2 ? 2 : 1;
        IBrush color = GetTrendBrush(ti.PctChange, t1, t2, t3);
        string valueText = $"{ti.Value.ToString(fmt)}({ti.Average.ToString(fmt)})";

        // 심볼 컨트롤
        Control symbol;
        if (tier == 2)
        {
            // 포개진 삼각형 (주식 스타일)
            symbol = BuildOverlapTriangles(isUp, color);
        }
        else if (tier == 3)
        {
            // 굵은 큰 화살표
            symbol = new TextBlock
            {
                Text = isUp ? "⬆" : "⬇",
                FontSize = AppTheme.FontMD + 4,
                Foreground = color,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, -2, 2, -2),
            };
        }
        else
        {
            // 빈 삼각형 (가벼운 느낌)
            symbol = new TextBlock
            {
                Text = isUp ? "△" : "▽",
                FontSize = AppTheme.FontMD - 1, FontFamily = Font,
                Foreground = color,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 1, 0),
            };
        }

        return new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Children =
            {
                symbol,
                new TextBlock
                {
                    Text = valueText,
                    FontSize = AppTheme.FontMD, FontFamily = Font,
                    Foreground = color,
                    VerticalAlignment = VerticalAlignment.Center,
                },
            },
        };
    }

    /// <summary>포개진 삼각형 (두 개가 수직으로 겹침, 주식 스타일)</summary>
    private static Control BuildOverlapTriangles(bool isUp, IBrush color)
    {
        string ch = isUp ? "▲" : "▼";
        double sz = AppTheme.FontMD * 0.65;

        var top = new TextBlock
        {
            Text = ch, FontSize = sz, FontFamily = Font,
            Foreground = color,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = isUp ? new Thickness(0, 0, 0, -5) : new Thickness(0),
        };
        var bottom = new TextBlock
        {
            Text = ch, FontSize = sz, FontFamily = Font,
            Foreground = color,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = isUp ? new Thickness(0) : new Thickness(0, -5, 0, 0),
        };

        return new StackPanel
        {
            Orientation = Orientation.Vertical,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 1, 0),
            Children = { top, bottom },
        };
    }

    /// <summary>3단계 색상: 상승(노랑→주황→빨강), 하락(하늘→녹색→파랑)</summary>
    private static IBrush GetTrendBrush(double pctChange, double t1, double t2, double t3)
    {
        double absPct = Math.Abs(pctChange);
        if (pctChange > 0 && absPct >= t1)
        {
            if (absPct >= t3) return new SolidColorBrush(Color.Parse("#ff4444"));  // 빨강
            if (absPct >= t2) return new SolidColorBrush(Color.Parse("#ff8844"));  // 주황
            return new SolidColorBrush(Color.Parse("#ddcc44"));  // 노랑
        }
        if (pctChange < 0 && absPct >= t1)
        {
            if (absPct >= t3) return new SolidColorBrush(Color.Parse("#4466ff"));  // 파랑
            if (absPct >= t2) return new SolidColorBrush(Color.Parse("#44bb88"));  // 녹색
            return new SolidColorBrush(Color.Parse("#66ccee"));  // 하늘
        }
        return AppTheme.FgPrimary;
    }

    // ── 유틸 ─────────────────────────────────────────────────────────────
    private static string FormatValue(string raw, string item)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        if (!double.TryParse(raw, out var v)) return raw;
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

    /// <summary>날짜 노드에 구분별 건수 표시: "01/19 (월) 여수10 율촌3"</summary>
    private string FormatDayWithCounts(string d)
    {
        string baseText = FormatDay(d);

        if (_dateGroupCounts == null || !_dateGroupCounts.TryGetValue(d, out var counts))
            return baseText;

        var parts = new List<string>();
        foreach (var g in new[] { "여수", "율촌", "세풍" })
        {
            if (counts.TryGetValue(g, out int cnt) && cnt > 0)
                parts.Add($"{g}{cnt}");
        }

        return parts.Count > 0 ? $"{baseText}  {string.Join(" ", parts)}" : baseText;
    }
}
