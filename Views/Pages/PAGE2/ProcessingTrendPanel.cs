using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE2;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 처리시설 추이 패널 (Show2): 시설 선택 시 30/60/90/120/150/180일 범위의
/// 시료명(공정)별 항목 추이를 세로 스택으로 표시.
/// · 상단 체크박스로 표시 항목 필터
/// · 시료 순서는 분석계획 시료순서 기준 (FacilityResultService.GetRowsInRange)
/// · 각 시료 카드는 라인/영역/바 세 차트를 표시
/// </summary>
public class ProcessingTrendPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static readonly int[] Periods = { 30, 60, 90, 120, 150, 180 };
    private static readonly Color[] Palette =
    {
        Color.Parse("#ff7755"), Color.Parse("#5588ff"), Color.Parse("#ffcc33"),
        Color.Parse("#55dd55"), Color.Parse("#dd55dd"), Color.Parse("#55ddcc"),
        Color.Parse("#ff9944"), Color.Parse("#aa99ff"), Color.Parse("#99cc44"),
        Color.Parse("#ff6699"), Color.Parse("#66ccff"),
    };

    private readonly TextBlock _tbTitle;
    private readonly WrapPanel _topBar;
    private readonly WrapPanel _itemPanel;
    private readonly StackPanel _chartsHost;
    private Button? _btnLine, _btnArea, _btnBar;

    private string? _facility;
    private int _periodDays = 30;
    private TrendChart.Kind _chartKind = TrendChart.Kind.Line;
    private bool _layoutSeparate; // false=겹침, true=별도
    private Dictionary<string, List<(string Date, Dictionary<string, string> Values)>>? _cachedData;
    private DateTime _cachedFrom, _cachedTo;
    private List<string> _allKeys = new();
    private readonly HashSet<string> _selectedItems = new();
    private readonly Dictionary<string, Color> _colorMap = new();
    private bool _firstLoad = true;

    public ProcessingTrendPanel()
    {
        _tbTitle = new TextBlock
        {
            Text         = "처리시설을 선택하세요",
            FontFamily   = Font,
            FontSize     = AppTheme.FontLG,
            FontWeight   = FontWeight.SemiBold,
            Foreground   = AppTheme.FgMuted,
            Margin       = new Thickness(10, 8, 10, 4),
        };

        _topBar = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            Margin      = new Thickness(10, 2, 10, 4),
            VerticalAlignment = VerticalAlignment.Center,
        };

        _btnLine = MakeKindButton("Line", TrendChart.Kind.Line);
        _btnArea = MakeKindButton("Area", TrendChart.Kind.Area);
        _btnBar  = MakeKindButton("Bar",  TrendChart.Kind.Bar);
        _topBar.Children.Add(_btnLine);
        _topBar.Children.Add(_btnArea);
        _topBar.Children.Add(_btnBar);
        UpdateKindButtonStyle();

        var periodCombo = new ComboBox
        {
            FontSize   = AppTheme.FontSM,
            FontFamily = Font,
            Padding    = new Thickness(6, 2),
            Margin     = new Thickness(12, 0, 0, 0),
            MinWidth   = 70,
            ItemsSource  = Periods,
            SelectedItem = _periodDays,
            VerticalAlignment = VerticalAlignment.Center,
        };
        periodCombo.SelectionChanged += (_, _) =>
        {
            if (periodCombo.SelectedItem is int v)
            {
                _periodDays = v;
                if (_facility != null)
                {
                    _tbTitle.Text = $"{_facility} · 최근 {_periodDays}일 공정별 추이";
                    LoadAndRender();
                }
            }
        };
        _topBar.Children.Add(new TextBlock
        {
            Text = "기간",
            FontFamily = Font,
            FontSize   = AppTheme.FontSM,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 4, 0),
        });
        _topBar.Children.Add(periodCombo);

        // 겹침/별도 토글스위치
        var layoutToggle = new ToggleSwitch
        {
            OffContent = "겹침",
            OnContent  = "별도",
            IsChecked  = _layoutSeparate,
            FontFamily = Font,
            FontSize   = AppTheme.FontSM,
            Padding    = new Thickness(4, 0),
            Margin     = new Thickness(12, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
        layoutToggle.IsCheckedChanged += (_, _) =>
        {
            _layoutSeparate = layoutToggle.IsChecked == true;
            RebuildCharts();
        };
        _topBar.Children.Add(layoutToggle);

        _itemPanel = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            Margin      = new Thickness(10, 0, 10, 6),
        };

        _chartsHost = new StackPanel { Spacing = 10, Margin = new Thickness(10, 0, 10, 10) };
        var scroll = new ScrollViewer
        {
            Content = _chartsHost,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };

        var root = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,Auto,Auto,*") };
        Grid.SetRow(_tbTitle,   0); root.Children.Add(_tbTitle);
        Grid.SetRow(_topBar,    1); root.Children.Add(_topBar);
        Grid.SetRow(_itemPanel, 2); root.Children.Add(_itemPanel);
        Grid.SetRow(scroll,     3); root.Children.Add(scroll);
        Content = root;
    }

    public void SetFacility(string facility)
    {
        _facility = facility;
        _tbTitle.Text = $"{facility} · 최근 {_periodDays}일 공정별 추이";
        _tbTitle.Foreground = AppTheme.FgPrimary;
        LoadAndRender();
    }

    private Button MakeKindButton(string label, TrendChart.Kind kind)
    {
        var btn = new Button
        {
            Content         = label,
            Tag             = kind,
            FontFamily      = Font,
            FontSize        = AppTheme.FontSM,
            Padding         = new Thickness(10, 3),
            Margin          = new Thickness(0, 0, 4, 0),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
        };
        btn.Click += (_, _) =>
        {
            _chartKind = kind;
            UpdateKindButtonStyle();
            RebuildCharts();
        };
        return btn;
    }

    private void UpdateKindButtonStyle()
    {
        foreach (var b in new[] { _btnLine, _btnArea, _btnBar })
        {
            if (b?.Tag is not TrendChart.Kind k) continue;
            bool on = k == _chartKind;
            b.Background  = on ? AppTheme.BgActiveGreen : AppTheme.BgCard;
            b.Foreground  = on ? AppTheme.FgSuccess     : AppTheme.FgMuted;
            b.BorderBrush = on ? AppTheme.BorderActive  : AppTheme.BorderMuted;
            b.FontWeight  = on ? FontWeight.SemiBold    : FontWeight.Normal;
        }
    }

    private void LoadAndRender()
    {
        if (string.IsNullOrEmpty(_facility)) return;

        var to   = DateTime.Today;
        var from = to.AddDays(-_periodDays + 1);
        _cachedFrom = from; _cachedTo = to;

        try
        {
            _cachedData = FacilityResultService.GetRowsInRange(
                _facility!, from.ToString("yyyy-MM-dd"), to.ToString("yyyy-MM-dd"));
        }
        catch (Exception ex)
        {
            _cachedData = null;
            _itemPanel.Children.Clear();
            _chartsHost.Children.Clear();
            _chartsHost.Children.Add(new TextBlock
            {
                Text       = $"조회 오류: {ex.Message}",
                Foreground = AppTheme.FgDanger,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
            });
            return;
        }

        RebuildItemCheckboxes();
        RebuildCharts();
    }

    private void RebuildItemCheckboxes()
    {
        _itemPanel.Children.Clear();
        if (_cachedData == null || _cachedData.Count == 0) return;

        // 분석항목 테이블 순서 기준으로 키 정렬 (GetAnalysisItems 는 순서 컬럼 ORDER BY)
        var itemOrder = FacilityResultService.GetAnalysisItems(activeOnly: false)
            .Select(i => i.컬럼명.Trim('`'))
            .ToList();
        var dataKeys = _cachedData.Values
            .SelectMany(rows => rows.SelectMany(r => r.Values.Keys))
            .Distinct()
            .ToHashSet();

        _allKeys = itemOrder.Where(k => dataKeys.Contains(k)).ToList();
        foreach (var k in dataKeys)
            if (!_allKeys.Contains(k)) _allKeys.Add(k);  // fallback: 미등록 항목 뒤에 붙임

        for (int i = 0; i < _allKeys.Count; i++)
            _colorMap[_allKeys[i]] = Palette[i % Palette.Length];

        if (_firstLoad)
        {
            _selectedItems.Clear();
            foreach (var k in _allKeys) _selectedItems.Add(k);
            _firstLoad = false;
        }
        else
        {
            _selectedItems.RemoveWhere(k => !_allKeys.Contains(k));
        }

        var lbl = new TextBlock
        {
            Text       = "항목:",
            FontFamily = Font,
            FontSize   = AppTheme.FontSM,
            Foreground = AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
            Margin     = new Thickness(0, 0, 8, 0),
        };
        _itemPanel.Children.Add(lbl);

        foreach (var k in _allKeys)
        {
            var color = _colorMap[k];
            var dot = new Ellipse
            {
                Width  = 8, Height = 8,
                Fill   = new SolidColorBrush(color),
                Margin = new Thickness(0, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center,
            };
            var cb = new CheckBox
            {
                Content   = k,
                IsChecked = _selectedItems.Contains(k),
                FontFamily = Font,
                FontSize   = AppTheme.FontSM,
                Foreground = AppTheme.FgSecondary,
                Margin     = new Thickness(0, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center,
            };
            string key = k;
            cb.Checked   += (_, _) => { _selectedItems.Add(key);    RebuildCharts(); };
            cb.Unchecked += (_, _) => { _selectedItems.Remove(key); RebuildCharts(); };

            var wrap = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin      = new Thickness(0, 0, 8, 0),
            };
            wrap.Children.Add(dot);
            wrap.Children.Add(cb);
            _itemPanel.Children.Add(wrap);
        }

        var selectAllBtn = MakeShortcutButton("전체");
        var clearBtn     = MakeShortcutButton("해제");
        selectAllBtn.Click += (_, _) =>
        {
            _selectedItems.Clear();
            foreach (var k in _allKeys) _selectedItems.Add(k);
            RebuildItemCheckboxes();
            RebuildCharts();
        };
        clearBtn.Click += (_, _) =>
        {
            _selectedItems.Clear();
            RebuildItemCheckboxes();
            RebuildCharts();
        };
        _itemPanel.Children.Add(selectAllBtn);
        _itemPanel.Children.Add(clearBtn);
    }

    private static Button MakeShortcutButton(string text) => new()
    {
        Content         = text,
        Padding         = new Thickness(8, 2),
        FontFamily      = Font,
        FontSize        = AppTheme.FontSM,
        Background      = AppTheme.BgCard,
        Foreground      = AppTheme.FgMuted,
        BorderBrush     = AppTheme.BorderMuted,
        BorderThickness = new Thickness(1),
        CornerRadius    = new CornerRadius(4),
        Margin          = new Thickness(4, 0, 0, 0),
    };

    private void RebuildCharts()
    {
        _chartsHost.Children.Clear();
        if (_cachedData == null || _cachedData.Count == 0)
        {
            _chartsHost.Children.Add(new TextBlock
            {
                Text       = $"{_cachedFrom:yyyy-MM-dd} ~ {_cachedTo:yyyy-MM-dd} 기간 데이터가 없습니다.",
                Foreground = AppTheme.FgMuted,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
                Margin     = new Thickness(0, 20, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            return;
        }
        if (_selectedItems.Count == 0)
        {
            _chartsHost.Children.Add(new TextBlock
            {
                Text       = "표시할 항목을 선택하세요.",
                Foreground = AppTheme.FgMuted,
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
                Margin     = new Thickness(0, 20, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            return;
        }

        // 별도 모드: 항목 순서대로 한 ROW에 컬럼으로 나눠 표시. 선택된 _allKeys 순서 유지.
        var orderedSelected = _allKeys.Where(k => _selectedItems.Contains(k)).ToList();

        foreach (var (sample, rows) in _cachedData)
            _chartsHost.Children.Add(BuildSampleCard(
                sample, rows, _cachedFrom, _cachedTo,
                _selectedItems, _colorMap, _chartKind,
                _layoutSeparate, orderedSelected));
    }

    private static Border BuildSampleCard(
        string sample,
        List<(string Date, Dictionary<string, string> Values)> rows,
        DateTime from, DateTime to,
        HashSet<string> selected,
        Dictionary<string, Color> colorMap,
        TrendChart.Kind kind,
        bool separate,
        List<string> orderedSelected)
    {
        var header = new TextBlock
        {
            Text       = sample,
            FontFamily = Font,
            FontSize   = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
            Margin     = new Thickness(0, 0, 0, 4),
        };

        Control chartContent;
        if (separate && orderedSelected.Count > 0)
        {
            var grid = new Grid { Height = 180, ColumnSpacing = 6 };
            for (int i = 0; i < orderedSelected.Count; i++)
                grid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

            for (int i = 0; i < orderedSelected.Count; i++)
            {
                var k = orderedSelected[i];
                var oneSet = new HashSet<string> { k };

                var colHeader = new TextBlock
                {
                    Text       = k,
                    FontFamily = Font,
                    FontSize   = AppTheme.FontSM,
                    Foreground = new SolidColorBrush(colorMap[k]),
                    Margin     = new Thickness(4, 0, 0, 2),
                    HorizontalAlignment = HorizontalAlignment.Left,
                };
                var one = new TrendChart(rows, from, to, oneSet, colorMap, kind);
                var cell = new Grid { RowDefinitions = RowDefinitions.Parse("Auto,*") };
                Grid.SetRow(colHeader, 0); cell.Children.Add(colHeader);
                Grid.SetRow(one,       1); cell.Children.Add(one);
                Grid.SetColumn(cell, i);
                grid.Children.Add(cell);
            }
            chartContent = grid;
        }
        else
        {
            chartContent = new TrendChart(rows, from, to, selected, colorMap, kind) { Height = 180 };
        }

        var stack = new StackPanel { Spacing = 2 };
        stack.Children.Add(header);
        stack.Children.Add(chartContent);

        return new Border
        {
            Background      = AppTheme.BgCard,
            BorderBrush     = AppTheme.BorderMuted,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(10, 8),
            Child           = stack,
        };
    }

    // ─────────────────────────────────────────────────────────────────────
    // TrendChart: 라인 / 영역 / 바 (항목별 정규화, x=날짜)
    // ─────────────────────────────────────────────────────────────────────
    private sealed class TrendChart : Control
    {
        public enum Kind { Line, Area, Bar }

        private static readonly Color ExceedColor = Color.Parse("#ff3b30");

        private readonly List<(DateTime D, Dictionary<string, double> V)> _points;
        private readonly DateTime _from, _to;
        private readonly List<string> _keys;
        private readonly Dictionary<string, Color> _colorMap;
        private readonly Kind _kind;
        private readonly Dictionary<string, double> _normMax; // 항목별 정규화 최대값 = 평균*3
        private DateTime? _hoverDate;

        public TrendChart(
            List<(string Date, Dictionary<string, string> Values)> rows,
            DateTime from, DateTime to,
            HashSet<string> selected,
            Dictionary<string, Color> colorMap,
            Kind kind)
        {
            _from = from; _to = to; _colorMap = colorMap; _kind = kind;
            _points = new();
            foreach (var (d, v) in rows)
            {
                if (!DateTime.TryParse(d, out var dt)) continue;
                var numv = new Dictionary<string, double>();
                foreach (var (k, s) in v)
                {
                    if (!selected.Contains(k)) continue;
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                        numv[k] = n;
                }
                if (numv.Count > 0) _points.Add((dt, numv));
            }
            _points.Sort((a, b) => a.D.CompareTo(b.D));

            // 항목 순서: colorMap 순서(= 분석항목 테이블 순서) 유지
            var present = _points.SelectMany(p => p.V.Keys).ToHashSet();
            _keys = colorMap.Keys.Where(k => present.Contains(k)).ToList();

            _normMax = new Dictionary<string, double>();
            foreach (var k in _keys)
            {
                var ys = _points.Where(p => p.V.ContainsKey(k)).Select(p => p.V[k]).ToList();
                if (ys.Count == 0) { _normMax[k] = 1; continue; }
                double mean = ys.Average();
                double nm   = mean * 3;
                if (nm < 1e-9) nm = ys.Max(y => Math.Abs(y));
                if (nm < 1e-9) nm = 1;
                _normMax[k] = nm;
            }

            PointerMoved += (_, e) => OnHoverMove(e);
            PointerExited += (_, _) =>
            {
                if (_hoverDate != null) { _hoverDate = null; InvalidateVisual(); }
            };
        }

        private void OnHoverMove(PointerEventArgs e)
        {
            if (_points.Count == 0) return;
            var pos = e.GetPosition(this);
            double padL = 6, padR = 10;
            double plotW = Bounds.Width - padL - padR;
            if (plotW <= 0) return;
            double totalDays = Math.Max(1, (_to - _from).TotalDays);

            DateTime? best = null;
            double bestDist = double.MaxValue;
            foreach (var (dt, _) in _points)
            {
                double x = padL + (dt - _from).TotalDays / totalDays * plotW;
                double d = Math.Abs(x - pos.X);
                if (d < bestDist) { bestDist = d; best = dt; }
            }
            var newHover = bestDist <= 25 ? best : null;
            if (newHover != _hoverDate) { _hoverDate = newHover; InvalidateVisual(); }
        }

        public override void Render(DrawingContext ctx)
        {
            base.Render(ctx);
            var w = Bounds.Width; var h = Bounds.Height;
            if (w < 20 || h < 20) return;

            // 포인터 이벤트가 전체 영역에서 동작하도록 투명 배경을 깔아준다.
            ctx.FillRectangle(Brushes.Transparent, new Rect(0, 0, w, h));

            double padL = 6, padR = 10, padT = 8, padB = 20;
            double plotW = w - padL - padR;
            double plotH = h - padT - padB;
            if (plotW <= 0 || plotH <= 0) return;

            var axisPen = new Pen(AppTheme.BorderMuted, 1);
            ctx.DrawLine(axisPen, new Point(padL, padT + plotH), new Point(padL + plotW, padT + plotH));

            double totalDays = Math.Max(1, (_to - _from).TotalDays);
            var face = new Typeface(Font);
            foreach (var (dt, label) in new[] {
                (_from, _from.ToString("MM-dd")),
                (_from.AddDays(totalDays / 2), _from.AddDays(totalDays / 2).ToString("MM-dd")),
                (_to, _to.ToString("MM-dd")),
            })
            {
                double x = padL + (dt - _from).TotalDays / totalDays * plotW;
                var ft = new FormattedText(label, CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight, face, 9, AppTheme.FgMuted);
                ctx.DrawText(ft, new Point(x - ft.Width / 2, padT + plotH + 3));
            }

            if (_points.Count == 0 || _keys.Count == 0) return;

            switch (_kind)
            {
                case Kind.Line: RenderLine(ctx, padL, padT, plotW, plotH, totalDays); break;
                case Kind.Area: RenderArea(ctx, padL, padT, plotW, plotH, totalDays); break;
                case Kind.Bar:  RenderBar (ctx, padL, padT, plotW, plotH, totalDays); break;
            }

            if (_hoverDate != null)
                DrawTooltip(ctx, padL, padT, plotW, plotH, totalDays);
        }

        private double YNorm(string key, double y, double padT, double plotH)
        {
            double nm = _normMax[key];
            double ratio = Math.Min(1.0, y / nm);
            return padT + plotH - ratio * plotH;
        }

        private void RenderLine(DrawingContext ctx, double padL, double padT, double plotW, double plotH, double totalDays)
        {
            foreach (var k in _keys)
            {
                var color = _colorMap[k];
                var normalPen = new Pen(new SolidColorBrush(color), 1.5);
                var exceedPen = new Pen(new SolidColorBrush(ExceedColor), 1.8);
                var normalBrush = new SolidColorBrush(color);
                var exceedBrush = new SolidColorBrush(ExceedColor);
                double nm = _normMax[k];

                var series = _points.Where(p => p.V.ContainsKey(k))
                    .Select(p => (p.D, Y: p.V[k])).ToList();
                if (series.Count == 0) continue;

                Point? prev = null;
                bool prevExceed = false;
                foreach (var (dt, y) in series)
                {
                    double x = padL + (dt - _from).TotalDays / totalDays * plotW;
                    double yn = YNorm(k, y, padT, plotH);
                    bool exceed = y > nm;
                    var p = new Point(x, yn);
                    if (prev != null)
                    {
                        var segPen = (exceed || prevExceed) ? exceedPen : normalPen;
                        ctx.DrawLine(segPen, prev.Value, p);
                    }
                    ctx.DrawEllipse(exceed ? exceedBrush : normalBrush, null, p, exceed ? 3.0 : 2.2, exceed ? 3.0 : 2.2);
                    prev = p; prevExceed = exceed;
                }
            }
        }

        private void RenderArea(DrawingContext ctx, double padL, double padT, double plotW, double plotH, double totalDays)
        {
            foreach (var k in _keys)
            {
                var color = _colorMap[k];
                var fill  = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B));
                var pen   = new Pen(new SolidColorBrush(color), 1.2);
                double nm = _normMax[k];

                var series = _points.Where(p => p.V.ContainsKey(k))
                    .Select(p => (p.D, Y: p.V[k])).ToList();
                if (series.Count == 0) continue;

                var geom = new StreamGeometry();
                using (var g = geom.Open())
                {
                    var first = series[0];
                    double x0 = padL + (first.D - _from).TotalDays / totalDays * plotW;
                    double baseY = padT + plotH;
                    g.BeginFigure(new Point(x0, baseY), isFilled: true);

                    double lastX = x0;
                    foreach (var (dt, y) in series)
                    {
                        double x = padL + (dt - _from).TotalDays / totalDays * plotW;
                        double yn = YNorm(k, y, padT, plotH);
                        g.LineTo(new Point(x, yn));
                        lastX = x;
                    }
                    g.LineTo(new Point(lastX, baseY));
                    g.EndFigure(isClosed: true);
                }
                ctx.DrawGeometry(fill, pen, geom);

                // 초과 포인트만 빨간 점으로 오버레이
                var exceedBrush = new SolidColorBrush(ExceedColor);
                foreach (var (dt, y) in series)
                {
                    if (y <= nm) continue;
                    double x = padL + (dt - _from).TotalDays / totalDays * plotW;
                    double yn = YNorm(k, y, padT, plotH);
                    ctx.DrawEllipse(exceedBrush, null, new Point(x, yn), 3.0, 3.0);
                }
            }
        }

        private void RenderBar(DrawingContext ctx, double padL, double padT, double plotW, double plotH, double totalDays)
        {
            int n = _keys.Count;
            if (n == 0) return;

            double slotW  = plotW / Math.Max(1, totalDays + 1);
            double groupW = Math.Max(1.0, slotW * 0.9);
            double barW   = Math.Max(0.5, groupW / n);
            var exceedBrush = new SolidColorBrush(ExceedColor);

            foreach (var (dt, vmap) in _points)
            {
                double xCenter = padL + (dt - _from).TotalDays / totalDays * plotW;
                for (int i = 0; i < n; i++)
                {
                    var k = _keys[i];
                    if (!vmap.TryGetValue(k, out var y)) continue;
                    double nm = _normMax[k];
                    bool exceed = y > nm;
                    double ratio = Math.Min(1.0, y / nm);
                    double barH = ratio * plotH;
                    double xLeft = xCenter - groupW / 2 + i * barW;
                    double yTop  = padT + plotH - barH;
                    var brush = exceed ? exceedBrush : new SolidColorBrush(_colorMap[k]);
                    ctx.DrawRectangle(brush, null,
                        new Rect(xLeft, yTop, barW, Math.Max(0.5, barH)));
                }
            }
        }

        private void DrawTooltip(DrawingContext ctx, double padL, double padT, double plotW, double plotH, double totalDays)
        {
            var dt = _hoverDate!.Value;
            double xLine = padL + (dt - _from).TotalDays / totalDays * plotW;

            var dashPen = new Pen(AppTheme.BorderActive, 1) { DashStyle = DashStyle.Dash };
            ctx.DrawLine(dashPen, new Point(xLine, padT), new Point(xLine, padT + plotH));

            int idx = _points.FindIndex(p => p.D == dt);
            if (idx < 0) return;
            var vmap = _points[idx].V;

            var face = new Typeface(Font);
            var headerFt = new FormattedText(dt.ToString("yyyy-MM-dd"),
                CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                face, 11, AppTheme.FgPrimary);

            var rows = new List<(FormattedText Ft, Color Dot, bool Exceed)>();
            foreach (var k in _keys)
            {
                if (!vmap.TryGetValue(k, out var y)) continue;
                bool exceed = y > _normMax[k];
                var fg = exceed ? (IBrush)new SolidColorBrush(ExceedColor) : AppTheme.FgSecondary;
                var ft = new FormattedText($"{k}: {FormatVal(y)}",
                    CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                    face, 10, fg);
                rows.Add((ft, _colorMap[k], exceed));
            }

            double rowsW = rows.Count > 0 ? rows.Max(r => r.Ft.Width) : 0;
            double maxW  = Math.Max(headerFt.Width, rowsW);
            double rowH  = rows.Count > 0 ? rows[0].Ft.Height + 2 : 0;
            double boxW  = maxW + 28;
            double boxH  = headerFt.Height + 6 + rows.Count * rowH + 8;

            double tx = xLine + 10;
            if (tx + boxW > padL + plotW) tx = xLine - boxW - 10;
            if (tx < padL) tx = padL;
            double ty = padT + 5;
            if (ty + boxH > padT + plotH) ty = padT + plotH - boxH - 2;

            var bg = new SolidColorBrush(Color.FromArgb(235, 20, 24, 30));
            ctx.DrawRectangle(bg, new Pen(AppTheme.BorderActive, 1),
                new Rect(tx, ty, boxW, boxH), 4, 4);

            double y0 = ty + 4;
            ctx.DrawText(headerFt, new Point(tx + 8, y0));
            y0 += headerFt.Height + 4;
            foreach (var (ft, dot, _) in rows)
            {
                ctx.DrawEllipse(new SolidColorBrush(dot), null,
                    new Point(tx + 12, y0 + ft.Height / 2), 3, 3);
                ctx.DrawText(ft, new Point(tx + 20, y0));
                y0 += rowH;
            }
        }

        private static string FormatVal(double v)
        {
            if (Math.Abs(v) >= 100) return v.ToString("0.#");
            if (Math.Abs(v) >= 10)  return v.ToString("0.##");
            return v.ToString("0.###");
        }
    }
}
